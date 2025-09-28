<# 
.SYNOPSIS
  Detects and fixes OneDrive Information Barrier (IB) segment stamping issues.
.DESCRIPTION
  Weekly job. Scans all OneDrives owned by IB users. If a site has fewer than the expected
  number of segment stamps (typical issue: only 1 stamped), determines whether the user’s
  segment is in Public or Private class, computes the allowed segments, and stamps the missing
  ones via SharePoint REST (ProcessQuery). Writes logs to disk and uploads to SharePoint.
  If a fixes log is produced (filename starts with "fixes" and ends with ".txt"), your Power Automate
  flow will post an alert to a Teams channel.
.NOTES
  Auth: App Registration (certificate or secret) + Connect-PnPOnline. Uses Graph/SPO perms.
  Placeholders below must be filled for your tenant.
#>

# =========================
# Parameters / Constants
# =========================

param(
  [Parameter(Mandatory=$true)][string]$TenantName,             # e.g. contoso.onmicrosoft.com
  [Parameter(Mandatory=$true)][string]$TenantId,               # e.g. 00000000-0000-0000-0000-000000000000
  [Parameter(Mandatory=$true)][string]$ClientId,               # App registration (Enterprise App) ID
  [Parameter(Mandatory=$true)][string]$AuthMode,               # "Certificate" or "Secret"
  [string]$CertificatePath = "C:\certs\ib-app.pfx",
  [string]$CertificatePassword = "<PFX-PASSWORD>",             # or use a secure asset in Automation
  [string]$ClientSecret = "<APP-CLIENT-SECRET>",               # only if AuthMode = "Secret"

  # SharePoint URLs and paths
  [Parameter(Mandatory=$true)][string]$SPOAdminUrl,            # e.g. https://contoso-admin.sharepoint.com
  [Parameter(Mandatory=$true)][string]$SPODomain,              # e.g. contoso (used to compose tenant URLs)
  [string]$OneDriveUrlPrefix = "-my.sharepoint.com/personal/", # default prefix pattern
  [string]$SharePointLogsSiteRelativeUrl = "/sites/M365Management",
  [string]$SharePointLogsLibrary = "MIBOneDriveSegmentCheck",  # doc library to upload logs

  # Inputs
  [string]$UserGroupsCsv = "C:\Input\IB_Groups.csv",           # CSV listing Azure AD group IDs to scan
  [string]$AllOneDrivesCache = "C:\Temp\AllOneDrives.json",    # optional cache (speed-up)

  # Local log path
  [string]$LogsBasePath = "C:\Logs\MIB_FixOneDriveSegments"
)

# Expected segment count (your environment stamps 3: user's own + 2 allowed)
[int]$ExpectedSegmentCount = 3

# Arrays of known GUIDs (Dummy placeholders — replace with real GUIDs in your secured env)
$PrivateSegmentIds = @(
  "00000000-0000-0000-0000-0000000000A1",
  "00000000-0000-0000-0000-0000000000A2",
  "00000000-0000-0000-0000-0000000000A3"
)
$PublicSegmentIds = @(
  "00000000-0000-0000-0000-0000000000B1",
  "00000000-0000-0000-0000-0000000000B2",
  "00000000-0000-0000-0000-0000000000B3",
  "00000000-0000-0000-0000-0000000000B4"
)

# Date stamp and log files (Power Automate trigger: file starts with "fixes" and ends with ".txt")
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ActionLog          = Join-Path $LogsBasePath "action_$Stamp.txt"
$ErrorLog           = Join-Path $LogsBasePath "errors_$Stamp.txt"
$LessThan3Log       = Join-Path $LogsBasePath "lessThan3_$Stamp.txt"
$FixesLog           = Join-Path $LogsBasePath "fixes_$Stamp.txt"   # <-- Flow watches: startswith(fixes) && endswith(.txt)

# Ensure TLS
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# Create log folder
New-Item -Path $LogsBasePath -ItemType Directory -Force | Out-Null

# =========================
# Logging helpers
# =========================
function Write-Log([string]$Message)     { ("{0}  {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $Message) | Tee-Object -FilePath $ActionLog -Append }
function Write-ErrorLog([string]$Message){ ("{0}  {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $Message) | Tee-Object -FilePath $ErrorLog  -Append }
function Write-LessThan3([string]$Msg)   { ("{0}  {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $Msg)     | Tee-Object -FilePath $LessThan3Log -Append }
function Write-FixLog([string]$Msg)      { ("{0}  {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $Msg)     | Tee-Object -FilePath $FixesLog    -Append }

# =========================
# XML payload templates for ProcessQuery (SPO)
# =========================
# These are minimal placeholders; in your environment you likely have full ProcessQuery envelopes.
# Replace [[SITE_URL]] and [[IB_SEGMENTS]] placeholders at runtime.
$GetIbSegmentsXml = @"
<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="PnP.PowerShell">
  <!-- GET IB segment info for [[SITE_URL]] -->
</Request>
"@

$SetIbSegmentsXml = @"
<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="PnP.PowerShell">
  <!-- SET IB segments [[IB_SEGMENTS]] for [[SITE_URL]] -->
</Request>
"@

# =========================
# Utility: Convert SPO date string to readable format (optional)
# =========================
function Convert-CustomDate([string]$DateString) {
  if (-not $DateString) { return "" }
  if ($DateString -match "\/Date\((\d+)" ) {
    $ms = [int64]$matches[1]
    return (Get-Date ([DateTimeOffset]::FromUnixTimeMilliseconds($ms).DateTime) -f "MM/dd/yyyy HH:mm:ss tt")
  }
  return $DateString
}

# =========================
# Functions: Data acquisition
# =========================

function Connect-SharePoint {
  Write-Log "Connecting to SharePoint Online ($SPOAdminUrl) as app…"
  if ($AuthMode -ieq "Certificate") {
    $secure = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
    Connect-PnPOnline -Url $SPOAdminUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword $secure -WarningAction SilentlyContinue
  } else {
    Connect-PnPOnline -Url $SPOAdminUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -WarningAction SilentlyContinue
  }
}

function Get-AllOneDrives {
  if (Test-Path $AllOneDrivesCache) {
    Write-Log "Loading cached OneDrives from $AllOneDrivesCache"
    return Get-Content $AllOneDrivesCache | ConvertFrom-Json
  }
  Write-Log "Querying tenant for all OneDrive sites…"
  $sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '$($SPODomain)$OneDriveUrlPrefix'"
  $sites | ConvertTo-Json -Depth 4 | Out-File $AllOneDrivesCache -Encoding UTF8
  return $sites
}

# CSV format option 1: list of Azure AD group IDs (col: GroupId)
# We then enumerate members via PnP Graph cmdlets
function Get-IBUsersFromGroups([string]$CsvPath){
  if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }
  $groups = Import-Csv $CsvPath
  $upns = @()
  foreach($g in $groups){
    try {
      $members = Get-PnPMicrosoft365GroupMember -Identity $g.GroupId
      $upns += $members.UserPrincipalName
    } catch {
      Write-ErrorLog "Failed to read members for GroupId=$($g.GroupId) :: $_"
    }
  }
  $upns | Sort-Object -Unique
}

# Filter OneDrive sites owned by IB users
function Filter-OneDrivesByIBUsers($AllSites, $UserUPNs){
  $hash = @{}
  foreach($u in $UserUPNs){ $hash[$u.ToLower()] = $true }
  $AllSites | Where-Object { 
    $_.Owner -and $hash.ContainsKey( ($_.Owner).ToLower() )
  }
}

# =========================
# Functions: IB segment info + fix
# =========================

function Get-IBSegmentInfo([string]$SiteUrl){
  Write-Log "Get IB segments for $SiteUrl"
  try {
    $resp = Invoke-PnPSPRestMethod -Url "/_vti_bin/client.svc/ProcessQuery" `
      -Method Post -ContentType "text/xml" -Raw `
      -Content ($GetIbSegmentsXml.Replace("[[SITE_URL]]", $SiteUrl))

    # TODO: parse $resp to extract segments + creation date
    # For placeholder, return a hashtable shape:
    return @{
      IBSegment   = @("00000000-0000-0000-0000-0000000000XX")  # replace with parsed list
      ODCreation  = (Get-Date)
    }
  } catch {
    Write-ErrorLog "Failed to query IB segments for $SiteUrl :: $_"
    return $null
  }
}

function Fix-MissingSegments($Site, $IBSegment, $ODCreation){
  # Combine known classes
  $allKnown = $PrivateSegmentIds + $PublicSegmentIds

  # Determine stamped segment and its class
  $stamped = $IBSegment | Where-Object { $_ -ne "" } | Select-Object -First 1
  if (-not $stamped) { Write-Log "No stamped segment found for $($Site.Url). Skipping."; return $false }

  $expected = @()
  if ($PrivateSegmentIds -contains $stamped)      { $expected = $PrivateSegmentIds }
  elseif ($PublicSegmentIds -contains $stamped)   { $expected = $PublicSegmentIds }
  else {
    Write-Log     "Stamped segment not recognized for $($Site.Url). Skipping."
    Write-ErrorLog "Stamped segment not recognized for $($Site.Url). Value=$stamped"
    return $false
  }

  # Work out missing IDs
  $missing = $expected | Where-Object { $IBSegment -notcontains $_ }
  if (-not $missing -or $missing.Count -eq 0) { Write-Log "No missing segments for $($Site.Url)"; return $false }

  # Add each missing segment via ProcessQuery
  foreach($seg in $missing){
    try{
      $payload = $SetIbSegmentsXml.Replace("[[SITE_URL]]",$Site.Url).Replace("[[IB_SEGMENTS]]",$seg)
      $setResp = Invoke-PnPSPRestMethod -Url "/_vti_bin/client.svc/ProcessQuery" `
                 -Method Post -ContentType "text/xml" -Raw -Content $payload
      Write-Log   "Added missing segment $seg to $($Site.Url)"
      Write-FixLog "Added missing segment $seg to $($Site.Url)"
    } catch {
      Write-ErrorLog "Failed to add segment $seg to $($Site.Url) :: $_"
    }
  }
  return $true
}

# =========================
# Main
# =========================

Write-Log "Program begin"

# 1) Connect to SPO
Connect-SharePoint

# 2) Get OneDrives
$allOD = Get-AllOneDrives
Write-Log "All OneDrives discovered: $($allOD.Count)"

# 3) Get IB user UPNs from groups CSV
$ibUsers = Get-IBUsersFromGroups -CsvPath $UserGroupsCsv
Write-Log "IB users loaded: $($ibUsers.Count)"

# 4) Filter ODs to IB users only
$ibODs = Filter-OneDrivesByIBUsers -AllSites $allOD -UserUPNs $ibUsers
Write-Log "OneDrives owned by IB users: $($ibODs.Count)"

# 5) Iterate and fix
foreach($site in $ibODs){
  $info = Get-IBSegmentInfo -SiteUrl $site.Url
  if ($null -eq $info) {
    Write-ErrorLog "No IB info returned for $($site.Url)"
    continue
  }

  $segments = @($info.IBSegment)
  $count    = $segments.Count
  $odc      = $info.ODCreation

  if ($count -lt $ExpectedSegmentCount) {
    Write-LessThan3 "Less than $ExpectedSegmentCount segments: $count  | $($site.Url) | Created: $(Convert-CustomDate $odc)"

    if ($count -eq 1) {
      $updated = Fix-MissingSegments -Site $site -IBSegment $segments -ODCreation $odc
      if ($updated) {
        Write-Log   ("After fix: {0} | {1}" -f $site.Url, ($(Get-Date)))
        Write-FixLog("After fix: {0} | {1}" -f $site.Url, ($(Get-Date)))
      }
    }
  }
}

# 6) Upload logs to SharePoint (so your flow can trigger)
try {
  # reconnect with same app (PnP requires reconnect when switching target site)
  $targetSiteUrl = "https://$SPODomain.sharepoint.com$SharePointLogsSiteRelativeUrl"
  if ($AuthMode -ieq "Certificate") {
    $secure = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
    Connect-PnPOnline -Url $targetSiteUrl -ClientId $ClientId -Tenant $TenantId -CertificatePath $CertificatePath -CertificatePassword $secure -WarningAction SilentlyContinue
  } else {
    Connect-PnPOnline -Url $targetSiteUrl -ClientId $ClientId -Tenant $TenantId -ClientSecret $ClientSecret -WarningAction SilentlyContinue
  }

  $folderServerRelativeUrl = "$SharePointLogsSiteRelativeUrl/$SharePointLogsLibrary"

  foreach($log in @($ActionLog,$ErrorLog,$LessThan3Log,$FixesLog)){
    if (Test-Path $log) {
      Add-PnPFile -Path $log -Folder $folderServerRelativeUrl -Values @{ CheckInComment = "Weekly MIB scan and fix" } | Out-Null
      Write-Log "Uploaded $(Split-Path $log -Leaf) to $folderServerRelativeUrl"
    }
  }
} catch {
  Write-ErrorLog "SharePoint upload failed :: $_"
}

Write-Log "Program end"
