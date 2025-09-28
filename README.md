# OneDrive IB Segment Auto-Remediation

Weekly PowerShell automation that scans all OneDrives for **Information Barrier (IB) segment** stamping issues and auto-fixes them.  
Root cause: a Microsoft-side issue occasionally stamps only the user's own IB segment, **omitting allowed segments**, which breaks expected collaboration.

## What it does
- Scans OneDrives owned by IB users
- Detects sites with **< 3 segments** (typical error: only 1)
- Decides if the stamped segment is **Private** or **Public**
- Stamps the **missing allowed segments** via SharePoint REST (ProcessQuery)
- Logs **actions**, **errors**, **less-than-3**, and **fixes** to disk and **uploads to SharePoint**
- A **Power Automate** flow watches for files that **start with `fixes` and end with `.txt`** and posts to **Teams**

## Tech
- PowerShell + PnP.PowerShell
- App Registration (Graph/SPO perms)  
- Connect-PnPOnline (cert or secret)
- SharePoint REST ProcessQuery for IB stamping
- Power Automate (files trigger) → Teams

## Parameters to set
Fill these in runbook or script:
- `TenantName`, `TenantId`, `ClientId`
- `AuthMode` = `Certificate` or `Secret`
- `CertificatePath` + `CertificatePassword` **or** `ClientSecret`
- `SPOAdminUrl`, `SPODomain`
- `SharePointLogsSiteRelativeUrl`, `SharePointLogsLibrary`
- `UserGroupsCsv` (CSV with **GroupId** column)

## Expected segment sets
Replace dummy GUIDs in the script with your **Private** and **Public** segment ID arrays.

## Logs produced
- `action_*.txt` – major steps
- `errors_*.txt` – errors
- `lessThan3_*.txt` – sites with < 3 segments
- `fixes_*.txt` – **fixes applied** (triggers your Power Automate flow)

## Flow trigger rule
When a file is added to the SharePoint logs library:
- **Name starts with:** `fixes`
- **Name ends with:** `.txt`
→ Post message in Teams (attach file or include its contents)

## Scheduling
- Run **weekly** (takes ~2 days in a large tenant)

## Safety
- Script is **idempotent** per missing segment
- Write-protected via app permissions
- Full audit trail in SharePoint

## 🏗️ Architecture

```mermaid
flowchart TD
    A[Weekly Schedule - Azure Automation Runbook or Task Scheduler] --> B[PowerShell Script - Hybrid or Azure Worker]
    B --> PNP[Connect-PnPOnline - Cert or Secret]
    B --> API[Graph and SharePoint Online APIs]
    API --> SCAN[Scan all OneDrives owned by IB users]

    subgraph Decision
      SCAN --> D1{Site has &lt; 3 IB segments?}
      D1 -- No --> SKIP[Healthy - skip]
      D1 -- Yes --> D2{Stamped segment is Private or Public?}
      D2 --> PRIV[Determine Private]
      D2 --> PUB[Determine Public]
      PRIV --> STAMP[Stamp missing allowed segments - SharePoint REST ProcessQuery]
      PUB --> STAMP
    end

    STAMP --> LOG_ACTION[Write action star .txt]
    B --> LOG_ERRORS[Write errors star .txt]
    D1 -->|Yes| LOG_LT3[Write lessThan3 star .txt]
    STAMP --> LOG_FIXES[Write fixes star .txt]

    subgraph SharePoint Logs Library
      UP_ACTION[Upload action logs] --> SP[(SharePoint Library)]
      UP_ERRORS[Upload error logs] --> SP
      UP_LT3[Upload lessThan3 logs] --> SP
      UP_FIXES[Upload fixes logs] --> SP
    end

    LOG_ACTION --> UP_ACTION
    LOG_ERRORS --> UP_ERRORS
    LOG_LT3 --> UP_LT3
    LOG_FIXES --> UP_FIXES

    subgraph Power Automate Flow
      TRIG{On file created - name starts with fixes and ends with .txt?}
      TRIG -- Yes --> TEAMS[Post message in Microsoft Teams - attach file or include contents]
      TRIG -- No --> NOOP[Do nothing]
    end

    SP --> TRIG
