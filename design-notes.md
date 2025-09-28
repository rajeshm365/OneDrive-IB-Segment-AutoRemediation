# Design Notes – OneDrive IB Segment Auto-Remediation

## Overview
- Detect OneDrive sites with missing IB allowed segments and remediate automatically.

## Flow
1. Initialize variables, logging
2. Connect to SPO via `Connect-PnPOnline` (app reg)
3. Load IB users from CSV (Azure AD groups)
4. Get all OneDrives; filter to IB users
5. For each site:
   - Get current IB segments (REST ProcessQuery)
   - If count < expected (3), branch:
     - If 1: identify stamped segment → compute allowed set (Private/Public) → add missing
     - If 0 or legacy duplicates: log & skip (re-join scenario)
6. Write logs; upload to SharePoint
7. Power Automate: if `fixes*.txt`, post to Teams

## Permissions (App Registration)
- SharePoint: `Sites.ReadWrite.All`
- Graph (if using group enumeration): `Group.Read.All` (app)
- PnP.PowerShell used for SPO + Graph operations

## Operational notes
- Runtime ~2 days for large tenants; schedule weekly
- Logs are separated by purpose for easier triage
- Replace dummy segment GUIDs in script before use
