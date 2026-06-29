# M365 Data Migration Scripts

Professional, enterprise-grade PowerShell scripts for Microsoft 365 migration and administration projects.

## Repository Purpose

This repository contains battle-tested scripts used in real-world M365 data migration engagements. All scripts follow consistent standards for logging, error handling, progress reporting, WhatIf support, and CSV output for downstream migration tools.

## Available Scripts

### Pre-Provision OneDrive for Azure AD Security Group (E5 Licensing) — v1.2

**Script:** `PreProvision-OneDrive-E5-Group.ps1`

**Purpose:**
Queries the specified Azure AD security group (default ID linked to M365 E5 licensing), pre-provisions OneDrive personal sites for all user members using batched `Request-SPOPersonalSite` calls, and exports a detailed CSV report.

**Key Fix in v1.2:**
Removed all `Import-Module` calls for Microsoft.Graph sub-modules. The cmdlets now load on-demand when `Connect-MgGraph` is executed. This completely resolves the recurring `function capacity 4096 exceeded` / `SessionStateOverflowException` errors that many users hit with both the full `Microsoft.Graph` package and even selective sub-module imports.

**Quick Start:**

```powershell
iwr "https://raw.githubusercontent.com/jessemooreuk/M365-Data-Migration-Scripts/main/PreProvision-OneDrive-E5-Group.ps1" -OutFile PreProvision-OneDrive-E5-Group.ps1

.\PreProvision-OneDrive-E5-Group.ps1 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com"
```

**Strongly Recommended:** Run this script in **PowerShell 7.x** (not Windows PowerShell 5.1) for best compatibility with large Microsoft Graph modules.

**Parameters:**
- `-TenantAdminUrl` (mandatory)
- `-GroupObjectId` (optional)
- `-WhatIf`
- `-BatchSize`
- `-OutputCsvPath`

## Prerequisites
- PowerShell 7.x (preferred) or 5.1
- Microsoft Graph PowerShell SDK (installed on-demand)
- SharePoint Online Management Shell
- SharePoint Administrator or Global Admin rights

---
*Maintained by the M365 Data Migration Team*