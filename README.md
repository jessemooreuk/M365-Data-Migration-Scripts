# M365 Data Migration Scripts

Professional, enterprise-grade PowerShell scripts for Microsoft 365 migration and administration projects.

## Repository Purpose

This repository contains battle-tested scripts used in real-world M365 data migration engagements. All scripts follow consistent standards for logging, error handling, progress reporting, WhatIf support, and CSV output for downstream migration tools.

## Available Scripts

### 1. Pre-Provision OneDrive for Azure AD Security Group (E5 Licensing) v1.1

**Script:** `PreProvision-OneDrive-E5-Group.ps1`

**Purpose:**
Queries the specified Azure AD security group (default ID linked to M365 E5 licensing), pre-provisions OneDrive personal sites for all user members using batched `Request-SPOPersonalSite` calls, and exports a detailed CSV report containing:
- DisplayName
- UPN
- OneDriveUrl (constructed in standard format)
- Pre-provision status
- Timestamp

**Why use this?**
- OneDrive sites must exist before many migration tools can migrate content or permissions.
- Pre-provisioning in advance avoids delays during cutover.
- Ideal for group-based licensed users (E5, E3 + SharePoint, etc.).

**Quick Start (PowerShell):**

```powershell
# Download the latest version directly
iwr "https://raw.githubusercontent.com/jessemooreuk/M365-Data-Migration-Scripts/main/PreProvision-OneDrive-E5-Group.ps1" -OutFile PreProvision-OneDrive-E5-Group.ps1

# Review the script (always recommended)
# Then run with your tenant admin URL
.\PreProvision-OneDrive-E5-Group.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

**Parameters:**
- `-TenantAdminUrl` (mandatory) – Your SharePoint Admin Center URL
- `-GroupObjectId` (optional) – Defaults to the E5 licensing group
- `-WhatIf` – Dry run mode
- `-BatchSize` – Control batch size for provisioning requests (default 100)
- `-OutputCsvPath` – Custom CSV location

See the script header (Get-Help) for full documentation and examples.

**v1.1 Update:** Fixed the common `Import-Module Microsoft.Graph` error ("function capacity 4096 has been exceeded") by using selective sub-module imports instead of the monolithic package.

## How to Use These Scripts

1. Download the .ps1 file(s)
2. Review the code and comments
3. Run in a PowerShell session with appropriate M365 admin permissions
4. Use the generated CSV reports as input for your migration platform (ShareGate, AvePoint, BitTitan, etc.)

## Prerequisites (Common to All Scripts)
- PowerShell 5.1 or 7+
- Microsoft Graph PowerShell SDK (selective modules)
- SharePoint Online Management Shell
- Appropriate M365 admin roles (Global Admin, SharePoint Admin, User Admin, etc.)
- MFA-enabled accounts supported via interactive login

## Contributing

These scripts are maintained for internal M365 migration projects. Feel free to fork, improve, and submit pull requests.

## License

Internal use for licensed Microsoft 365 tenants.

---
*Maintained by the M365 Data Migration Team*