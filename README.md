# M365 Data Migration Scripts

A collection of ready-to-use PowerShell scripts for Microsoft 365 data migration projects.

These scripts are designed to help with common pre- and post-migration tasks such as granting temporary permissions, cleaning up access, and preparing sites/OneDrives for migration tools (SPMT, ShareGate, etc.).

## Scripts

### OneDrive SCA Grant / Revoke
**File:** `Grant-Revoke-SCA-OneDriveOnly.ps1`

Grants or revokes Site Collection Administrator rights **only on personal OneDrive sites**.

**Use case:** Before migrating user OneDrives or when a migration account needs temporary access to all user content libraries.

See the top of the script for full parameters, examples, and warnings about large tenants.

## Getting Started

```powershell
# Install prerequisite module
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force

# Grant SCA on all OneDrives
.\Grant-Revoke-SCA-OneDriveOnly.ps1 `
    -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    -LoginName "migration-svc@contoso.onmicrosoft.com" `
    -Action Grant

# Revoke after migration
.\Grant-Revoke-SCA-OneDriveOnly.ps1 `
    -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    -LoginName "migration-svc@contoso.onmicrosoft.com" `
    -Action Revoke
```

## Best Practices
- Always revoke access after migration is complete and validated.
- For large tenants, consider running against a specific batch of users instead of the entire tenant.
- Run scripts with a dedicated migration service account.
- Review the transcript logs and failed CSV output.

More scripts (full tenant SCA, targeted site lists, PnP versions, etc.) will be added over time.

Contributions and suggestions welcome!