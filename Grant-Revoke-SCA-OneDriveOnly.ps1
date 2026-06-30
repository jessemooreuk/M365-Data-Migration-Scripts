<#
.SYNOPSIS
    M365 Data Migration Script – Grant or Revoke SCA rights on ALL OneDrive sites only.

.DESCRIPTION
    Targets personal OneDrive sites (https://tenant-my.sharepoint.com/personal/*) exclusively.
    Use before/after migrating user OneDrives or when temporary admin access to all user content is required.

.PARAMETER TenantAdminUrl
    SharePoint Admin Center URL (e.g. https://contoso-admin.sharepoint.com)

.PARAMETER LoginName
    UPN of the migration/service account

.PARAMETER Action
    Grant or Revoke (default = Grant)

.EXAMPLE
    .\Grant-Revoke-SCA-OneDriveOnly.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -LoginName "mig-svc@contoso.onmicrosoft.com" -Action Grant

.EXAMPLE
    .\Grant-Revoke-SCA-OneDriveOnly.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -LoginName "mig-svc@contoso.onmicrosoft.com" -Action Revoke
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $true)]
    [string]$LoginName,

    [ValidateSet("Grant", "Revoke")]
    [string]$Action = "Grant"
)

$ErrorActionPreference = 'Continue'
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath   = "C:\MigrationLogs\SCA_OneDrive_$timestamp.log"
$failedCsv = "C:\MigrationLogs\Failed_OneDrive_SCA_$timestamp.csv"

Start-Transcript -Path $logPath -Append

try {
    Write-Host "=== M365 Data Migration – SCA for OneDrive Sites Only ($Action) ===" -ForegroundColor Cyan
    Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Yellow
    Connect-SPOService -Url $TenantAdminUrl

    Write-Host "Retrieving all OneDrive sites (this may take time on large tenants)..." -ForegroundColor Yellow
    
    $oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All | 
                     Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }

    $total = $oneDriveSites.Count
    Write-Host "Found $total OneDrive site(s) to process." -ForegroundColor Green

    if ($total -eq 0) {
        Write-Warning "No OneDrive sites found. Exiting."
        return
    }

    $failedSites   = @()
    $processed     = 0
    $successCount  = 0

    foreach ($site in $oneDriveSites) {
        $processed++
        $percent = [math]::Round(($processed / $total) * 100, 1)

        Write-Progress -Activity "Processing OneDrive SCA $Action" `
                       -Status "$processed of $total – $($site.Url)" `
                       -PercentComplete $percent

        try {
            $isAdmin = ($Action -eq "Grant")
            Set-SPOUser -Site $site.Url `
                        -LoginName $LoginName `
                        -IsSiteCollectionAdmin $isAdmin `
                        -ErrorAction Stop

            Write-Host "[$processed/$total] SUCCESS: $Action SCA on $($site.Url)" -ForegroundColor Green
            $successCount++
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Warning "[$processed/$total] FAILED on $($site.Url): $errMsg"
            $failedSites += [PSCustomObject]@{
                SiteUrl   = $site.Url
                Error     = $errMsg
                Timestamp = Get-Date
            }
        }
    }

    Write-Progress -Activity "Processing OneDrive SCA $Action" -Completed

    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host "Total OneDrive sites processed : $processed" -ForegroundColor White
    Write-Host "Successful                     : $successCount" -ForegroundColor Green
    Write-Host "Failed                         : $($failedSites.Count)" -ForegroundColor $(if ($failedSites.Count -gt 0) { "Red" } else { "Green" })

    if ($failedSites.Count -gt 0) {
        $failedSites | Export-Csv -Path $failedCsv -NoTypeInformation -Encoding UTF8
        Write-Host "Failed sites exported to   : $failedCsv" -ForegroundColor Yellow
    }

    Write-Host "`nScript completed. Full log: $logPath" -ForegroundColor Cyan
}
catch {
    Write-Error "Critical error: $($_.Exception.Message)"
}
finally {
    Stop-Transcript
}