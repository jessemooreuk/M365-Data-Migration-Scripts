<#
================================================================================
 M365 DATA MIGRATION SCRIPT
 Pre-Provision OneDrive Sites for Azure AD Security Group Members (E5 Licensing)
================================================================================

.SYNOPSIS
    Queries the specified Azure AD (Microsoft Entra ID) security group, 
    pre-provisions OneDrive for Business sites for all user members, 
    and exports a CSV containing UPNs and OneDrive URLs.

.DESCRIPTION
    This script is built for Microsoft 365 data migration projects.
    It targets the security group (ID: 09f7fb60-5675-4875-8374-313497f095da) 
    that is linked to M365 E5 licensing via group-based licensing.

    The script:
    1. Connects to Microsoft Graph (to read group + users)
    2. Retrieves all user members from the security group
    3. Batches and submits OneDrive pre-provisioning requests via SharePoint Online
    4. Constructs the standard OneDrive personal site URLs
    5. Exports a clean CSV report (DisplayName, UPN, OneDriveUrl, Status, Timestamp)

    Pre-provisioning is highly recommended before mailbox/OneDrive migrations, 
    large-scale sharing, or when users have never signed into OneDrive.

.AUTHOR
    M365 Data Migration Team

.VERSION
    1.2 - 29 Jun 2026

.PREREQUISITES
    - PowerShell 7.x strongly recommended (PowerShell 5.1 often hits function limits)
    - Run as a user with:
        • SharePoint Administrator or Global Admin role (for SPO)
        • Azure AD permissions to read groups/users (or Global Reader)
    - Internet access to login.microsoftonline.com and *.sharepoint.com
    - Modules will be auto-installed if missing (CurrentUser scope)

.USAGE
    1. Save this file as PreProvision-OneDrive-E5-Group.ps1
    2. Open PowerShell and navigate to the folder
    3. Run:
       .\PreProvision-OneDrive-E5-Group.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"
    4. Authenticate when prompted (Graph + SPO)
    5. Review the output CSV and transcript log

    Optional parameters:
       -GroupObjectId   (default is your E5 group)
       -OutputCsvPath   (default includes timestamp)
       -WhatIf          (dry-run mode - no provisioning requests sent)
       -BatchSize       (default 100 - safe for Request-SPOPersonalSite)

.EXAMPLE
    .\PreProvision-OneDrive-E5-Group.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" -WhatIf

.NOTES
    - Provisioning is asynchronous. Sites usually appear within minutes to 24 hours.
    - Run this script BEFORE any migration tool that requires OneDrive sites to exist.
    - The CSV can be used as input for migration tools (e.g. ShareGate, AvePoint, etc.).
    - Transcript log is created in the same folder for full audit trail.
    - v1.2: Removed all Import-Module calls for Microsoft.Graph sub-modules.
      The cmdlets now load on-demand when Connect-MgGraph is called.
      This resolves the persistent "function capacity 4096 exceeded" error.
================================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$GroupObjectId = "09f7fb60-5675-4875-8374-313497f095da",

    [Parameter(Mandatory = $true, HelpMessage = "https://yourtenant-admin.sharepoint.com")]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = ".\OneDrive_PreProvision_E5_Group_Report_$(Get-Date -Format yyyyMMdd_HHmmss).csv",

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 100,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

#--------------------------------------------------------------------------------
# INITIALISATION & LOGGING (M365 Migration Standard)
#--------------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

$scriptName = "PreProvision-OneDrive-E5-Group"
$startTime = Get-Date
$transcriptPath = ".\$scriptName`_Transcript_$(Get-Date -Format yyyyMMdd_HHmmss).log"

Start-Transcript -Path $transcriptPath -Append -Force | Out-Null

Write-Host "=================================================================================" -ForegroundColor Cyan
Write-Host " M365 DATA MIGRATION SCRIPT - Pre-Provision OneDrive for E5 Security Group" -ForegroundColor Cyan
Write-Host "=================================================================================" -ForegroundColor Cyan
Write-Host "Start Time          : $startTime" -ForegroundColor Gray
Write-Host "Target Group ID     : $GroupObjectId" -ForegroundColor Gray
Write-Host "Tenant Admin URL    : $TenantAdminUrl" -ForegroundColor Gray
Write-Host "Output CSV          : $OutputCsvPath" -ForegroundColor Gray
Write-Host "WhatIf Mode         : $WhatIf" -ForegroundColor Gray
Write-Host "=================================================================================" -ForegroundColor Cyan

#--------------------------------------------------------------------------------
# MODULE CHECK (No Import-Module for Graph to avoid capacity limit)
#--------------------------------------------------------------------------------
Write-Host "`n[1/5] Checking required modules (Graph modules load on-demand)..." -ForegroundColor Cyan

$graphModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Users"
)

foreach ($module in $graphModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "  Installing $module (CurrentUser scope)..." -ForegroundColor Yellow
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "    Installed successfully." -ForegroundColor Green
    } else {
        Write-Host "  $module already installed." -ForegroundColor Gray
    }
}

# SharePoint Online module still requires explicit import
if (-not (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")) {
    Write-Host "  Installing Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Yellow
    Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
Write-Host "  Microsoft.Online.SharePoint.PowerShell imported successfully." -ForegroundColor Green

#--------------------------------------------------------------------------------
# CONNECT TO MICROSOFT GRAPH
#--------------------------------------------------------------------------------
Write-Host "`n[2/5] Connecting to Microsoft Graph..." -ForegroundColor Cyan
$graphScopes = @("Group.Read.All", "GroupMember.Read.All", "User.Read.All")

try {
    Connect-MgGraph -Scopes $graphScopes -NoWelcome -ErrorAction Stop
    $context = Get-MgContext
    Write-Host "  Connected as: $($context.Account)" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Graph. Error: $_"
    Stop-Transcript | Out-Null
    exit 1
}

#--------------------------------------------------------------------------------
# CONNECT TO SHAREPOINT ONLINE
#--------------------------------------------------------------------------------
Write-Host "`n[3/5] Connecting to SharePoint Online Admin Center..." -ForegroundColor Cyan
try {
    Connect-SPOService -Url $TenantAdminUrl -ErrorAction Stop
    Write-Host "  Successfully connected to SPO Admin Center." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SharePoint Online. Verify you have SharePoint Admin rights. Error: $_"
    Stop-Transcript | Out-Null
    exit 1
}

#--------------------------------------------------------------------------------
# GET GROUP & MEMBERS
#--------------------------------------------------------------------------------
Write-Host "`n[4/5] Retrieving members from Azure AD Security Group..." -ForegroundColor Cyan

try {
    $group = Get-MgGroup -GroupId $GroupObjectId -ErrorAction Stop
    Write-Host "  Group found: $($group.DisplayName) (ID: $($group.Id))" -ForegroundColor Green
} catch {
    Write-Error "Group with ID $GroupObjectId not found or access denied. Error: $_"
    Stop-Transcript | Out-Null
    exit 1
}

# Get all members (paginated)
$groupMembers = Get-MgGroupMember -GroupId $GroupObjectId -All -ErrorAction Stop

$users = @()
foreach ($member in $groupMembers) {
    if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
        try {
            $user = Get-MgUser -UserId $member.Id `
                               -Property DisplayName, UserPrincipalName, Id, UserType `
                               -ErrorAction Stop
            
            if ($user.UserType -eq 'Member' -or $null -eq $user.UserType) {
                $users += $user
            }
        } catch {
            Write-Warning "Could not retrieve full details for member $($member.Id). Skipping."
        }
    }
}

if ($users.Count -eq 0) {
    Write-Warning "No user members found in the group. Exiting."
    Stop-Transcript | Out-Null
    exit 0
}

Write-Host "  Total user members found: $($users.Count)" -ForegroundColor Green

#--------------------------------------------------------------------------------
# PRE-PROVISION ONEDRIVE SITES (BATCHED)
#--------------------------------------------------------------------------------
Write-Host "`n[5/5] Pre-provisioning OneDrive sites..." -ForegroundColor Cyan

# Derive MySite host from TenantAdminUrl
$tenantName = ($TenantAdminUrl -replace 'https://', '' -replace '-admin\.sharepoint\.com.*', '').Trim()
$mySiteHost = "https://$tenantName-my.sharepoint.com"

Write-Host "  MySite Host URL   : $mySiteHost" -ForegroundColor Gray
Write-Host "  Batch size        : $BatchSize" -ForegroundColor Gray

$upnList = $users.UserPrincipalName

if (-not $WhatIf) {
    Write-Host "  Submitting pre-provisioning requests (this is asynchronous)..." -ForegroundColor Yellow
} else {
    Write-Host "  WHATIF MODE - No actual provisioning requests will be sent." -ForegroundColor Magenta
}

$batchNumber = 0
$totalBatches = [math]::Ceiling($upnList.Count / $BatchSize)

for ($i = 0; $i -lt $upnList.Count; $i += $BatchSize) {
    $batchNumber++
    $batch = $upnList[$i..([math]::Min($i + $BatchSize - 1, $upnList.Count - 1))]
    
    Write-Host "  Processing batch $batchNumber of $totalBatches ($($batch.Count) users)..." -ForegroundColor Cyan
    
    if (-not $WhatIf) {
        try {
            Request-SPOPersonalSite -UserEmails $batch -ErrorAction Stop
            Write-Host "    Batch $batchNumber submitted successfully." -ForegroundColor Green
        } catch {
            Write-Warning "    Error submitting batch $batchNumber : $_"
            Write-Warning "    Continuing with remaining batches..."
        }
    } else {
        Write-Host "    [WhatIf] Would have requested pre-provisioning for $($batch.Count) users." -ForegroundColor Magenta
    }
    
    # Small pause between batches (good migration practice)
    if ($i + $BatchSize -lt $upnList.Count) {
        Start-Sleep -Seconds 2
    }
}

#--------------------------------------------------------------------------------
# BUILD REPORT & EXPORT CSV
#--------------------------------------------------------------------------------
Write-Host "`nGenerating final CSV report..." -ForegroundColor Cyan

$report = @()
foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    $encodedUpn = ($upn.ToLower() -replace '@', '_' -replace '\.', '_')
    $oneDriveUrl = "$mySiteHost/personal/$encodedUpn"
    
    $report += [PSCustomObject]@{
        DisplayName            = $user.DisplayName
        UPN                    = $upn
        OneDriveUrl            = $oneDriveUrl
        PreProvisionRequested  = if ($WhatIf) { "No (WhatIf Mode)" } else { "Yes" }
        RequestTimestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        GroupName              = $group.DisplayName
        GroupObjectId          = $GroupObjectId
    }
}

# Sort for readability
$report = $report | Sort-Object DisplayName

$report | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8 -Force

Write-Host "`n=================================================================================" -ForegroundColor Green
Write-Host " SCRIPT COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "=================================================================================" -ForegroundColor Green
Write-Host "Total users processed     : $($users.Count)" -ForegroundColor White
Write-Host "CSV Report saved to       : $OutputCsvPath" -ForegroundColor White
Write-Host "Transcript log saved to   : $transcriptPath" -ForegroundColor White
Write-Host "=================================================================================" -ForegroundColor Green

Write-Host "`nNEXT STEPS (Recommended for M365 Migration):" -ForegroundColor Yellow
Write-Host "1. Open the CSV and verify the OneDrive URLs look correct." -ForegroundColor Yellow
Write-Host "2. Wait 15-60 minutes (or up to 24h for full provisioning)." -ForegroundColor Yellow
Write-Host "3. Validate a few sites by browsing to the URLs or using Get-SPOSite." -ForegroundColor Yellow
Write-Host "4. Use this CSV as source list for your migration tool." -ForegroundColor Yellow
Write-Host "5. Re-run with -WhatIf if you need to generate the list again without re-provisioning." -ForegroundColor Yellow

Stop-Transcript | Out-Null