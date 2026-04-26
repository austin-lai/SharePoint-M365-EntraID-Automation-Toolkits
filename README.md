# SharePoint, M365, and EntraID Automation Toolkits

```markdown
> Austin.Lai |
> -----------| April 25th, 2026
> -----------| Updated on April 26th, 2026
```

<!-- 
TODO:

- Check all command and script
- Check all module require inside the script
- Check the functionality
- Check if possible to refine it
- Check description or can have better one
 -->

---

## Table of Contents

<!-- TOC -->

- [SharePoint, M365, and EntraID Automation Toolkits](#sharepoint-m365-and-entraid-automation-toolkits)
    - [Table of Contents](#table-of-contents)
    - [Disclaimer](#disclaimer)
    - [📌 Description](#-description)
    - [SharePoint Toolkits](#sharepoint-toolkits)
        - [Simple PnPOnline Connect Command](#simple-pnponline-connect-command)
        - [Reset SharePoint Home Page back to Home.aspx of SitePages](#reset-sharepoint-home-page-back-to-homeaspx-of-sitepages)
        - [SharePoint change NewPosts back to Page](#sharepoint-change-newposts-back-to-page)
        - [SharePoint - Clear SitePageFlags](#sharepoint---clear-sitepageflags)
        - [SharePoint - Grant Enterprise Application with Write access to ALL sites](#sharepoint---grant-enterprise-application-with-write-access-to-all-sites)
    - [M365 or EntraID Toolkits](#m365-or-entraid-toolkits)
        - [M365 - Retrieve Group.Unified information](#m365---retrieve-groupunified-information)
            - [Method 1 - Using Microsoft.Graph to retrieve Group.Unified information](#method-1---using-microsoftgraph-to-retrieve-groupunified-information)
            - [Method 2 - Using Invoke-MgGraphRequest to retrieve Group.Unified information](#method-2---using-invoke-mggraphrequest-to-retrieve-groupunified-information)
        - [M365 - To enable Microsoft 365 Group and Teams creation for M365 User - Simple functionalities](#m365---to-enable-microsoft-365-group-and-teams-creation-for-m365-user---simple-functionalities)
        - [M365 - To enable Microsoft 365 Group and Teams creation for M365 User - Complex functionalities](#m365---to-enable-microsoft-365-group-and-teams-creation-for-m365-user---complex-functionalities)
        - [M365 - Restrict Microsoft 365 Group and Teams creation - Forced and disabled M365 user from creating M365 Group](#m365---restrict-microsoft-365-group-and-teams-creation---forced-and-disabled-m365-user-from-creating-m365-group)
        - [Automation to ensure all tenant enabled GDAP Auto-Extend](#automation-to-ensure-all-tenant-enabled-gdap-auto-extend)
    - [Utilities & Supporting Toolkits](#utilities--supporting-toolkits)
        - [SharePoint - PowerShell Modules required in this repo](#sharepoint---powershell-modules-required-in-this-repo)
        - [SharePoint - REST API to get Site Page ID](#sharepoint---rest-api-to-get-site-page-id)
        - [Invoke-MgGraphRequest command error](#invoke-mggraphrequest-command-error)
        - [Use -DisableNameChecking when import Microsoft.Online.Sharepoint.PowerShell](#use--disablenamechecking-when-import-microsoftonlinesharepointpowershell)

<!-- /TOC -->

<br>

## Disclaimer

> [!WARNING]
> This repository is licensed under the Apache License 2.0.
>
> All scripts and content are provided for **educational and reference purposes only**. They are not production-ready by default and may require modification to suit your specific environment, tenant configuration, and security policies.
>
> By using this repository, you acknowledge that:
>
> - You are responsible for reviewing and understanding the code before execution  
> - You will test all scripts in a **non-production environment** prior to use  
> - You will validate required permissions, scopes, and potential impact  
> - You assume full responsibility for any outcomes resulting from usage  
>
> The author shall not be held responsible for any misuse, damage, data loss, service disruption, or security issues resulting from the use of this repository.
>
> This project is **not affiliated with, endorsed by, or supported by Microsoft**.

<br>

## 📌 Description

<!-- Description -->

A curated collection of PowerShell scripts for Microsoft 365 and SharePoint Online administration, automation, and operational tasks.

This repository contains a growing set of reusable scripts designed to simplify common administrative workflows across SharePoint Online, Microsoft 365, and related services such as Entra ID and Microsoft Graph. The focus is on practical, real-world operations including site management, permission handling, reporting, and tenant-level automation.

<br>

> [!NOTE]
> 🧰 What’s Included:
>
> - SharePoint Online Management
>     - Site creation, configuration, and cleanup
>     - Hub site registration and deregistration
>     - Page operations (e.g. News ↔ Page conversion, homepage reset)
>     - File and content enumeration
> - Permissions & Access Control
>     - Assigning site admins and user roles
>     - App-level access (tenant-wide permissions)
>     - Group and M365 role management
> - Automation & Reporting
>     - Tenant configuration exports
>     - SharePoint structure and template exports
>     - Health checks and diagnostics
> - Microsoft Graph / Entra ID Integration
>     - App registration automation in Entra ID
>     - Graph-based operations (e.g. Copilot chat export)
>     - GDAP-related automation
> - Utilities & Supporting Scripts
>     - Bulk file listing and analysis
>     - Module installation helpers
>     - REST API examples and troubleshooting notes

<br>

🎯 Purpose

This repo is intended for:

- SharePoint or M365 administrators
- IT or Security engineers managing enterprise tenants
- Security and automation engineers working with Microsoft cloud

The scripts are built from hands-on operational use cases, not theoretical examples.

<br>

🚀 Notes

- Some scripts require:
    - PnP PowerShell
    - Microsoft Graph PowerShell SDK
    - SharePoint Online Management Shell
- App registration steps for Entra ID are included for automation scenarios.

<!-- /Description -->

<br>

## SharePoint Toolkits

### Simple PnPOnline Connect Command

Based on my experience, the most effective approach for connecting to SharePoint Online is to use `Connect-PnPOnline` with explicitly defined tenant and Client ID values to ensure consistent and secure authentication.

> [!TIP]
> Of course, you can always use a better and more secure method like using Non interactive Authentication using a certificate file.

The command structure as below:

```
Connect-PnPOnline -Url "[SHAREPOINT_SITE-or-SHAREPOINT_ADMIN_SITE]" -ClientId [CLIENT_ID] -Tenant "[TENANT_DOMAIN]" -Interactive
```

<br>

### Reset SharePoint Home Page back to Home.aspx of SitePages

This toolkit creation is specifically designed to address a common scenario:

- Restoring a SharePoint site to its original configuration, particularly when dealing with changes to the home page or site structure. It’s a streamlined process for quickly returning to a predefined state, ensuring a consistent and reliable setup.

```powershell
# 1) Connect to the affected site (modern auth; will reuse token for subsequent calls)
Import-Module PnP.PowerShell
Connect-PnPOnline -Url "[SHAREPOINT_SITE-or-SHAREPOINT_ADMIN_SITE]" -ClientId [CLIENT_ID] -Tenant "[TENANT_DOMAIN]" -Interactive


# 2) Reset the site home page to modern (usually SitePages/Home.aspx)
# If your modern home page has a different name, update the path accordingly.
try {
    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Home.aspx"
    Write-Host "[OK] Home page reset to SitePages/Home.aspx" -ForegroundColor Green
} catch {
    Write-Warning "Home page reset failed: $($_.Exception.Message)"
}


# 3) Turn OFF In-Place Records Management at the site collection level
# (This stops manual record declaration behaviors that can surface 'review' flows.)
try {
    Set-PnPInPlaceRecordsManagement -Enabled:$false
    Write-Host "[OK] In-Place Records Management disabled" -ForegroundColor Green
} catch {
    Write-Warning "Failed to disable In-Place Records Management: $($_.Exception.Message)"
}


# 4) Remove the 'Review Items' list (optional—only if you don't need it)
# If you want to keep it, skip this block.
$reviewList = Get-PnPList -Identity "Review Items" -ErrorAction SilentlyContinue
if ($reviewList) {
    try {
        Remove-PnPList -Identity "Review Items" -Force
        Write-Host "[OK] 'Review Items' list removed" -ForegroundColor Green
    } catch {
        Write-Warning "Couldn't remove 'Review Items': $($_.Exception.Message)"
    }
} else {
    Write-Host "[INFO] 'Review Items' list not found on this site." -ForegroundColor Yellow
}


# 5) (Optional) If modern still doesn’t show, ensure no custom master/alternate CSS is applied, and clear the DenyAddAndCustomizePages flag using SPO cmdlets (requires Connect-SPOService).
# Uncomment and run if needed:
# Connect-SPOService -Url "https://<tenant>-admin.sharepoint.com"
# Set-SPOSite -Identity $siteUrl -DenyAddAndCustomizePages 0
# Write-Host "[OK] DenyAddAndCustomizePages set to 0 (allows modern pages)" -ForegroundColor Green
```

<br>

### SharePoint change NewPosts back to Page

This toolkit creation is specifically designed to address a common scenario:

- Restoring a SharePoint NewPosts to Page.
- Useful when a Posts or NewPosts being promoted and user want to reset it back to Page for further tuning or drafting in order to re-promote to NewPosts to reach wider audiences.

```powershell
# 1) Install/Import PnP.PowerShell (once per machine)
# Install-Module PnP.PowerShell -Scope CurrentUser


# 2) Connect with modern auth
$SiteUrl = "[SHAREPOINT_SITE-or-SHAREPOINT_ADMIN_SITE]"
Connect-PnPOnline -Url $SiteUrl -ClientId [CLIENT_ID] -Tenant "[TENANT_DOMAIN]" -Interactive


# 3) Find the list item for your page by FileLeafRef (file name)
$FileName = "[NAME].aspx"
$item = Get-PnPListItem -List "Site Pages" -PageSize 2000 | Where-Object { $_["FileLeafRef"] -eq $FileName }
if ($null -eq $item) {
    Write-Host "Item not found in Site Pages: $FileName" -ForegroundColor Red
    return
}


# 4) Demote the News post to a normal Page by setting PromotedState to 0
Set-PnPListItem -List "Site Pages" -Identity $item.Id -Values @{ "PromotedState" = "0" }


# 5) (Optional) Re-publish to make sure the change is visible
# You can simply open the page and click Publish/Republish in the UI, or call Publish-PnPClientSidePage if you prefer scripting:
# Publish-PnPClientSidePage -Identity $FileName


Write-Host "Done. $FileName should now show Publish/Republish instead of Post/Update news."
```

<br>

### SharePoint - Clear SitePageFlags

This toolkit is useful when you dealing with messy SharePoint page's flag or you wish to clean up the SharePoint page's flag.

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [string]$LibraryTitle = "Site Pages",
    [string]$ColumnInternalName = "_SPSitePageFlags",

    [switch]$OnlyWhenHasValue,
    [switch]$WhatIf,

    [string]$BackupCsvPath = "$(Join-Path $PWD ('SitePageFlags_Backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'))"
)

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
# $SiteUrl = "[SHAREPOINT_SITE-or-SHAREPOINT_ADMIN_SITE]"
Connect-PnPOnline -Url $SiteUrl -ClientId [CLIENT_ID] -Tenant "[TENANT_DOMAIN]" -Interactive

$list = Get-PnPList -Identity $LibraryTitle -ErrorAction Stop
Write-Host "Found list: $($list.Title)" -ForegroundColor Green

# Verify field
$field = Get-PnPField -List $LibraryTitle | Where-Object { $_.InternalName -eq $ColumnInternalName }
if (-not $field) {
    Write-Host "ERROR: Field '$ColumnInternalName' not found in '$LibraryTitle'." -ForegroundColor Red
    exit 1
}
Write-Host "Field: Title='$($field.Title)', InternalName='$($field.InternalName)', Type='$($field.TypeAsString)'" -ForegroundColor Green

# Load items
$items = Get-PnPListItem -List $LibraryTitle -PageSize 2000 -Fields "FileLeafRef","ID",$ColumnInternalName
$pages = $items | Where-Object { $_["FileLeafRef"] -and ($_[ "FileLeafRef" ].ToLower().EndsWith(".aspx")) }

Write-Host ("Total items: {0}; .aspx pages: {1}" -f $items.Count, $pages.Count) -ForegroundColor Cyan

# Build candidates
$candidates = foreach ($p in $pages) {
    $name = $p["FileLeafRef"]; $id = $p.Id; $val = $p[$ColumnInternalName]
    $currentText = if ($null -ne $val -and ($val -is [System.Collections.IEnumerable]) -and -not ($val -is [string])) { ($val -join "; ") } else { [string]$val }
    $hasValue = -not [string]::IsNullOrWhiteSpace($currentText)
    if (-not $OnlyWhenHasValue -or $hasValue) {
        [pscustomobject]@{ ID = $id; Name = $name; Flags = $currentText }
    }
}

Write-Host ("Candidates to clear: {0}" -f $candidates.Count) -ForegroundColor Cyan
if ($candidates.Count -eq 0) { Write-Host "Nothing to do." -ForegroundColor Green; exit 0 }

# Backup
Write-Host "Writing backup to: $BackupCsvPath" -ForegroundColor Cyan
$candidates | Export-Csv -Path $BackupCsvPath -NoTypeInformation -Encoding UTF8

# Detect supported update parameter
$setCmd = Get-Command Set-PnPListItem
$hasUpdateType = ($setCmd.Parameters.Keys -contains 'UpdateType')
$hasSystemUpdate = ($setCmd.Parameters.Keys -contains 'SystemUpdate')

Write-Host ("Set-PnPListItem supports -UpdateType: {0}; -SystemUpdate: {1}" -f $hasUpdateType, $hasSystemUpdate) -ForegroundColor Cyan

# Clear flags
$errors = @(); $updated = @()

foreach ($row in $candidates) {
    $id = $row.ID; $name = $row.Name; $old = $row.Flags
    Write-Host ("{0} -> clearing flags: '{1}'" -f $name, $old) -ForegroundColor Yellow

    try {
        if ($WhatIf) {
            Write-Host ("WHATIF: Would clear '{0}' on item ID {1}" -f $ColumnInternalName, $id) -ForegroundColor DarkYellow
            continue
        }

        $values = @{ $ColumnInternalName = @() }   # MultiChoice: clear with empty array
        if ($hasUpdateType) {
            Set-PnPListItem -List $LibraryTitle -Identity $id -Values $values -UpdateType SystemUpdate
        }
        elseif ($hasSystemUpdate) {
            Set-PnPListItem -List $LibraryTitle -Identity $id -Values $values -SystemUpdate
        }
        else {
            # Fallback: normal update (will bump version/modified)
            Set-PnPListItem -List $LibraryTitle -Identity $id -Values $values
        }

        $updated += $row
    }
    catch {
        # Try a secondary attempt with $null for stubborn MultiChoice fields
        try {
            Set-PnPListItem -List $LibraryTitle -Identity $id -Values @{ $ColumnInternalName = $null }
            $updated += $row
            Write-Host ("{0} -> cleared via fallback ($null)" -f $name) -ForegroundColor Green
        }
        catch {
            $errors += [pscustomobject]@{ ID = $id; Name = $name; Error = $_.Exception.Message }
            Write-Host ("ERROR on {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
        }
    }
}

Write-Host ("Updated items: {0}" -f $updated.Count) -ForegroundColor Green
if ($errors.Count -gt 0) {
    $errPath = "$(Join-Path $PWD ('SitePageFlags_Errors_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'))"
    $errors | Export-Csv -Path $errPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Errors: {0} | Log: {1}" -f $errors.Count, $errPath) -ForegroundColor Red
}

Write-Host "Done." -ForegroundColor Green
```

<br>

### SharePoint - Grant Enterprise Application with `Write` access to ALL sites

This toolkit is useful when you are bulk managed all the SharePoint sites.

```powershell
# Connect as SharePoint Admin (interactive)
$SiteUrl = "[SHAREPOINT_SITE-or-SHAREPOINT_ADMIN_SITE]"
Connect-PnPOnline -Url $SiteUrl -ClientId [CLIENT_ID] -Tenant "[TENANT_DOMAIN]" -Interactive


# Replace the -ClientId to your own client-id and display name
$AppId       = "<YOUR-APP-CLIENT-ID>"
$DisplayName = "Austin-Copilot+PnP"


# Get all sites exclude OneDrive sites; include team/communication sites
$sites = Get-PnPTenantSite -IncludeOneDriveSites:$false


# Grant Write access to all sites with error handling
foreach ($site in $sites) {
    try {
        Grant-PnPAzureADAppSitePermission `
          -Site $site.Url `
          -AppId $AppId `
          -DisplayName $DisplayName `
          -Permissions Write

        Write-Host "Granted access to: $($site.Url)" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to grant access to: $($site.Url): $($_.Exception.Message)"
    }
}
```

<br><br>

## M365 or EntraID Toolkits

### M365 - Retrieve `Group.Unified` information

#### Method 1 - Using `Microsoft.Graph` to retrieve `Group.Unified` information

```powershell
# Requires -Module Microsoft.Graph.Groups
Connect-MgGraph -Scopes "User.Read.All,Directory.ReadWrite.All","Group.Read.All","GroupSettings.Read.All"  -NoWelcome 

# Fetch and display the 'Group.Unified' directory setting template
Get-MgBetaDirectorySettingTemplate | Where-Object DisplayName -eq 'Group.Unified'
```

Sample output as below:

```powershell
DeletedDateTime Id                                   Description
--------------- --                                   -----------
                62375ab9-6b52-47ed-826b-58e47e0e304b …
```

<br>

#### Method 2 - Using `Invoke-MgGraphRequest` to retrieve `Group.Unified` information

```powershell
# Fetch and display all settings for 'Group.Unified'
Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/groupSettings' |`
  Select-Object -ExpandProperty value |`
  Where-Object displayName -eq 'Group.Unified' |`
  ForEach-Object { $_.values | Format-Table name, value -AutoSize }
```

Sample output as below:

```powershell
Name  Value
----  -----
value true
name  NewUnifiedGroupWritebackDefault
value false
name  EnableMIPLabels
value
name  CustomBlockedWordsList
value false
name  EnableMSStandardBlockedWords
value
name  ClassificationDescriptions
value
name  DefaultClassification
value
name  PrefixSuffixNamingRequirement
value false
name  AllowGuestsToBeGroupOwner
value true
name  AllowGuestsToAccessGroups
value
name  GuestUsageGuidelinesUrl
value 81036dcd-c8d9-4ac5-b017-4f3426c0859c
name  GroupCreationAllowedGroupId
value true
name  AllowToAddGuests
value
name  UsageGuidelinesUrl
value
name  ClassificationList
value false
name  EnableGroupCreation 
```

<br>

### M365 - To enable Microsoft 365 Group (and Teams) creation for M365 User - Simple functionalities

```powershell
Import-Module Microsoft.Graph.Beta.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Beta.Groups

Connect-MgGraph -Scopes "Directory.ReadWrite.All", "Group.Read.All"

$GroupName = ""
$AllowGroupCreation = "False"

$settingsObjectID = (Get-MgBetaDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id

if(!$settingsObjectID)
{
    $params = @{
      templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
      values = @(
            @{
                   name = "EnableMSStandardBlockedWords"
                   value = $true
             }
              )
         }
    
    New-MgBetaDirectorySetting -BodyParameter $params
    
    $settingsObjectID = (Get-MgBetaDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).Id
}

$groupId = (Get-MgBetaGroup -all | Where-object {$_.displayname -eq $GroupName}).Id

$params = @{
    templateId = "62375ab9-6b52-47ed-826b-58e47e0e304b"
    values = @(
        @{
            name = "EnableGroupCreation"
            value = $AllowGroupCreation
        }
        @{
            name = "GroupCreationAllowedGroupId"
            value = $groupId
        }
    )
}

Update-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID -BodyParameter $params

(Get-MgBetaDirectorySetting -DirectorySettingId $settingsObjectID).Values
```

<br>

### M365 - To enable Microsoft 365 Group (and Teams) creation for M365 User - Complex functionalities

```powershell
# [ROLLBACK] To restore default (everyone can create groups), run these GA REST calls:
Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/groupSettings/8da592cd-e14c-484c-8afd-3ab9bc06ec7c' -Body (@{ values = @(@{name='EnableGroupCreation'; value='true' }, @{name='GroupCreationAllowedGroupId'; value=''}) } | ConvertTo-Json -Depth 6)
```

<br>

### M365 - Restrict Microsoft 365 Group (and Teams) creation - Forced and disabled M365 user from creating M365 Group

```powershell
# To Check existing Group.Unified settings:
# Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/groupSettings' |
# >>   Select-Object -ExpandProperty value |
# >>   Where-Object displayName -eq 'Group.Unified' |
# >>   ForEach-Object { $_.values | Format-Table name, value -AutoSize }

# OR #
# USE #
# get-Group.Unified-info-final-v1-30122025.ps1



# [ROLLBACK] To restore default (everyone can create groups), run these GA REST calls:
# Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/groupSettings/8da592cd-e14c-484c-8afd-3ab9bc06ec7c' -Body (@{ values = @(@{name='EnableGroupCreation'; value='true' }, @{name='GroupCreationAllowedGroupId'; value=''}) } | ConvertTo-Json -Depth 6)



<#
.SYNOPSIS
  Restrict Microsoft 365 Group (and Teams) creation to a specific security group by configuring tenant-wide "Group.Unified" settings using GA REST (Invoke-MgGraphRequest).

.DESCRIPTION
  - Creates or updates the tenant-wide Microsoft 365 Group settings (Group.Unified) via /v1.0/groupSettings.
  - Sets EnableGroupCreation=false (blocks everyone).
  - Sets GroupCreationAllowedGroupId=<ObjectId of your "Allowed Group Creators" security group> (re-enables only those members).
  - Optionally creates the security group and adds specified members.
  - Verbose output and optional transcript. No progress bars.

.REFERENCES
  - Create tenant-wide settings (POST /groupSettings) and Group.Unified usage in v1.0: https://learn.microsoft.com/en-us/graph/api/group-post-settings?view=graph-rest-1.0
  - Microsoft Entra group settings cmdlets (template ID example for Group.Unified = 62375ab9-6b52-47ed-826b-58e47e0e304b): https://learn.microsoft.com/en-us/entra/identity/users/groups-settings-cmdlets
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $AllowedGroupDisplayName = 'Queenstown IT - Allowed Group Creators',

    [Parameter(Mandatory = $false)]
    [switch] $AutoCreateAllowedGroupIfMissing = $true,

    [Parameter(Mandatory = $false)]
    [string[]] $AddMembersUpn = @(),  # e.g. 'admin@contoso.com','alice@contoso.com'

    [Parameter(Mandatory = $false)]
    [switch] $EnableTranscript,

    [Parameter(Mandatory = $false)]
    [string] $TranscriptPath = ".\GroupUnified_Transcript.log"
)

# GA template ID from Microsoft docs for Group.Unified
$GroupUnifiedTemplateId = '62375ab9-6b52-47ed-826b-58e47e0e304b' # [2](https://learn.microsoft.com/en-us/answers/questions/1186794/what-is-the-limtation-on-creating-number-of-aad-gr)

function Ensure-GraphAuth {
    Write-Verbose "[Ensure-GraphAuth] Installing/Importing Microsoft.Graph..."
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph -ErrorAction Stop

    # Request all scopes we might need (least privileged that cover this op)
    # v1.0 tenant-wide group settings creation/update requires GroupSettings.ReadWrite.All + Directory.ReadWrite.All
    $scopes = @('Directory.ReadWrite.All','Group.Read.All','Group.ReadWrite.All','GroupSettings.ReadWrite.All')
    Write-Verbose "[Ensure-GraphAuth] Connecting with scopes: $($scopes -join ', ')"
    Connect-MgGraph -Scopes $scopes -NoWelcome

    $ctx = Get-MgContext
    $org = Get-MgOrganization | Select-Object -First 1
    Write-Host ("[INFO] Connected. Tenant: {0}  Account: {1}" -f $org.Id, $ctx.Account) -ForegroundColor Green
}

function Get-OrCreate-AllowedSecurityGroup {
    param([string] $DisplayName, [switch] $AutoCreate)

    Write-Verbose "[Get-OrCreate-AllowedSecurityGroup] Searching for '$DisplayName'..."
    $group = Get-MgGroup -Filter "displayName eq '$DisplayName'" -ConsistencyLevel eventual -All | Where-Object {
        $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false
    }

    if ($group) {
        Write-Host "[INFO] Found security group: $($group.DisplayName) [$($group.Id)]" -ForegroundColor Green
        return $group
    }

    if ($AutoCreate) {
        Write-Host "[INFO] Creating security group '$DisplayName'..." -ForegroundColor Yellow
        $mailNickname = ($DisplayName -replace '[^A-Za-z0-9]', '').ToLower()
        if ([string]::IsNullOrWhiteSpace($mailNickname)) { $mailNickname = "allowedgroupcreators" }

        $new = New-MgGroup `
            -DisplayName $DisplayName `
            -MailEnabled:$false `
            -SecurityEnabled:$true `
            -MailNickname $mailNickname

        Write-Host "[INFO] Created security group: $($new.DisplayName) [$($new.Id)]" -ForegroundColor Green
        return $new
    }

    throw "Security group '$DisplayName' not found and AutoCreateAllowedGroupIfMissing is disabled."
}

function Add-Members-ToGroup {
    param([string] $GroupId, [string[]] $UsersUpn)

    if (-not $UsersUpn -or $UsersUpn.Count -eq 0) { Write-Verbose "[Add-Members-ToGroup] No members to add."; return }

    Write-Host "[INFO] Adding members to group [$GroupId]..." -ForegroundColor Cyan
    foreach ($upn in $UsersUpn) {
        try {
            Write-Verbose "[Add-Members-ToGroup] Resolving user: $upn"
            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ConsistencyLevel eventual -All
            if (-not $user) { Write-Host "[WARN] User not found: $upn" -ForegroundColor Yellow; continue }

            Write-Verbose "[Add-Members-ToGroup] Checking membership for $upn..."
            $existing = Get-MgGroupMember -GroupId $GroupId -All | Where-Object { $_.Id -eq $user.Id }
            if ($existing) { Write-Host "[INFO] Already a member: $upn" -ForegroundColor DarkGreen; continue }

            Write-Verbose "[Add-Members-ToGroup] Adding $upn by reference..."
            $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" }
            Add-MgGroupMemberByRef -GroupId $GroupId -BodyParameter $body
            Write-Host "[INFO] Added: $upn" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to add $upn : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Get-TenantGroupSettings {
    # v1.0: GET /groupSettings (tenant-wide settings)  [4](https://learn.microsoft.com/en-us/graph/api/group-post-settings?view=graph-rest-1.0)
    Write-Verbose "[Get-TenantGroupSettings] Listing existing tenant-wide group settings (v1.0)..."
    $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/groupSettings'
    return $resp.value
}

function New-TenantGroupSettings {
    param([string] $AllowedGroupId)

    # v1.0: POST /groupSettings, using published Group.Unified template ID  [4](https://learn.microsoft.com/en-us/graph/api/group-post-settings?view=graph-rest-1.0)[2](https://learn.microsoft.com/en-us/answers/questions/1186794/what-is-the-limtation-on-creating-number-of-aad-gr)
    Write-Verbose "[New-TenantGroupSettings] Creating Group.Unified settings..."
    $payload = @{
        templateId = $GroupUnifiedTemplateId
        values     = @(
            @{ name = 'EnableGroupCreation';          value = 'false' }
            @{ name = 'GroupCreationAllowedGroupId'; value = $AllowedGroupId }
        )
    }
    $created = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/groupSettings' -Body ($payload | ConvertTo-Json -Depth 6)
    Write-Host "[INFO] Created Group.Unified settings: $($created.id)" -ForegroundColor Green
    return $created
}

function Update-TenantGroupSettings {
    param([string] $SettingId, [string] $AllowedGroupId)

    # v1.0: PATCH /groupSettings/{id}  [4](https://learn.microsoft.com/en-us/graph/api/group-post-settings?view=graph-rest-1.0)
    Write-Verbose "[Update-TenantGroupSettings] Updating settings $SettingId..."
    $payload = @{
        values = @(
            @{ name = 'EnableGroupCreation';          value = 'false' }
            @{ name = 'GroupCreationAllowedGroupId'; value = $AllowedGroupId }
        )
    }
    Invoke-MgGraphRequest -Method PATCH -Uri ("https://graph.microsoft.com/v1.0/groupSettings/{0}" -f $SettingId) -Body ($payload | ConvertTo-Json -Depth 6) | Out-Null
    Write-Host "[INFO] Updated Group.Unified settings: $SettingId" -ForegroundColor Green
}

function Ensure-GroupUnifiedSetting {
    param([string] $AllowedGroupId)

    $settings = Get-TenantGroupSettings
    $unified  = $settings | Where-Object { $_.displayName -eq 'Group.Unified' }

    if (-not $unified) {
        $created = New-TenantGroupSettings -AllowedGroupId $AllowedGroupId
        return $created
    }
    else {
        Update-TenantGroupSettings -SettingId $unified.id -AllowedGroupId $AllowedGroupId
        # Re‑read for verification
        $settings2 = Get-TenantGroupSettings
        return ($settings2 | Where-Object { $_.id -eq $unified.id })
    }
}

function Show-Verification {
    param($SettingObject)

    Write-Host "[INFO] Applied values:" -ForegroundColor Green
    $SettingObject.values | Sort-Object name | Format-Table name, value -AutoSize
}

function Show-Rollback {
    param([string] $SettingId)

    Write-Host "`n[ROLLBACK] To restore default (everyone can create groups), run these GA REST calls:" -ForegroundColor Yellow
    Write-Host "Invoke-MgGraphRequest -Method PATCH -Uri 'https://graph.microsoft.com/v1.0/groupSettings/$SettingId' -Body (@{ values = @(@{name='EnableGroupCreation'; value='true' }, @{name='GroupCreationAllowedGroupId'; value=''}) } | ConvertTo-Json -Depth 6)" -ForegroundColor White
}

# ----------------------------- Main -----------------------------
$transcriptStarted = $false
try {
    if ($EnableTranscript) {
        $dir = Split-Path -Path $TranscriptPath -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Start-Transcript -Path $TranscriptPath -Append -ErrorAction Stop
        $transcriptStarted = $true
        Write-Host "[INFO] Transcript started: $TranscriptPath" -ForegroundColor Cyan
    }

    Ensure-GraphAuth

    # 1) Resolve/create the allowed creators security group
    $allowedGroup = Get-OrCreate-AllowedSecurityGroup -DisplayName $AllowedGroupDisplayName -AutoCreate:$AutoCreateAllowedGroupIfMissing
    Write-Host "[INFO] Allowed creators group ObjectId: $($allowedGroup.Id)" -ForegroundColor Green

    # 2) Optionally add members
    if ($AddMembersUpn.Count -gt 0) {
        Add-Members-ToGroup -GroupId $allowedGroup.Id -UsersUpn $AddMembersUpn
    }

    # 3) Configure Group.Unified tenant-wide settings
    $unifiedSetting = Ensure-GroupUnifiedSetting -AllowedGroupId $allowedGroup.Id
    Show-Verification -SettingObject $unifiedSetting
    Show-Rollback -SettingId $unifiedSetting.id

    Write-Host "`n[INFO] Configuration complete." -ForegroundColor Green
    Write-Host "[NOTE] Allow up to ~24–36 hours for Teams/Outlook/Planner to fully reflect the restriction in client UIs." -ForegroundColor Yellow
}
catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    throw
}
finally {
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
        Write-Host "[INFO] Transcript stopped." -ForegroundColor Cyan
    }
}
```

<br>

### Automation to ensure all tenant enabled GDAP Auto-Extend

This is a useful toolkit when you working as MSSP or in whichever circumstances that manage multiple tenant within your organization.

> [!IMPORTANT]
> Important Reminders for GDAP Remediation
>
> - **Global Admin Role**: Relationships containing the Global Administrator role are ineligible for auto-extend. Any script will fail to update these unless that role is first removed.
> - **Permissions Required**: To run these remediations, your service principal or admin account needs the DelegatedAdminRelationship.ReadWrite.All permission in Microsoft Graph.
> - **Partner Center UI Alternative**: You can now bulk-enable this for up to 25 customers at a time directly in the Partner Center under Customers > Expiring Granular Relationships, which may be faster than troubleshooting a custom script.

You can refer to [Enable GDAP auto-extend using Microsoft Graph and PowerShell](https://tommygjertsen.com/enable-gdap-auto-extend/#:~:text=Description,autoExtendDuration%E2%80%9D%2C%20to%20enable%20it.) for futher reading.

```powershell
<#
.SYNOPSIS
  Audits Microsoft Partner GDAP (delegated admin) relationships and enables auto-extend in bulk.
  Optionally removes the Global Administrator role from relationships (only when status=active) to allow auto-extend.

.SCRIPT NAME
  Invoke-GdapAutoExtendRemediation.ps1

.REQUIREMENTS
  - PowerShell 5.1+ (or 7+)
  - Microsoft Graph PowerShell SDK
  - Permission scope: DelegatedAdminRelationship.ReadWrite.All (delegated sign-in)

.NOTES (Grounded)
  - Can't auto-extend if Global Admin is included; remove it to become eligible [1](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/beta/api/delegatedadminrelationship-update.md)
  - Removing Global Admin is only allowed when status is active and can be long-running (202 Accepted possible)
  - autoExtendDuration supports PT0S (off) and P180D (on) among supported values
  - Update relationship: PATCH /tenantRelationships/delegatedAdminRelationships/{id} with If-Match (etag)
  - List relationships: GET /tenantRelationships/delegatedAdminRelationships [3](https://www.reddit.com/r/PowerShell/comments/zkzz7p/pim_role_ga_eligibility_with_graph_api/)
  - Global Administrator role template ID: 62e90394-69f5-4237-9190-012177145e10 [4](https://community.dynamics.com/forums/thread/details/?threadid=7eb95dfc-4a20-f011-998a-7c1e5266971b)

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  # If set, only produces the report and does not change anything.
  [switch] $ReportOnly,

  # If set, tries to remove Global Administrator role from eligible relationships (status must be active).
  [switch] $RemoveGlobalAdminIfPresent,

  # Auto-extend duration to set (ISO 8601 duration). P180D = 6 months.
  [ValidateSet("P180D")]
  [string] $TargetAutoExtendDuration = "P180D",

  # Where to write the CSV report
  [string] $ReportPath = (Join-Path $PWD ("GDAP_AutoExtend_Audit_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

  # How many attempts to poll after removing GA (because it can be long-running)
  [int] $MaxPollAttempts = 30,

  # Seconds to wait between polls
  [int] $PollIntervalSeconds = 10
)

# -----------------------------
# Constants
# -----------------------------
# Global Administrator role template ID [4](https://community.dynamics.com/forums/thread/details/?threadid=7eb95dfc-4a20-f011-998a-7c1e5266971b)
$GlobalAdminRoleDefinitionId = "62e90394-69f5-4237-9190-012177145e10"

# Relationship list endpoint [3](https://www.reddit.com/r/PowerShell/comments/zkzz7p/pim_role_ga_eligibility_with_graph_api/)
$ListUriBase = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships?`$top=300"

# -----------------------------
# Helpers
# -----------------------------
function Ensure-GraphConnection {
  if (-not (Get-Module Microsoft.Graph -ListAvailable)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force | Out-Null
  }
  Import-Module Microsoft.Graph | Out-Null

  Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
  Connect-MgGraph -Scopes "DelegatedAdminRelationship.ReadWrite.All" -NoWelcome | Out-Null
}

function Get-AllDelegatedAdminRelationships {
  # Uses paging via @odata.nextLink [3](https://www.reddit.com/r/PowerShell/comments/zkzz7p/pim_role_ga_eligibility_with_graph_api/)
  $all = @()
  $uri = $ListUriBase

  while ($uri) {
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
    if ($resp.value) { $all += $resp.value }
    $uri = $resp.'@odata.nextLink'
  }
  return $all
}

function Get-RelationshipById([string] $RelationshipId) {
  return Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$RelationshipId"
}

function Get-Etag($Relationship) {
  # ETag exists on delegatedAdminRelationship objects [3](https://www.reddit.com/r/PowerShell/comments/zkzz7p/pim_role_ga_eligibility_with_graph_api/)
  return $Relationship.'@odata.etag'
}

function Has-GlobalAdminRole($Relationship) {
  $roles = $Relationship.accessDetails.unifiedRoles
  if (-not $roles) { return $false }
  return $roles.roleDefinitionId -contains $GlobalAdminRoleDefinitionId
}

function Set-AutoExtendDuration([string] $RelationshipId, [string] $Etag, [string] $Duration) {
  # PATCH requires If-Match header with the last known ETag
  $headers = @{
    "If-Match"     = $Etag
    "Content-Type" = "application/json"
  }

  $body = @{ autoExtendDuration = $Duration } | ConvertTo-Json -Depth 10

  Invoke-MgGraphRequest -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$RelationshipId" `
    -Headers $headers `
    -Body $body | Out-Null
}

function Remove-GlobalAdminRole([object] $Relationship, [string] $Etag) {
  # Only removable when status is active; operation can be long-running (202 Accepted possible).
  $relationshipId = $Relationship.id

  $currentRoles = @($Relationship.accessDetails.unifiedRoles)
  $newRoles = $currentRoles | Where-Object { $_.roleDefinitionId -ne $GlobalAdminRoleDefinitionId }

  $headers = @{
    "If-Match"     = $Etag
    "Content-Type" = "application/json"
  }

  $bodyObj = @{
    accessDetails = @{
      unifiedRoles = $newRoles
    }
  }
  $body = $bodyObj | ConvertTo-Json -Depth 10

  Invoke-MgGraphRequest -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$relationshipId" `
    -Headers $headers `
    -Body $body | Out-Null
}

function Wait-UntilGlobalAdminRemoved([string] $RelationshipId) {
  for ($i = 1; $i -le $MaxPollAttempts; $i++) {
    $r = Get-RelationshipById -RelationshipId $RelationshipId
    if (-not (Has-GlobalAdminRole $r)) {
      return $true
    }
    Start-Sleep -Seconds $PollIntervalSeconds
  }
  return $false
}

# -----------------------------
# Main
# -----------------------------
Ensure-GraphConnection

Write-Host "Fetching delegated admin (GDAP) relationships..." -ForegroundColor Cyan
$relationships = Get-AllDelegatedAdminRelationships

# Audit model
$audit = foreach ($r in $relationships) {
  [pscustomobject]@{
    RelationshipId     = $r.id
    RelationshipName   = $r.displayName
    CustomerTenantId   = $r.customer.tenantId
    CustomerName       = $r.customer.displayName
    Status             = $r.status
    AutoExtendDuration = $r.autoExtendDuration
    AutoExtendEnabled  = ($r.autoExtendDuration -eq $TargetAutoExtendDuration)
    HasGlobalAdminRole = (Has-GlobalAdminRole $r)
  }
}

# Export audit CSV
$audit | Export-Csv -NoTypeInformation -Path $ReportPath
Write-Host "Audit exported to: $ReportPath" -ForegroundColor Green

if ($ReportOnly) {
  Write-Host "ReportOnly specified. No changes were made." -ForegroundColor Yellow
  Disconnect-MgGraph | Out-Null
  return
}

# Targets: relationships where auto-extend is not enabled and status is created or active
$targets = $audit | Where-Object {
  (-not $_.AutoExtendEnabled) -and ($_.Status -in @("created","active"))
}

Write-Host "Found $($targets.Count) relationships that are 'created/active' and not set to auto-extend ($TargetAutoExtendDuration)." -ForegroundColor Yellow

foreach ($t in $targets) {
  Write-Host "`nProcessing: $($t.RelationshipName) / $($t.CustomerName) [$($t.Status)]" -ForegroundColor Cyan

  try {
    # Always re-fetch latest relationship and ETag before patching to satisfy If-Match requirement [3](https://www.reddit.com/r/PowerShell/comments/zkzz7p/pim_role_ga_eligibility_with_graph_api/)
    $fresh = Get-RelationshipById -RelationshipId $t.RelationshipId
    $freshEtag = Get-Etag $fresh

    # If Global Admin role is present, auto-extend is not allowed; remove GA if user enabled this option. [1](https://github.com/microsoftgraph/microsoft-graph-docs-contrib/blob/main/api-reference/beta/api/delegatedadminrelationship-update.md)
    if ($RemoveGlobalAdminIfPresent -and (Has-GlobalAdminRole $fresh)) {
      if ($fresh.status -ne "active") {
        # Graph only allows removing GA when status is active
        Write-Host " - Global Admin present but status is '$($fresh.status)'; cannot remove GA unless status is active." -ForegroundColor DarkYellow
      }
      else {
        if ($PSCmdlet.ShouldProcess("$($t.RelationshipName) ($($t.CustomerName))","Remove Global Administrator role from relationship")) {
          Write-Host " - Removing Global Administrator role from relationship accessDetails..." -ForegroundColor Yellow
          Remove-GlobalAdminRole -Relationship $fresh -Etag $freshEtag

          # Removal can be long-running; poll until the role is gone (or we hit attempts)
          Write-Host " - Waiting for Global Admin removal to complete (polling)..." -ForegroundColor Yellow
          $removed = Wait-UntilGlobalAdminRemoved -RelationshipId $t.RelationshipId
          if (-not $removed) {
            Write-Host " - Warning: Global Admin still present after polling. Auto-extend may still fail." -ForegroundColor DarkYellow
          }
        }
      }
    }

    # Re-fetch again for latest ETag and then enable auto-extend (P180D)
    $after = Get-RelationshipById -RelationshipId $t.RelationshipId
    $afterEtag = Get-Etag $after

    if ($PSCmdlet.ShouldProcess("$($t.RelationshipName) ($($t.CustomerName))","Set autoExtendDuration = $TargetAutoExtendDuration")) {
      Write-Host " - Enabling auto-extend (autoExtendDuration = $TargetAutoExtendDuration)..." -ForegroundColor Yellow
      Set-AutoExtendDuration -RelationshipId $t.RelationshipId -Etag $afterEtag -Duration $TargetAutoExtendDuration
      Write-Host " - Done." -ForegroundColor Green
    }
  }
  catch {
    Write-Host " - FAILED: $($_.Exception.Message)" -ForegroundColor Red
  }
}

Disconnect-MgGraph | Out-Null
Write-Host "`nCompleted." -ForegroundColor Green
```

<br><br>

## Utilities & Supporting Toolkits

### SharePoint - PowerShell Modules required in this repo

```powershell
Install-Module ExchangeOnlineManagement
Install-Module AzureAD
Install-Module MSOnline
Install-Module Microsoft.Online.SharePoint.PowerShell
Install-Module ImportExcel
Install-Module MSOnline
Install-Module Microsoft.Graph
Install-Module Microsoft.Graph.Beta 
Install-Module PnP.PowerShell
Install-Module ImportExcel
```

<br>

### SharePoint - REST API to get Site Page ID

> [!TIP]
> You can use any tool like cURL, wget, or even in browser.

```
https://[TENANT_NAME].sharepoint.com/sites/[SITE_NAME]/_api/web/lists/getbytitle('Site Pages')/id
```

Sample output/result:

```
This XML file does not appear to have any style information associated with it. The document tree is shown below.
<d:Id xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns:georss="http://www.georss.org/georss" xmlns:gml="http://www.opengis.net/gml" m:type="Edm.Guid">[GUID]</d:Id>
```

<br>

### `Invoke-MgGraphRequest` command error

In some cases, the `Invoke-MgGraphRequest` will prompt error as below:

```powershell
Invoke-MgGraphRequest: Could not load type 'Microsoft.Graph.Authentication.AzureIdentityAccessTokenProvider' from assembly 'Microsoft.Graph.Core, Version=1.25.1.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35'.
```

This is because the installed `Microsoft.Graph` modules are conflicted, either you have beta or difference version installed.

To temporary resolve it at the time being, you can start the PowerShell session with `-NoProfile`:

```powershell
pwsh -NoProfile
```

To permanent fix it, you either find out which version is conflicted **OR** remove all and install the correct and specific version of `Microsoft.Graph`.

<br>

### Use `-DisableNameChecking` when import `Microsoft.Online.Sharepoint.PowerShell`

The reason you use -DisableNameChecking when importing the Microsoft.Online.SharePoint.PowerShell module is to suppress warning messages about unapproved verbs.

In PowerShell, Microsoft has a strict list of "Approved Verbs" (like Get, Set, New, Remove) to keep things consistent. If a module includes cmdlets that use non-standard verbs (e.g., Upgrade or Connect), PowerShell will throw a yellow warning every time you import it, notifying you that the module contains "unapproved verbs".

Why this matters for SharePoint:

- **Legacy Naming**: Many SharePoint Online cmdlets were created before certain naming conventions were strictly enforced or they use specific service-oriented verbs that aren't on the official "approved" list.
- **Cleaner Scripts**: Using -DisableNameChecking makes your scripts look cleaner by preventing these warnings from cluttering the console or your logs.
- **No Impact on Functionality**: This parameter only hides the warning. It doesn't change how the cmdlets work or affect their performance.





<details>

<summary><span style="padding-left:10px;">Click here for "Sample 1"</span></summary>

```

```

</details>


