# SharePoint, M365, and EntraID Automation Toolkits

```markdown
> Austin.Lai |
> -----------| April 25th, 2026
> -----------| Updated on April 26th, 2026
```

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
    - [M365 or EntraID Toolkits](#m365-or-entraid-toolkits)
        - [Automation to ensure all tenant enabled GDAP Auto-Extend](#automation-to-ensure-all-tenant-enabled-gdap-auto-extend)
    - [Utilities & Supporting Toolkits](#utilities--supporting-toolkits)
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

<br>

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

## M365 or EntraID Toolkits

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




<br>

## Utilities & Supporting Toolkits

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

To permanent fix it, you either find out which version is conflicted or remove all and install the correct and specific version of `Microsoft.Graph`.

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


