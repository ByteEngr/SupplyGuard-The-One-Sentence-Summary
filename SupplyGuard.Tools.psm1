<#
.SYNOPSIS
    SupplyGuard helper module for Microsoft 365 supply-chain scenarios.

.EXPORTS
    Connect-SupplyGuardGraph
    New-SupplyGuardSupplier
    Get-SupplyGuardExternalTenantSummary
#>

# Ensure Microsoft.Graph is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Warning "Microsoft.Graph module not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
}

function Connect-SupplyGuardGraph {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with the required scopes for SupplyGuard scripts.
    .DESCRIPTION
        For MVP we use delegated auth (interactive login).
        Later you can switch to app-only auth for automation.
    #>
    [CmdletBinding()]
    param(
        [string[]]$Scopes = @(
            "User.Read.All",
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
    )

    Import-Module Microsoft.Graph -ErrorAction Stop

    Write-Host "Connecting to Microsoft Graph with scopes: $($Scopes -join ', ')" -ForegroundColor Cyan
    Connect-MgGraph -Scopes $Scopes
    Select-MgProfile -Name "v1.0"
}

function New-SupplyGuardSupplier {
    <#
    .SYNOPSIS
        One-click supplier onboarding for a single tenant.

    .DESCRIPTION
        - Creates an M365 group (Unified) for the supplier
        - Creates a Team (and associated SharePoint site)
        - Invites supplier contacts as B2B guests and adds them to the group
        - Sets expiry metadata on the group description
        - OPTIONAL:
            - Applies a sensitivity label to the SharePoint site
            - Adds the site to a DLP policy scope

        NOTE:
            For sensitivity labels, this function can use the SharePoint Online module.
            For DLP policies, it can use the Purview/Security & Compliance PowerShell session.

    .PARAMETER SupplierName
        Friendly name of the supplier, e.g. "Contoso Logistics".

    .PARAMETER SupplierDomain
        Primary email/domain of the supplier, e.g. "contoso-logistics.com".

    .PARAMETER SupplierUsers
        One or more supplier contact email addresses.

    .PARAMETER ExpiryDays
        Number of days until the supplier workspace should expire (metadata).

    .PARAMETER SharePointAdminUrl
        Your SPO admin URL, e.g. "https://yourtenant-admin.sharepoint.com".
        Required if you want to apply a sensitivity label.

    .PARAMETER SensitivityLabelId
        GUID of the container sensitivity label to apply to the site (optional).

    .PARAMETER DlpPolicyName
        Name of an existing DLP compliance policy to which the site should be added (optional).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SupplierName,

        [Parameter(Mandatory = $true)]
        [string]$SupplierDomain,

        [Parameter(Mandatory = $true)]
        [string[]]$SupplierUsers,

        [int]$ExpiryDays = 90,

        [string]$SharePointAdminUrl,

        [Guid]$SensitivityLabelId,

        [string]$DlpPolicyName
    )

    # Ensure Graph connection
    if (-not (Get-MgContext)) {
        Connect-SupplyGuardGraph
    }

    function New-MailNicknameFromName {
        param([string]$name)
        $nickname = $name.ToLowerInvariant() -replace '[^a-z0-9]', ''
        if ($nickname.Length -gt 40) {
            $nickname = $nickname.Substring(0, 40)
        }
        return $nickname
    }

    $mailNickname = New-MailNicknameFromName -name $SupplierName
    $expiryDate = (Get-Date).AddDays($ExpiryDays).ToString("yyyy-MM-dd")

    Write-Host "Creating M365 group for supplier '$SupplierName'..." -ForegroundColor Cyan

    $groupBody = @{
        displayName     = "SUPPLIER - $SupplierName"
        mailNickname    = $mailNickname
        mailEnabled     = $true
        securityEnabled = $false
        groupTypes      = @("Unified")
        description     = "Supplier: $SupplierName ($SupplierDomain) | Expires on $expiryDate"
    }

    $group = New-MgGroup -BodyParameter $groupBody
    Write-Host "Created group: $($group.DisplayName) ($($group.Id))"

    Write-Host "Creating Team for the supplier group (standard template)..." -ForegroundColor Cyan

    $teamBody = @{
        "memberSettings" = @{
            "allowCreateUpdateChannels" = $true
        }
        "messagingSettings" = @{
            "allowUserEditMessages"   = $true
            "allowUserDeleteMessages" = $true
        }
        "funSettings" = @{
            "allowGiphy" = $true
        }
        "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
    }

    New-MgTeam -GroupId $group.Id -BodyParameter $teamBody

    # Give Teams/SharePoint some time to provision
    Write-Host "Waiting briefly for site provisioning..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10

    $siteUrl = $null
    try {
        $sites = Get-MgGroupSite -GroupId $group.Id -ErrorAction Stop
        if ($sites) {
            $siteUrl = $sites[0].WebUrl
            Write-Host "Supplier SharePoint site URL: $siteUrl" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Could not retrieve associated SharePoint site yet. It may still be provisioning."
    }

    # === OPTIONAL: Apply Sensitivity Label ===
    if ($SensitivityLabelId -and $SharePointAdminUrl -and $siteUrl) {
        Write-Host "Applying sensitivity label $SensitivityLabelId to site $siteUrl..." -ForegroundColor Cyan

        if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
            Write-Warning "SharePoint Online module not found. Install with: Install-Module Microsoft.Online.SharePoint.PowerShell"
        } else {
            Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue

            # Assumes you have SPO admin credentials or modern auth cached.
            Connect-SPOService -Url $SharePointAdminUrl

            # This uses the modern sensitivity label binding for sites.
            Set-SPOSite -Identity $siteUrl -SensitivityLabel $SensitivityLabelId.Guid

            Write-Host "Sensitivity label applied." -ForegroundColor Green
        }
    }

    # === OPTIONAL: Add site to a DLP Policy ===
    if ($DlpPolicyName -and $siteUrl) {
        Write-Host "Adding site to DLP policy '$DlpPolicyName'..." -ForegroundColor Cyan

        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Warning "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement"
        } else {
            Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

            # Connect to security & compliance (Purview) PowerShell
            # In newer tenants use Connect-ExchangeOnline -ConnectionUri for SCC or Purview.
            try {
                Connect-IPPSSession -ErrorAction Stop
            }
            catch {
                Write-Warning "Could not connect to Purview/SCC PowerShell. Ensure you have the right permissions."
            }

            # Simplified example – in reality you may need to manage GUIDs/locations more explicitly.
            $policy = Get-DlpCompliancePolicy -Identity $DlpPolicyName -ErrorAction SilentlyContinue
            if ($policy) {
                # Merge existing SharePoint locations with new site URL
                $existingLocations = @()
                if ($policy.SharePointLocation) {
                    $existingLocations = $policy.SharePointLocation
                }

                if ($existingLocations -contains $siteUrl) {
                    Write-Host "Site already in DLP policy locations." -ForegroundColor Yellow
                } else {
                    $newLocations = $existingLocations + $siteUrl
                    Set-DlpCompliancePolicy -Identity $policy.Identity -SharePointLocation $newLocations
                    Write-Host "Site added to DLP policy scope." -ForegroundColor Green
                }
            } else {
                Write-Warning "DLP policy '$DlpPolicyName' not found."
            }
        }
    }

    # === Invite supplier users as guests and add them to the group ===
    foreach ($email in $SupplierUsers) {
        Write-Host "Inviting supplier user: $email" -ForegroundColor Cyan

        $inviteBody = @{
            invitedUserEmailAddress = $email
            inviteRedirectUrl       = $siteUrl ? $siteUrl : "https://teams.microsoft.com"
            sendInvitationMessage   = $true
            invitedUserMessageInfo  = @{
                customizedMessageBody = "You've been granted secure supplier access for $SupplierName."
            }
        }

        $invitation = New-MgInvitation -BodyParameter $inviteBody
        $guestUserId = $invitation.invitedUser.Id

        if (-not $guestUserId) {
            Write-Warning "No guest user ID returned for $email. Skipping group membership."
            continue
        }

        $memberBody = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$guestUserId"
        }

        New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $memberBody
        Write-Host "Added $email as member of $($group.DisplayName)." -ForegroundColor Green
    }

    Write-Host "Supplier onboarding complete." -ForegroundColor Green
    [PSCustomObject]@{
        GroupId         = $group.Id
        GroupName       = $group.DisplayName
        SiteUrl         = $siteUrl
        ExpiresOn       = $expiryDate
        SupplierName    = $SupplierName
        SupplierDomain  = $SupplierDomain
        UsersInvited    = ($SupplierUsers -join "; ")
    }
}

function Get-SupplyGuardExternalTenantSummary {
    <#
    .SYNOPSIS
        Discover external suppliers/tenants based on guest accounts and export a summary.

    .DESCRIPTION
        - Retrieves all guest users (userType eq 'Guest')
        - Infers external domain from their mail / UPN
        - Groups by domain and counts users
        - Returns objects and optionally exports CSV
    #>
    [CmdletBinding()]
    param(
        [string]$OutputPath
    )

    if (-not (Get-MgContext)) {
        Connect-SupplyGuardGraph -Scopes "User.Read.All"
    }

    Write-Host "Retrieving guest users..." -ForegroundColor Cyan
    $guests = Get-MgUser -Filter "userType eq 'Guest'" -All

    if (-not $guests) {
        Write-Warning "No guest users found."
        return
    }

    function Get-ExternalDomainFromUser {
        param([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User)

        if ($User.Mail) {
            $parts = $User.Mail.Split("@")
            if ($parts.Count -eq 2) { return $parts[1].ToLower() }
        }

        if ($User.UserPrincipalName -like "*#EXT#*") {
            $prefix = $User.UserPrincipalName.Split("#EXT#")[0]
            $parts  = $prefix.Split("_")
            if ($parts.Count -gt 1) {
                return $parts[-1].ToLower()
            }
        }

        return $null
    }

    $rows = @()

    foreach ($g in $guests) {
        $domain = Get-ExternalDomainFromUser -User $g
        if (-not $domain) { continue }

        $rows += [PSCustomObject]@{
            ExternalDomain = $domain
            UserPrincipal  = $g.UserPrincipalName
            Mail           = $g.Mail
        }
    }

    if (-not $rows) {
        Write-Warning "No external domains could be inferred from guest users."
        return
    }

    $summary = $rows |
        Group-Object -Property ExternalDomain |
        ForEach-Object {
            [PSCustomObject]@{
                ExternalDomain = $_.Name
                UserCount      = $_.Count
                SampleUsers    = ($_.Group | Select-Object -First 3 -ExpandProperty UserPrincipal) -join "; "
            }
        } |
        Sort-Object -Property UserCount -Descending

    if ($OutputPath) {
        $summary | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported summary to $OutputPath" -ForegroundColor Green
    }

    return $summary
}

Export-ModuleMember -Function `
    Connect-SupplyGuardGraph, `
    New-SupplyGuardSupplier, `
    Get-SupplyGuardExternalTenantSummary
