Add-Supplier.ps1
<#
.SYNOPSIS
    One-click supplier onboarding for a single Microsoft 365 tenant.

.DESCRIPTION
    - Creates an M365 group (Unified)
    - Creates a Team (and associated SharePoint site)
    - Invites supplier users as B2B guests
    - Adds them to the group
    - Sets expiry metadata (in description) and optionally hooks into a lifecycle policy

.NOTES
    Requires:
    - Microsoft.Graph PowerShell SDK
    - App with sufficient permissions or admin delegated sign-in
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SupplierName,

    [Parameter(Mandatory = $true)]
    [string]$SupplierDomain,

    [Parameter(Mandatory = $true)]
    [string[]]$SupplierUsers,   # list of email addresses

    [int]$ExpiryDays = 90
)

# 1. Connect to Microsoft Graph
# For delegated (interactive) – good for testing:
# Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All","Directory.ReadWrite.All"

# For app-only (recommended for automation), use:
# Connect-MgGraph -ClientId "YOUR_APP_ID" -TenantId "YOUR_TENANT_ID" -CertificateThumbprint "THUMBPRINT"

Import-Module Microsoft.Graph

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All","Directory.ReadWrite.All"
Select-MgProfile -Name "v1.0"

# Helper: sanitize supplier name for mail nickname
function New-MailNicknameFromName {
    param([string]$name)
    $nickname = $name.ToLowerInvariant() -replace '[^a-z0-9]', ''
    if ($nickname.Length -gt 40) {
        $nickname = $nickname.Substring(0, 40)
    }
    return $nickname
}

$nickname = New-MailNicknameFromName -name $SupplierName
$expiryDate = (Get-Date).AddDays($ExpiryDays).ToString("yyyy-MM-dd")

Write-Host "Creating M365 group for supplier '$SupplierName'..." -ForegroundColor Cyan

# 2. Create Unified Group (M365 group)
$groupBody = @{
    displayName     = "SUPPLIER - $SupplierName"
    mailNickname    = $nickname
    mailEnabled     = $true
    securityEnabled = $false
    groupTypes      = @("Unified")
    description     = "Supplier: $SupplierName ($SupplierDomain) | Expires on $expiryDate"
}

$group = New-MgGroup -BodyParameter $groupBody
Write-Host "Created group: $($group.DisplayName) ($($group.Id))"

# 3. Create a Team on top of the group (auto-creates SharePoint site)
# This call is asynchronous; in production you may want to poll until complete.
Write-Host "Creating Team for the supplier group..." -ForegroundColor Cyan

$teamBody = @{
    "memberSettings" = @{
        "allowCreateUpdateChannels" = $true
    }
    "messagingSettings" = @{
        "allowUserEditMessages" = $true
        "allowUserDeleteMessages" = $true
    }
    "funSettings" = @{
        "allowGiphy" = $true
    }
    "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
}

New-MgTeam -GroupId $group.Id -BodyParameter $teamBody

# Retrieve primary site URL (optional, for invite redirect)
Write-Host "Fetching associated SharePoint site (may take a few seconds to provision)..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

$sites = Get-MgGroupSite -GroupId $group.Id -ErrorAction SilentlyContinue
$siteUrl = $null
if ($sites) {
    $siteUrl = $sites[0].WebUrl
    Write-Host "Supplier SharePoint/Teams site URL: $siteUrl"
} else {
    Write-Host "Could not retrieve site yet. It may still be provisioning." -ForegroundColor Yellow
}

# 4. Invite each supplier user as a guest and add to group
foreach ($email in $SupplierUsers) {
    Write-Host "Inviting supplier user: $email" -ForegroundColor Cyan

    $inviteBody = @{
        invitedUserEmailAddress = $email
        inviteRedirectUrl       = $siteUrl # will send them here after redemption
        sendInvitationMessage   = $true
        invitedUserMessageInfo  = @{
            customizedMessageBody = "You've been granted secure access as a supplier to $SupplierName."
        }
    }

    $invitation = New-MgInvitation -BodyParameter $inviteBody

    $guestUserId = $invitation.invitedUser.Id
    if (-not $guestUserId) {
        Write-Warning "Could not retrieve guest user ID for $email. Skipping group membership."
        continue
    }

    # Add guest to group
    $directoryObjectBody = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$guestUserId"
    }

    New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter $directoryObjectBody
    Write-Host "Added $email as member of group $($group.DisplayName)."
}

Write-Host "Supplier onboarding complete!" -ForegroundColor Green
Write-Host "Group: $($group.DisplayName)"
Write-Host "Site:  $siteUrl"
Write-Host "Expires (metadata): $expiryDate"
