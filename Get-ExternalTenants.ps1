<#
.SYNOPSIS
    Discover external suppliers/tenants based on guest accounts and export to CSV.

.DESCRIPTION
    - Fetches all guest users (userType eq 'Guest')
    - Extracts the original email domain from their UPN or mail
    - Aggregates by domain and counts users
    - Exports to CSV: ExternalDomain, UserCount, SampleUsers

.NOTES
    For real tenant IDs (GUIDs) you can later enrich using sign-in logs (Get-MgAuditLogSignIn)
#>

param(
    [string]$OutputPath = ".\ExternalTenants.csv"
)

Import-Module Microsoft.Graph

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All"
Select-MgProfile -Name "v1.0"

Write-Host "Retrieving guest users..." -ForegroundColor Cyan
$guests = Get-MgUser -Filter "userType eq 'Guest'" -All

if (-not $guests) {
    Write-Host "No guest users found." -ForegroundColor Yellow
    return
}

function Get-ExternalDomainFromUser {
    param([Microsoft.Graph.PowerShell.Models.IMicrosoftGraphUser]$User)

    # Prefer mail attribute if present (often original email)
    if ($User.Mail) {
        $parts = $User.Mail.Split("@")
        if ($parts.Count -eq 2) { return $parts[1].ToLower() }
    }

    # Fallback: parse UPN like user_domain.com#EXT#@yourtenant.onmicrosoft.com
    if ($User.UserPrincipalName -like "*#EXT#*") {
        $upnParts = $User.UserPrincipalName.Split("#EXT#")[0].Split("_")
        if ($upnParts.Count -gt 1) {
            # Last element often approximates domain (e.g. supplier.com)
            return $upnParts[-1].ToLower()
        }
    }

    return $null
}

$results = @()

foreach ($g in $guests) {
    $domain = Get-ExternalDomainFromUser -User $g
    if ($domain) {
        $results += [PSCustomObject]@{
            ExternalDomain = $domain
            UserId         = $g.Id
            UserPrincipal  = $g.UserPrincipalName
            Mail           = $g.Mail
        }
    }
}

if (-not $results) {
    Write-Host "No external domains could be inferred from guest users." -ForegroundColor Yellow
    return
}

# Aggregate by ExternalDomain
$agg = $results |
    Group-Object -Property ExternalDomain |
    ForEach-Object {
        [PSCustomObject]@{
            ExternalDomain = $_.Name
            UserCount      = $_.Count
            SampleUsers    = ($_.Group | Select-Object -First 3 -ExpandProperty UserPrincipal) -join "; "
        }
    } |
    Sort-Object -Property UserCount -Descending

Write-Host "Found $($agg.Count) external domains." -ForegroundColor Green
$agg | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Exported summary to $OutputPath"

# OPTIONAL (more advanced): enrich with sign-in logs to get tenant IDs
# Connect-MgGraph -Scopes "AuditLog.Read.All"
# $signins = Get-MgAuditLogSignIn -Filter "userType eq 'Guest'" -All
# Then group by ResourceTenantId from $signins and join with $agg by userPrincipal/mail
