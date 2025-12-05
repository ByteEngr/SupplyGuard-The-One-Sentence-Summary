<#
Creates an external supplier identity in under 60 seconds.
#>

param(
  [string]$SupplierName,
  [string]$Email
)

. .\Connect-Graph.ps1

Write-Host "Creating external identity for $SupplierName..." -ForegroundColor Yellow

$invite = New-MgInvitation -InvitedUserEmailAddress $Email -InviteRedirectUrl "https://myapps.microsoft.com" -SendInvitationMessage:$true

Write-Host "Created external user: $($invite.InvitedUser.Id)" -ForegroundColor Green
