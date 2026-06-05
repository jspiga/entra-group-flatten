<#
.SYNOPSIS
    Registers the HTTP URL ACL required for Start-WebhookListener.ps1 to bind on all
    network interfaces without requiring Administrator privileges on every run.

.DESCRIPTION
    Windows HTTP.sys requires a URL ACL reservation before a non-administrator process can
    bind an HttpListener to a public prefix (http://+:port/).

    This script must be run ONCE as Administrator. After that, Start-WebhookListener.ps1
    can be run as a standard user.

.PARAMETER Port
    The port number to register. Must match the -Port used in Start-WebhookListener.ps1.
    Default: 8080.

.PARAMETER UserName
    The Windows user account to grant access to. Defaults to the current user.

.PARAMETER Remove
    If specified, removes the URL ACL reservation instead of adding it.

.EXAMPLE
    # Run as Administrator -- register port 8080 for the current user
    .\Register-UrlAcl.ps1 -Port 8080

.EXAMPLE
    # Register for a specific user
    .\Register-UrlAcl.ps1 -Port 8080 -UserName "DOMAIN\serviceaccount"

.EXAMPLE
    # Remove the reservation
    .\Register-UrlAcl.ps1 -Port 8080 -Remove
#>

[CmdletBinding()]
param(
    [int]$Port       = 8080,
    [string]$UserName = "$env:USERDOMAIN\$env:USERNAME",
    [switch]$Remove
)

# Check for admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as administrator'."
    exit 1
}

$url = "http://+:$Port/"

if ($Remove) {
    Write-Host "Removing URL ACL for $url ..."
    $result = netsh http delete urlacl url=$url 2>&1
    Write-Host $result
} else {
    Write-Host "Registering URL ACL: $url for user '$UserName' ..."
    $result = netsh http add urlacl url=$url user=$UserName 2>&1
    Write-Host $result

    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nSuccess! Start-WebhookListener.ps1 can now bind to http://+:$Port/ without admin rights." -ForegroundColor Green
        Write-Host "Start the listener with:" -ForegroundColor Cyan
        Write-Host "  .\Start-WebhookListener.ps1 -Port $Port" -ForegroundColor Cyan
    } else {
        Write-Host "`nFailed to register URL ACL. Exit code: $LASTEXITCODE" -ForegroundColor Red
    }
}
