<#
.SYNOPSIS
    Shared authentication helper for Microsoft Graph API access.
    Acquires and caches an OAuth2 access token using the client credentials flow.

.DESCRIPTION
    Exports:
        Get-GraphToken     - Returns a valid Bearer token string, refreshing if needed.
        Get-GraphHeaders   - Returns an Authorization header hashtable ready for Invoke-RestMethod.
#>

# Module-level token cache
$script:_tokenCache = $null
$script:_tokenExpiry = [datetime]::MinValue

function Get-GraphToken {
    <#
    .SYNOPSIS
        Acquires (or returns a cached) OAuth2 access token for Microsoft Graph.

    .PARAMETER TenantId
        The Azure AD tenant ID (GUID or domain).

    .PARAMETER ClientId
        The App Registration client ID (GUID).

    .PARAMETER ClientSecret
        The client secret string for the App Registration.

    .EXAMPLE
        $token = Get-GraphToken -TenantId $tid -ClientId $cid -ClientSecret $secret
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$ClientSecret
    )

    # Return cached token if still valid (with 60-second buffer)
    if ($script:_tokenCache -and [datetime]::UtcNow -lt $script:_tokenExpiry.AddSeconds(-60)) {
        Write-Verbose "[GraphAuth] Using cached token (expires $($script:_tokenExpiry.ToString('u')))"
        return $script:_tokenCache
    }

    Write-Verbose "[GraphAuth] Acquiring new token for tenant '$TenantId', client '$ClientId'"

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    }
    catch {
        $errMsg = $_.Exception.Message
        throw "Failed to acquire Graph API token: $errMsg"
    }

    $script:_tokenCache = $response.access_token
    $script:_tokenExpiry = [datetime]::UtcNow.AddSeconds($response.expires_in)

    Write-Verbose "[GraphAuth] Token acquired, expires at $($script:_tokenExpiry.ToString('u'))"
    return $script:_tokenCache
}

function Get-GraphHeaders {
    <#
    .SYNOPSIS
        Returns a hashtable of HTTP headers (including Authorization) for Graph API calls.

    .PARAMETER TenantId
        The Azure AD tenant ID.

    .PARAMETER ClientId
        The App Registration client ID.

    .PARAMETER ClientSecret
        The client secret string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$ClientSecret
    )

    $token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    return @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "application/json"
    }
}
