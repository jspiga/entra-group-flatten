<#
.SYNOPSIS
    Version 2: Registers Microsoft Graph change notification subscriptions for a source group
    and all of its nested sub-groups, so that membership changes trigger the webhook endpoint.

.DESCRIPTION
    Creates (or renews) a Graph change notification subscription for each group that is part of
    the source group hierarchy. When any of these groups change, Microsoft Graph will POST a
    notification to the configured webhook URL.

    The webhook receiver (Start-WebhookListener.ps1) then calls Sync-FlattenedGroup.ps1 to
    re-flatten the group and sync the target group.

    Subscriptions are valid for a maximum of 3 days for group resources. This script should be
    run on a schedule (e.g. every 2 days) to renew them before they expire.

.PARAMETER SourceGroup
    Display name of the source group whose full hierarchy should be subscribed.

.PARAMETER TenantId
    Azure AD tenant ID. Overrides config.json if provided.

.PARAMETER ClientId
    App Registration client ID. Overrides config.json if provided.

.PARAMETER ClientSecret
    App Registration client secret. Overrides config.json if provided.

.PARAMETER NotificationUrl
    The HTTPS endpoint that will receive change notifications. Overrides config.json if provided.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER SubscriptionStorePath
    Path to a JSON file where active subscription IDs are persisted for renewal tracking.
    Defaults to ./state/subscriptions.json.

.EXAMPLE
    .\Register-ChangeNotification.ps1 -SourceGroup "All-Engineering" `
        -NotificationUrl "https://myserver.example.com/notify"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourceGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$NotificationUrl,

    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),
    [string]$SubscriptionStorePath = (Join-Path $PSScriptRoot "state/subscriptions.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $colour = switch ($Level) {
        "INFO"    { "Cyan"   }
        "SUCCESS" { "Green"  }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red"    }
        default   { "White"  }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $colour
}

# ── Load libraries ────────────────────────────────────────────────────────────

. (Join-Path $PSScriptRoot "lib/GraphAuth.ps1")
. (Join-Path $PSScriptRoot "lib/GraphGroups.ps1")

# ── Load config ───────────────────────────────────────────────────────────────

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading config from '$ConfigPath'"
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable
}

if (-not $TenantId)       { $TenantId        = $config["tenantId"]              }
if (-not $ClientId)       { $ClientId        = $config["clientId"]              }
if (-not $ClientSecret)   { $ClientSecret    = $config["clientSecret"]          }
if (-not $NotificationUrl){ $NotificationUrl = $config["webhookNotificationUrl"] }

foreach ($required in @("TenantId","ClientId","ClientSecret","NotificationUrl")) {
    if (-not (Get-Variable $required -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required value: '$required'. Provide it as a parameter or in config.json."
    }
}

# ── Authenticate ─────────────────────────────────────────────────────────────

Write-Log "Authenticating to Microsoft Graph (tenant: $TenantId)"
$headers = Get-GraphHeaders -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Log "Authentication successful." "SUCCESS"

# ── Resolve source group ──────────────────────────────────────────────────────

Write-Log "Resolving source group: '$SourceGroup'"
$sourceGroupObj = Get-GroupByName -Headers $headers -DisplayName $SourceGroup
if (-not $sourceGroupObj) {
    throw "Source group '$SourceGroup' was not found in Entra ID."
}
Write-Log "Source group: '$($sourceGroupObj.displayName)' (ID: $($sourceGroupObj.id))" "SUCCESS"

# ── Discover all nested group IDs ────────────────────────────────────────────

Write-Log "Discovering nested groups in '$SourceGroup' hierarchy..."
$groupIds = Get-NestedGroupIds -Headers $headers -GroupId $sourceGroupObj.id
Write-Log "Found $($groupIds.Count) group(s) to subscribe (including root)." "SUCCESS"

# ── Load existing subscriptions ───────────────────────────────────────────────

$subscriptionStore = @{}
if (Test-Path $SubscriptionStorePath) {
    Write-Log "Loading existing subscription records from '$SubscriptionStorePath'"
    $subscriptionStore = Get-Content $SubscriptionStorePath -Raw | ConvertFrom-Json -AsHashtable
}

# ── Create or renew subscriptions ─────────────────────────────────────────────

# Graph allows subscriptions on group resources for up to 3 days (4230 minutes max)
$expiryUtc = [datetime]::UtcNow.AddDays(2).ToString("yyyy-MM-ddTHH:mm:ssZ")
$graphSubUrl = "https://graph.microsoft.com/v1.0/subscriptions"

$newStore = @{}

foreach ($groupId in $groupIds) {
    $resource = "groups/$groupId/members"
    $existingSubId = $subscriptionStore[$groupId]

    if ($existingSubId) {
        # Attempt renewal (PATCH)
        Write-Log "Renewing subscription '$existingSubId' for group '$groupId'..."
        try {
            $renewBody = @{ expirationDateTime = $expiryUtc }
            Invoke-GraphRequest -Headers $headers -Method PATCH `
                -Uri "$graphSubUrl/$existingSubId" -Body $renewBody | Out-Null
            Write-Log "Subscription '$existingSubId' renewed (expires $expiryUtc)." "SUCCESS"
            $newStore[$groupId] = $existingSubId
            continue
        }
        catch {
            Write-Log "Failed to renew subscription '$existingSubId' — will create a new one. ($_)" "WARN"
        }
    }

    # Create new subscription
    Write-Log "Creating new subscription for group '$groupId' (resource: $resource)..."
    $subBody = @{
        changeType             = "updated"
        notificationUrl        = $NotificationUrl
        resource               = $resource
        expirationDateTime     = $expiryUtc
        clientState            = "entra-group-flatten-$groupId"
    }

    try {
        $sub = Invoke-GraphRequest -Headers $headers -Method POST -Uri $graphSubUrl -Body $subBody
        Write-Log "Subscription created: '$($sub.id)' (expires $($sub.expirationDateTime))" "SUCCESS"
        $newStore[$groupId] = $sub.id
    }
    catch {
        Write-Log "Failed to create subscription for group '$groupId': $_" "ERROR"
    }
}

# ── Persist updated subscription store ───────────────────────────────────────

$storeDir = Split-Path $SubscriptionStorePath -Parent
if ($storeDir -and -not (Test-Path $storeDir)) {
    New-Item -ItemType Directory -Path $storeDir -Force | Out-Null
}
$newStore | ConvertTo-Json -Depth 5 | Set-Content -Path $SubscriptionStorePath -Encoding UTF8
Write-Log "Subscription store saved to '$SubscriptionStorePath'" "SUCCESS"
Write-Log "Done. $($newStore.Count) active subscription(s) registered." "SUCCESS"
