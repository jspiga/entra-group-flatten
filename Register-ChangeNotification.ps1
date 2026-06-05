<#
.SYNOPSIS
    Registers a single Microsoft Graph change notification subscription covering all groups in
    the tenant, and builds/updates a group registry mapping each monitored group ID to its
    source group name.

.DESCRIPTION
    Replaces the old per-group subscription model with a single tenant-wide subscription on
    /groups, which scales to any number of source groups without hitting Graph subscription limits.

    On each run this script:
      1. Resolves each SourceGroup name to an Entra ID object and discovers its full nested
         group hierarchy.
      2. Writes/updates state/group-registry.json, mapping every group ID in every hierarchy
         to the source group display name it belongs to.
      3. Creates or renews a single Graph change-notification subscription on the /groups
         resource. The subscription ID is stored in state/subscriptions.json.

    Subscriptions on /groups are valid for up to 3 days. Schedule this script to run every
    2 days to keep the subscription alive.

    The webhook receiver (Start-WebhookListener.ps1) loads group-registry.json to determine
    which source group is affected by each incoming notification and triggers the appropriate
    Invoke-DeltaSync.ps1 job.

.PARAMETER SourceGroup
    One or more source group display names whose hierarchies should be monitored.
    Accepts an array: -SourceGroup "Group A","Group B"

.PARAMETER TargetGroup
    One or more corresponding target group display names (must match SourceGroup order).
    Used to record the source-to-target mapping in the registry.

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
    Path to the JSON file where the active subscription ID is stored.
    Defaults to ./state/subscriptions.json.

.PARAMETER RegistryPath
    Path to the JSON group registry file.
    Defaults to ./state/group-registry.json.

.EXAMPLE
    # Single group
    .\Register-ChangeNotification.ps1 `
        -SourceGroup "All-Engineering" `
        -TargetGroup "EngineeringFlat" `
        -NotificationUrl "https://myserver.example.com/notify"

.EXAMPLE
    # Multiple groups
    .\Register-ChangeNotification.ps1 `
        -SourceGroup "All-Engineering","All-Finance" `
        -TargetGroup "EngineeringFlat","FinanceFlat" `
        -NotificationUrl "https://myserver.example.com/notify"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$SourceGroup,

    [Parameter(Mandatory)]
    [string[]]$TargetGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,
    [string]$NotificationUrl,

    [string]$ConfigPath            = "",
    [string]$SubscriptionStorePath = "",
    [string]$RegistryPath          = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ConfigPath)            { $ConfigPath            = Join-Path $PSScriptRoot "config.json"               }
if (-not $SubscriptionStorePath) { $SubscriptionStorePath = Join-Path $PSScriptRoot "state/subscriptions.json"  }
if (-not $RegistryPath)          { $RegistryPath          = Join-Path $PSScriptRoot "state/group-registry.json" }

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

# ── Validate inputs ───────────────────────────────────────────────────────────

if ($SourceGroup.Count -ne $TargetGroup.Count) {
    throw "SourceGroup and TargetGroup arrays must have the same number of entries."
}

# ── Load libraries ────────────────────────────────────────────────────────────

. (Join-Path $PSScriptRoot "lib/GraphAuth.ps1")
. (Join-Path $PSScriptRoot "lib/GraphGroups.ps1")

# ── Load config ───────────────────────────────────────────────────────────────

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading config from '$ConfigPath'"
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ForEach-Object {
        $h = @{}; $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h
    }
}

if (-not $TenantId)        { $TenantId        = $config["tenantId"]              }
if (-not $ClientId)        { $ClientId        = $config["clientId"]              }
if (-not $ClientSecret)    { $ClientSecret    = $config["clientSecret"]          }
if (-not $NotificationUrl) { $NotificationUrl = $config["webhookNotificationUrl"] }

foreach ($required in @("TenantId","ClientId","ClientSecret","NotificationUrl")) {
    if (-not (Get-Variable $required -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required value: '$required'. Provide it as a parameter or in config.json."
    }
}

# ── Authenticate ──────────────────────────────────────────────────────────────

Write-Log "Authenticating to Microsoft Graph (tenant: $TenantId)"
$headers = Get-GraphHeaders -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Log "Authentication successful." "SUCCESS"

# ── Ensure state directory exists ─────────────────────────────────────────────

$stateDir = Split-Path $RegistryPath -Parent
if ($stateDir -and -not (Test-Path $stateDir)) {
    New-Item -ItemType Directory -Path $stateDir -Force | Out-Null
}

# ── Load existing registry (preserve any existing entries) ───────────────────

$registry = @{}   # groupId -> @[{ sourceGroup = "..."; targetGroup = "..." }]
$excludedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
if (Test-Path $RegistryPath) {
    Write-Log "Loading existing group registry from '$RegistryPath'"
    $raw = Get-Content $RegistryPath -Raw | ConvertFrom-Json
    $raw.PSObject.Properties | ForEach-Object {
        if ($_.Name -eq '__excludedTargetGroupIds__') {
            # Restore excluded target group IDs
            foreach ($id in @($_.Value)) { $null = $excludedIds.Add($id) }
        } else {
            $entries = @()
            foreach ($item in @($_.Value)) {
                $e = @{}
                $item.PSObject.Properties | ForEach-Object { $e[$_.Name] = $_.Value }
                $entries += $e
            }
            $registry[$_.Name] = $entries
        }
    }
    Write-Log "Loaded $($registry.Count) existing registry entries ($($excludedIds.Count) excluded target group IDs)."
}

# ── Discover and register all group hierarchies ───────────────────────────────

for ($i = 0; $i -lt $SourceGroup.Count; $i++) {
    $srcName = $SourceGroup[$i]
    $tgtName = $TargetGroup[$i]

    Write-Log "--- Processing source group: '$srcName' -> target: '$tgtName' ---"

    $sourceGroupObj = Get-GroupByName -Headers $headers -DisplayName $srcName
    if (-not $sourceGroupObj) {
        Write-Log "Source group '$srcName' not found in Entra ID -- skipping." "WARN"
        continue
    }
    Write-Log "Resolved '$srcName' to ID: $($sourceGroupObj.id)" "SUCCESS"

    Write-Log "Discovering nested group hierarchy for '$srcName'..."
    $groupIds = Get-NestedGroupIds -Headers $headers -GroupId $sourceGroupObj.id
    Write-Log "Found $($groupIds.Count) group(s) in hierarchy (including root)." "SUCCESS"

    foreach ($gId in $groupIds) {
        # A group ID may belong to multiple source hierarchies (e.g. L2 is both a child of L1
        # and itself a source group). Store an array of mappings per group ID.
        $newEntry = @{ sourceGroup = $srcName; targetGroup = $tgtName }
        if ($registry.ContainsKey($gId)) {
            $existing = @($registry[$gId])
            # Only add if this source group isn't already present
            $alreadyPresent = $existing | Where-Object { $_["sourceGroup"] -eq $srcName }
            if (-not $alreadyPresent) {
                $registry[$gId] = $existing + $newEntry
            }
        } else {
            $registry[$gId] = @($newEntry)
        }
    }
    Write-Log "Registry updated with $($groupIds.Count) entries for '$srcName'."

    # ── Resolve target group ID and add to exclusion list ────────────────────
    # Target groups modified by our sync must be excluded from notification processing
    # to prevent circular dependency loops (our own writes triggering re-syncs).
    $tgtGroupObj = Get-GroupByName -Headers $headers -DisplayName $tgtName
    if ($tgtGroupObj) {
        $null = $excludedIds.Add($tgtGroupObj.id)
        Write-Log "Target group '$tgtName' (ID: $($tgtGroupObj.id)) added to exclusion list."
    } else {
        Write-Log "Target group '$tgtName' not found in Entra -- exclusion skipped (group may not exist yet)." "WARN"
    }
}

# ── Save updated registry ─────────────────────────────────────────────────────

# Embed excluded target group IDs in the registry file under a reserved key
$registryWithExclusions = @{}
$registry.Keys | ForEach-Object { $registryWithExclusions[$_] = $registry[$_] }
$registryWithExclusions['__excludedTargetGroupIds__'] = @($excludedIds)

$registryWithExclusions | ConvertTo-Json -Depth 5 | Set-Content -Path $RegistryPath -Encoding UTF8
Write-Log "Group registry saved to '$RegistryPath' ($($registry.Count) source entries, $($excludedIds.Count) excluded target IDs)." "SUCCESS"

# ── Load existing subscription store ─────────────────────────────────────────

$subStore = @{}
if (Test-Path $SubscriptionStorePath) {
    Write-Log "Loading existing subscription store from '$SubscriptionStorePath'"
    $raw2 = Get-Content $SubscriptionStorePath -Raw | ConvertFrom-Json
    $raw2.PSObject.Properties | ForEach-Object { $subStore[$_.Name] = $_.Value }
}

# ── Create or renew the single tenant-wide subscription ──────────────────────

# Graph allows subscriptions on /groups for up to 3 days (4230 minutes max)
$expiryUtc   = [datetime]::UtcNow.AddDays(2).ToString("yyyy-MM-ddTHH:mm:ssZ")
$graphSubUrl = "https://graph.microsoft.com/v1.0/subscriptions"
$resource    = "groups"
$clientState = "entra-group-flatten-tenant"

$existingSubId = $subStore["tenantSubscriptionId"]

if ($existingSubId) {
    Write-Log "Attempting to renew existing subscription '$existingSubId'..."
    try {
        $renewBody = @{ expirationDateTime = $expiryUtc }
        Invoke-GraphRequest -Headers $headers -Method PATCH `
            -Uri "$graphSubUrl/$existingSubId" -Body $renewBody | Out-Null
        Write-Log "Subscription '$existingSubId' renewed (expires $expiryUtc)." "SUCCESS"
        $subStore["tenantSubscriptionId"] = $existingSubId
        $subStore["expiresAt"]            = $expiryUtc
    }
    catch {
        Write-Log "Failed to renew subscription '$existingSubId' -- will create a new one. ($_)" "WARN"
        $existingSubId = $null
    }
}

if (-not $existingSubId) {
    Write-Log "Creating new tenant-wide subscription on resource '$resource'..."
    $subBody = @{
        changeType         = "updated"
        notificationUrl    = $NotificationUrl
        resource           = $resource
        expirationDateTime = $expiryUtc
        clientState        = $clientState
    }
    try {
        $sub = Invoke-GraphRequest -Headers $headers -Method POST -Uri $graphSubUrl -Body $subBody
        Write-Log "Subscription created: '$($sub.id)' (expires $($sub.expirationDateTime))" "SUCCESS"
        $subStore["tenantSubscriptionId"] = $sub.id
        $subStore["expiresAt"]            = $sub.expirationDateTime
        $subStore["clientState"]          = $clientState
    }
    catch {
        Write-Log "Failed to create subscription: $_" "ERROR"
        throw
    }
}

# ── Persist subscription store ────────────────────────────────────────────────

$subDir = Split-Path $SubscriptionStorePath -Parent
if ($subDir -and -not (Test-Path $subDir)) {
    New-Item -ItemType Directory -Path $subDir -Force | Out-Null
}
$subStore | ConvertTo-Json -Depth 5 | Set-Content -Path $SubscriptionStorePath -Encoding UTF8
Write-Log "Subscription store saved to '$SubscriptionStorePath'." "SUCCESS"
Write-Log "Done. Monitoring $($registry.Count) group(s) via 1 subscription." "SUCCESS"
