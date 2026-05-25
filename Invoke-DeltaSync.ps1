<#
.SYNOPSIS
    Version 3: Optimised group sync using Microsoft Graph delta queries.

.DESCRIPTION
    On the first run against a source group, performs a full transitive member sync (identical
    to Version 1) and stores delta links for every group in the source hierarchy plus the target
    group. On all subsequent runs, only fetches changes since the last run using delta queries,
    significantly reducing API calls and rate-limit exposure.

    Workflow per run:
      1. Load persisted state (delta links + member snapshots) from disk.
      2. For each group of interest (source root + all nested groups):
           a. If no delta link stored → full member fetch + store delta link.
           b. If delta link stored    → delta query → apply changes to tracked snapshot.
      3. Aggregate all tracked member snapshots into a flat de-duplicated set of user IDs.
      4. Load the tracked target membership snapshot.
      5. Diff desired vs tracked target membership.
      6. If a diff exists, apply adds/removes to the live target group and update snapshot.
      7. Persist updated state to disk.

    This script is designed to be called by Start-WebhookListener.ps1 on notification receipt,
    but can also be run manually or on a schedule.

.PARAMETER SourceGroup
    Display name of the source group to flatten.

.PARAMETER TargetGroup
    Display name of the target group to create/update. Prefix is applied from config if needed.

.PARAMETER TenantId
    Azure AD tenant ID. Overrides config.json if provided.

.PARAMETER ClientId
    App Registration client ID. Overrides config.json if provided.

.PARAMETER ClientSecret
    App Registration client secret. Overrides config.json if provided.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER StateFilePath
    Path to the JSON state file. Overrides config.json stateFilePath if provided.
    Defaults to ./state/group-state.json.

.PARAMETER ForceFullSync
    Discard all stored state and perform a full sync from scratch.

.EXAMPLE
    # Normal run -- uses delta links if available
    .\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

.EXAMPLE
    # Force a complete re-sync (discard stored delta state)
    .\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -ForceFullSync
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourceGroup,

    [Parameter(Mandatory)]
    [string]$TargetGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,

    [string]$ConfigPath    = (Join-Path $PSScriptRoot "config.json"),
    [string]$StateFilePath = "",

    [switch]$ForceFullSync
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
. (Join-Path $PSScriptRoot "lib/StateStore.ps1")

# ── Load config ───────────────────────────────────────────────────────────────

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading config from '$ConfigPath'"
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ForEach-Object { $h = @{}; $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h }
}

if (-not $TenantId)     { $TenantId     = $config["tenantId"]     }
if (-not $ClientId)     { $ClientId     = $config["clientId"]     }
if (-not $ClientSecret) { $ClientSecret = $config["clientSecret"] }
$prefix = $config["targetGroupPrefix"]

if (-not $StateFilePath) {
    $StateFilePath = if ($config["stateFilePath"]) {
        # Resolve relative paths against script root
        if ([System.IO.Path]::IsPathRooted($config["stateFilePath"])) {
            $config["stateFilePath"]
        } else {
            Join-Path $PSScriptRoot $config["stateFilePath"]
        }
    } else {
        Join-Path $PSScriptRoot "state/group-state.json"
    }
}

foreach ($required in @("TenantId","ClientId","ClientSecret")) {
    if (-not (Get-Variable $required -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required value: '$required'. Provide it as a parameter or in config.json."
    }
}

# ── Resolve target group name (apply prefix if needed) ───────────────────────

$TargetGroup = Resolve-TargetGroupName -Name $TargetGroup -Prefix $prefix

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

# ── Resolve or create target group ────────────────────────────────────────────

Write-Log "Resolving target group: '$TargetGroup'"
$targetGroupObj  = Get-GroupByName -Headers $headers -DisplayName $TargetGroup
$targetGroupIsNew = $false
if (-not $targetGroupObj) {
    Write-Log "Target group '$TargetGroup' not found -- creating..." "WARN"
    $targetGroupObj  = New-EntraGroup -Headers $headers -DisplayName $TargetGroup `
        -Description "Flattened membership of '$SourceGroup', managed by Invoke-DeltaSync.ps1"
    $targetGroupIsNew = $true
}
Write-Log "Target group: '$($targetGroupObj.displayName)' (ID: $($targetGroupObj.id))" "SUCCESS"

# ── Load state ────────────────────────────────────────────────────────────────

if ($ForceFullSync -and (Test-Path $StateFilePath)) {
    Write-Log "ForceFullSync: removing existing state file '$StateFilePath'" "WARN"
    Remove-Item $StateFilePath -Force
}

$state = Read-State -Path $StateFilePath

# ── Discover groups of interest ───────────────────────────────────────────────

Write-Log "Discovering all groups in source hierarchy..."
$groupIds = Get-NestedGroupIds -Headers $headers -GroupId $sourceGroupObj.id
Write-Log "Found $($groupIds.Count) group(s) of interest (including root)." "SUCCESS"

# ── Process each group via delta query ───────────────────────────────────────

foreach ($groupId in $groupIds) {
    $storedDeltaLink = Get-DeltaLink -State $state -GroupId $groupId

    Write-Log "Processing group '$groupId' ($(if ($storedDeltaLink) { 'incremental delta' } else { 'full sync' }))..."

    if (-not $storedDeltaLink) {
        # ── First run: full member fetch + bootstrap delta link ───────────────
        $deltaResult = Get-GroupMembersDelta -Headers $headers -GroupId $groupId -DeltaLink $null

        # Filter to user objects only (nested groups will be in transitive members, handled separately)
        $userMembers = @($deltaResult.Changes | Where-Object {
            $odataType = $_.PSObject.Properties['@odata.type']
            $removed   = $_.PSObject.Properties['@removed']
            ($odataType -and $odataType.Value -eq '#microsoft.graph.user') -or
            (-not $odataType -and -not $removed)
        })

        Set-TrackedMembers -State $state -GroupId $groupId -Members $userMembers
        Set-DeltaLink      -State $state -GroupId $groupId -DeltaLink $deltaResult.DeltaLink

        Write-Log "Full sync: $($userMembers.Count) user member(s) tracked for group '$groupId'." "SUCCESS"
    }
    else {
        # ── Incremental: apply delta changes to tracked snapshot ──────────────
        $deltaResult = Get-GroupMembersDelta -Headers $headers -GroupId $groupId -DeltaLink $storedDeltaLink

        $currentTracked = Get-TrackedMembers -State $state -GroupId $groupId

        # Build a mutable dictionary keyed by user ID for fast lookup
        $memberMap = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in $currentTracked) {
            if ($m.id) { $memberMap[$m.id] = $m }
        }

        $added   = 0
        $removed = 0

        foreach ($change in $deltaResult.Changes) {
            if (-not $change.id) { continue }

            $changeRemoved   = $change.PSObject.Properties['@removed']
            $changeOdataType = $change.PSObject.Properties['@odata.type']
            if ($changeRemoved) {
                # Member was removed from the group
                if ($memberMap.ContainsKey($change.id)) {
                    $memberMap.Remove($change.id) | Out-Null
                    $removed++
                }
            }
            elseif ((-not $changeOdataType) -or $changeOdataType.Value -eq '#microsoft.graph.user') {
                # Member was added or updated -- upsert
                $memberMap[$change.id] = $change
                $added++
            }
            # Ignore nested group changes -- transitive membership is handled at the aggregate level
        }

        $updatedMembers = @($memberMap.Values)
        Set-TrackedMembers -State $state -GroupId $groupId -Members $updatedMembers
        Set-DeltaLink      -State $state -GroupId $groupId -DeltaLink $deltaResult.DeltaLink

        Write-Log "Delta applied for '$groupId': +$added / -$removed (tracked: $($updatedMembers.Count))" "SUCCESS"
    }
}

# ── Aggregate flattened desired membership ────────────────────────────────────

Write-Log "Aggregating flattened membership across all $($groupIds.Count) group(s)..."

$desiredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($groupId in $groupIds) {
    $members = Get-TrackedMembers -State $state -GroupId $groupId
    foreach ($m in $members) {
        if ($m.id) { $desiredSet.Add($m.id) | Out-Null }
    }
}

Write-Log "Desired flattened membership: $($desiredSet.Count) unique user(s)." "SUCCESS"

# ── Load tracked target membership ────────────────────────────────────────────

$trackedTargetIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($tm in @($state.targetMembers | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.id } })) {
    if ($tm) { $null = $trackedTargetIds.Add($tm) }
}

# ── Compute diff against tracked target state ─────────────────────────────────

$toAdd    = @($desiredSet      | Where-Object { -not $trackedTargetIds.Contains($_) })
$toRemove = @($trackedTargetIds | Where-Object { -not $desiredSet.Contains($_) })

Write-Log "Diff vs tracked target -- Add: $($toAdd.Count)  |  Remove: $($toRemove.Count)"

if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
    Write-Log "No change in flattened membership since last run -- target group update skipped." "SUCCESS"
}
else {
    # Verify actual live target membership before mutating (guard against out-of-band changes)
    # Skip live fetch if the group was just created (empty by definition, avoids 404 propagation delay)
    Write-Log "Fetching live target group membership for final verification..."
    $liveMembers = if ($targetGroupIsNew) { @() } else { Get-GroupMembers -Headers $headers -GroupId $targetGroupObj.id }
    $liveIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($lid in @($liveMembers | Where-Object { $p = $_.PSObject.Properties['@odata.type']; -not $p -or $p.Value -eq '#microsoft.graph.user' } | ForEach-Object { $_.id })) {
        if ($lid) { $null = $liveIds.Add($lid) }
    }

    $liveToAdd    = @($desiredSet | Where-Object { -not $liveIds.Contains($_) })
    $liveToRemove = @($liveIds    | Where-Object { -not $desiredSet.Contains($_) })

    Write-Log "Live diff -- Add: $($liveToAdd.Count)  |  Remove: $($liveToRemove.Count)"

    if ($liveToAdd.Count -gt 0) {
        Write-Log "Adding $($liveToAdd.Count) member(s) to '$TargetGroup'..."
        Add-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $liveToAdd
    }
    if ($liveToRemove.Count -gt 0) {
        Write-Log "Removing $($liveToRemove.Count) member(s) from '$TargetGroup'..."
        Remove-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $liveToRemove
    }

    Write-Log "Target group sync complete." "SUCCESS"

    # Update tracked target membership in state
    $state.targetMembers = @($desiredSet)
}

# ── Persist updated state ─────────────────────────────────────────────────────

Write-State -Path $StateFilePath -State $state
Write-Log "State persisted to '$StateFilePath'." "SUCCESS"
Write-Log "Delta sync run complete." "SUCCESS"
