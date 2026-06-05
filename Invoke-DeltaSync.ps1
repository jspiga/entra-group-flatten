<#
.SYNOPSIS
    Option 2: Optimised group sync using Microsoft Graph delta queries.

.DESCRIPTION
    On the first run against a source group, performs a full transitive member sync (identical
    to Option 1) and stores delta links for every group in the source hierarchy plus the target
    group. On all subsequent runs, only fetches changes since the last run using delta queries,
    significantly reducing API calls and rate-limit exposure.

    Supports multiple source/target group pairs in a single run. State is keyed per source group
    so each pair maintains independent delta tracking.

    Workflow per source/target pair per run:
      1. Load persisted state (delta links + member snapshots) from disk.
      2. For each group in the source hierarchy (root + all nested groups):
           a. If no delta link stored -> full member fetch + store delta link.
           b. If delta link stored    -> delta query -> apply changes to tracked snapshot.
      3. Aggregate all tracked member snapshots into a flat de-duplicated set of user IDs.
      4. Diff desired vs tracked target membership.
      5. If a diff exists, apply adds/removes to the live target group and update snapshot.
      6. Persist updated state to disk.

    This script is designed to be called by Start-WebhookListener.ps1 on notification receipt,
    but can also be run manually or on a schedule.

.PARAMETER SourceGroup
    One or more source group display names to flatten. Accepts an array.

.PARAMETER TargetGroup
    One or more corresponding target group display names (must match SourceGroup order).
    Prefix is applied from config if needed.

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
    # Single group -- uses delta links if available
    .\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

.EXAMPLE
    # Multiple groups
    .\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering","All-Finance" -TargetGroup "EngineeringFlat","FinanceFlat"

.EXAMPLE
    # Force a complete re-sync (discard stored delta state)
    .\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -ForceFullSync
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

    [string]$ConfigPath    = "",
    [string]$StateFilePath = "",

    [switch]$ForceFullSync,

    # Where to sync the flattened membership. Defaults to "Entra" if not specified.
    # Use "Atlassian" alone if Entra SCIM provisioning is already active (avoids race conditions).
    # Use "Entra","Atlassian" to sync to both. Can also be set via "syncTo" in config.json.
    [ValidateSet("Entra", "Atlassian")]
    [string[]]$SyncTo = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot "config.json" }

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
. (Join-Path $PSScriptRoot "lib/AtlassianScim.ps1")

# ── Load config ───────────────────────────────────────────────────────────────

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading config from '$ConfigPath'"
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ForEach-Object {
        $h = @{}; $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h
    }
}

if (-not $TenantId)     { $TenantId     = $config["tenantId"]     }
if (-not $ClientId)     { $ClientId     = $config["clientId"]     }
if (-not $ClientSecret) { $ClientSecret = $config["clientSecret"] }
$prefix = $config["targetGroupPrefix"]

if (-not $StateFilePath) {
    $StateFilePath = if ($config["stateFilePath"]) {
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

# ── Resolve SyncTo destinations ───────────────────────────────────────────────

$effectiveSyncTo = if ($SyncTo.Count -gt 0) {
    $SyncTo
} elseif ($config["syncTo"]) {
    @($config["syncTo"])
} else {
    @("Entra")
}

$syncToEntra     = $effectiveSyncTo -contains "Entra"
$syncToAtlassian = $effectiveSyncTo -contains "Atlassian"

Write-Log "Sync destinations: $($effectiveSyncTo -join ', ')"

if ($syncToAtlassian) {
    $atlassianDirectoryId = $config["atlassianDirectoryId"]
    $atlassianApiKey      = $config["atlassianScimApiKey"]
    if (-not $atlassianDirectoryId -or -not $atlassianApiKey) {
        throw "atlassianDirectoryId and atlassianScimApiKey must be set in config.json when syncing to Atlassian."
    }
    Write-Log "Atlassian SCIM enabled (directory: $atlassianDirectoryId)"
    Initialize-ScimClient -DirectoryId $atlassianDirectoryId -ApiKey $atlassianApiKey
}

# ── Validate inputs ───────────────────────────────────────────────────────────

if ($SourceGroup.Count -ne $TargetGroup.Count) {
    throw "SourceGroup and TargetGroup arrays must have the same number of entries (got $($SourceGroup.Count) source(s) and $($TargetGroup.Count) target(s))."
}

# ── Authenticate (once, shared across all group pairs) ────────────────────────

Write-Log "Authenticating to Microsoft Graph (tenant: $TenantId)"
$headers = Get-GraphHeaders -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Log "Authentication successful." "SUCCESS"

# ── Load state (once, shared — keyed by source group name internally) ─────────

if ($ForceFullSync -and (Test-Path $StateFilePath)) {
    Write-Log "ForceFullSync: removing existing state file '$StateFilePath'" "WARN"
    Remove-Item $StateFilePath -Force
}

$state = Read-State -Path $StateFilePath

# ── Process each source/target group pair ─────────────────────────────────────

for ($i = 0; $i -lt $SourceGroup.Count; $i++) {
    $srcName = $SourceGroup[$i]
    $tgtName = Resolve-TargetGroupName -Name $TargetGroup[$i] -Prefix $prefix

    Write-Log "--- [$($i+1)/$($SourceGroup.Count)] '$srcName' -> '$tgtName' ---"

    # ── Resolve source group ──────────────────────────────────────────────────

    Write-Log "Resolving source group: '$srcName'"
    $sourceGroupObj = Get-GroupByName -Headers $headers -DisplayName $srcName
    if (-not $sourceGroupObj) {
        Write-Log "Source group '$srcName' was not found in Entra ID -- skipping." "ERROR"
        continue
    }
    Write-Log "Source group: '$($sourceGroupObj.displayName)' (ID: $($sourceGroupObj.id))" "SUCCESS"

    # ── Resolve or create target group ────────────────────────────────────────

    $targetGroupObj   = $null
    $targetGroupIsNew = $false

    if ($syncToEntra) {
        Write-Log "Resolving target group: '$tgtName'"
        $targetGroupObj = Get-GroupByName -Headers $headers -DisplayName $tgtName
        if (-not $targetGroupObj) {
            Write-Log "Target group '$tgtName' not found -- creating..." "WARN"
            $targetGroupObj  = New-EntraGroup -Headers $headers -DisplayName $tgtName `
                -Description "Flattened membership of '$srcName', managed by Invoke-DeltaSync.ps1"
            $targetGroupIsNew = $true
        }
        Write-Log "Target group: '$($targetGroupObj.displayName)' (ID: $($targetGroupObj.id))" "SUCCESS"
    } else {
        Write-Log "Skipping Entra target group resolution (not syncing to Entra)." "INFO"
    }

    # ── Build per-pair state keys (namespaced by source group name) ───────────

    # State keys are namespaced as "<sourceGroupName>|<groupId>" for deltaLinks/trackedMembers
    # and "<sourceGroupName>|targetMembers" for the target snapshot.
    # This allows multiple source groups to coexist in a single state file.
    $stateKey = $srcName

    # ── Discover groups of interest ───────────────────────────────────────────

    Write-Log "Discovering all groups in source hierarchy for '$srcName'..."
    $groupIds = Get-NestedGroupIds -Headers $headers -GroupId $sourceGroupObj.id
    Write-Log "Found $($groupIds.Count) group(s) of interest (including root)." "SUCCESS"

    # ── Process each group via delta query ────────────────────────────────────

    foreach ($groupId in $groupIds) {
        $namespacedId    = "$stateKey|$groupId"
        $storedDeltaLink = Get-DeltaLink -State $state -GroupId $namespacedId

        Write-Log "Processing group '$groupId' ($(if ($storedDeltaLink) { 'incremental delta' } else { 'full sync' }))..."

        if (-not $storedDeltaLink) {
            # Full member fetch + bootstrap delta link
            $deltaResult = Get-GroupMembersDelta -Headers $headers -GroupId $groupId -DeltaLink $null

            $userMembers = @($deltaResult.Changes | Where-Object {
                $odataType = $_.PSObject.Properties['@odata.type']
                $removed   = $_.PSObject.Properties['@removed']
                ($odataType -and $odataType.Value -eq '#microsoft.graph.user') -or
                (-not $odataType -and -not $removed)
            })

            Set-TrackedMembers -State $state -GroupId $namespacedId -Members $userMembers
            Set-DeltaLink      -State $state -GroupId $namespacedId -DeltaLink $deltaResult.DeltaLink

            Write-Log "Full sync: $($userMembers.Count) user member(s) tracked for group '$groupId'." "SUCCESS"
        }
        else {
            # Incremental: apply delta changes to tracked snapshot
            $deltaResult    = Get-GroupMembersDelta -Headers $headers -GroupId $groupId -DeltaLink $storedDeltaLink
            $currentTracked = Get-TrackedMembers -State $state -GroupId $namespacedId

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
                    if ($memberMap.ContainsKey($change.id)) {
                        $memberMap.Remove($change.id) | Out-Null
                        $removed++
                    }
                }
                elseif ((-not $changeOdataType) -or $changeOdataType.Value -eq '#microsoft.graph.user') {
                    $memberMap[$change.id] = $change
                    $added++
                }
            }

            $updatedMembers = @($memberMap.Values)
            Set-TrackedMembers -State $state -GroupId $namespacedId -Members $updatedMembers
            Set-DeltaLink      -State $state -GroupId $namespacedId -DeltaLink $deltaResult.DeltaLink

            Write-Log "Delta applied for '$groupId': +$added / -$removed (tracked: $($updatedMembers.Count))" "SUCCESS"
        }
    }

    # ── Aggregate flattened desired membership ────────────────────────────────

    Write-Log "Aggregating flattened membership across all $($groupIds.Count) group(s)..."

    $desiredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($groupId in $groupIds) {
        $members = Get-TrackedMembers -State $state -GroupId "$stateKey|$groupId"
        foreach ($m in $members) {
            if ($m -is [string] -and $m) { $null = $desiredSet.Add($m) }
            elseif ($m.id)               { $null = $desiredSet.Add($m.id) }
        }
    }

    Write-Log "Desired flattened membership: $($desiredSet.Count) unique user(s)." "SUCCESS"

    # ── Load tracked target membership (namespaced by source group) ───────────

    $targetMembersKey = "$stateKey|targetMembers"
    $rawTargetMembers = @()
    if ($state.trackedMembers -and $state.trackedMembers.ContainsKey($targetMembersKey)) {
        $rawTargetMembers = @($state.trackedMembers[$targetMembersKey])
    }

    $trackedTargetIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tm in @($rawTargetMembers | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.id } })) {
        if ($tm) { $null = $trackedTargetIds.Add($tm) }
    }

    # ── Compute diff against tracked target state ─────────────────────────────

    $toAdd    = @($desiredSet       | Where-Object { -not $trackedTargetIds.Contains($_) })
    $toRemove = @($trackedTargetIds | Where-Object { -not $desiredSet.Contains($_) })

    Write-Log "Diff vs tracked target -- Add: $($toAdd.Count)  |  Remove: $($toRemove.Count)"

    if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
        Write-Log "No change in flattened membership since last run -- target group update skipped." "SUCCESS"
    }
    else {
        if ($syncToEntra -and $targetGroupObj) {
            # Verify actual live target membership before mutating (guard against out-of-band changes)
            Write-Log "Fetching live target group membership for final verification..."
            $liveMembers = if ($targetGroupIsNew) { @() } else {
                @(Get-GroupMembers -Headers $headers -GroupId $targetGroupObj.id)
            }

            $liveIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($lid in @($liveMembers | Where-Object {
                $p = $_.PSObject.Properties['@odata.type']; -not $p -or $p.Value -eq '#microsoft.graph.user'
            } | ForEach-Object { $_.id })) {
                if ($lid) { $null = $liveIds.Add($lid) }
            }

            $liveToAdd    = @($desiredSet | Where-Object { -not $liveIds.Contains($_) })
            $liveToRemove = @($liveIds    | Where-Object { -not $desiredSet.Contains($_) })

            Write-Log "Live diff -- Add: $($liveToAdd.Count)  |  Remove: $($liveToRemove.Count)"

            if ($liveToAdd.Count -gt 0) {
                Write-Log "[Entra] Adding $($liveToAdd.Count) member(s) to '$tgtName'..."
                Add-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $liveToAdd
            }
            if ($liveToRemove.Count -gt 0) {
                Write-Log "[Entra] Removing $($liveToRemove.Count) member(s) from '$tgtName'..."
                Remove-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $liveToRemove
            }
            Write-Log "[Entra] Target group '$tgtName' sync complete." "SUCCESS"
        } else {
            Write-Log "[Entra] Skipping Entra ID write (not in SyncTo destinations)." "INFO"
        }

        # Update tracked target membership in state (namespaced)
        if (-not $state.trackedMembers) { $state.trackedMembers = @{} }
        $state.trackedMembers[$targetMembersKey] = @($desiredSet)
    }

    # ── Apply changes to Atlassian Cloud (SCIM) ───────────────────────────────

    if ($syncToAtlassian) {
        Write-Log "[Atlassian] Syncing '$tgtName' to Atlassian Cloud via SCIM..."
        # Build user objects from tracked members (need userPrincipalName for SCIM lookup)
        # Fetch UPNs for all desired user IDs from the tracked member snapshots
        $desiredUserObjects = @()
        foreach ($groupId in $groupIds) {
            $members = Get-TrackedMembers -State $state -GroupId "$stateKey|$groupId"
            foreach ($m in $members) {
                if ($m -isnot [string] -and $m.id -and $m.userPrincipalName) {
                    $desiredUserObjects += $m
                }
            }
        }
        # De-duplicate by id
        $desiredUserObjects = @($desiredUserObjects | Sort-Object { $_.id } -Unique)

        Sync-ScimGroupMembership `
            -GroupDisplayName  $tgtName `
            -DesiredEntraUsers $desiredUserObjects `
            -WhatIf            $false
    }
}

# ── Rebuild registry to reflect current group hierarchy ──────────────────────
# After each sync, re-discover which groups are in each source hierarchy.
# This ensures that if a child group was removed from the hierarchy, its ID
# (and its descendants' IDs) are no longer mapped to this source group,
# preventing unnecessary syncs when those orphaned groups change.

$registryPath = Join-Path $PSScriptRoot "state\group-registry.json"
if (Test-Path $registryPath) {
    Write-Log "Rebuilding group registry to reflect current hierarchy..."

    # Load existing registry
    $regRaw = Get-Content $registryPath -Raw | ConvertFrom-Json
    $reg = @{}
    $regExcludedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $regRaw.PSObject.Properties | ForEach-Object {
        if ($_.Name -eq '__excludedTargetGroupIds__') {
            foreach ($id in @($_.Value)) { $null = $regExcludedIds.Add($id) }
        } else {
            $gId = $_.Name
            $entries = @()
            foreach ($item in @($_.Value)) {
                $e = @{}
                $item.PSObject.Properties | ForEach-Object { $e[$_.Name] = $_.Value }
                $entries += $e
            }
            $reg[$gId] = $entries
        }
    }

    # Process each source group that was synced in this run
    for ($ri = 0; $ri -lt $SourceGroup.Count; $ri++) {
        $rSrcName = $SourceGroup[$ri]
        $rTgtName = Resolve-TargetGroupName -Name $TargetGroup[$ri] -Prefix $prefix

        # Resolve source group ID
        $rSrcObj = Get-GroupByName -Headers $headers -DisplayName $rSrcName
        if (-not $rSrcObj) {
            Write-Log "Registry rebuild: source group '$rSrcName' not found -- skipping." "WARN"
            continue
        }

        # Get current hierarchy
        $currentIds = Get-NestedGroupIds -Headers $headers -GroupId $rSrcObj.id
        $currentIdSet = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]$currentIds, [System.StringComparer]::OrdinalIgnoreCase)

        # Remove this source group's mapping from any group IDs no longer in the hierarchy
        $toClean = @($reg.Keys | Where-Object { -not $currentIdSet.Contains($_) })
        foreach ($gId in $toClean) {
            $reg[$gId] = @($reg[$gId] | Where-Object { $_["sourceGroup"] -ne $rSrcName })
            if ($reg[$gId].Count -eq 0) { $reg.Remove($gId) }
        }

        # Add/update mappings for all current hierarchy IDs
        $newEntry = @{ sourceGroup = $rSrcName; targetGroup = $rTgtName }
        foreach ($gId in $currentIds) {
            if ($reg.ContainsKey($gId)) {
                $existing = @($reg[$gId])
                $alreadyPresent = $existing | Where-Object { $_["sourceGroup"] -eq $rSrcName }
                if (-not $alreadyPresent) {
                    $reg[$gId] = $existing + $newEntry
                }
            } else {
                $reg[$gId] = @($newEntry)
            }
        }

        # Resolve the target group and ensure its ID is in the exclusion list
        $rTgtObj = Get-GroupByName -Headers $headers -DisplayName $rTgtName
        if ($rTgtObj) {
            if ($regExcludedIds.Add($rTgtObj.id)) {
                Write-Log "Registry rebuild: target group '$rTgtName' (ID: $($rTgtObj.id)) added to exclusion list."
            }
        }
    }

    # Save updated registry (preserving excluded target group IDs)
    $regWithExclusions = @{}
    $reg.Keys | ForEach-Object { $regWithExclusions[$_] = $reg[$_] }
    $regWithExclusions['__excludedTargetGroupIds__'] = @($regExcludedIds)
    $regWithExclusions | ConvertTo-Json -Depth 5 | Set-Content -Path $registryPath -Encoding UTF8
    Write-Log "Group registry rebuilt ($($reg.Count) entries, $($regExcludedIds.Count) excluded target IDs)." "SUCCESS"

    # Signal the listener to reload the registry on next loop iteration
    $reloadFlagPath = Join-Path $PSScriptRoot "state\.registry-reload"
    [System.IO.File]::WriteAllText($reloadFlagPath, (Get-Date).ToString("o"))
}

# ── Persist updated state (once, after all pairs processed) ──────────────────

Write-State -Path $StateFilePath -State $state
Write-Log "State persisted to '$StateFilePath'." "SUCCESS"
Write-Log "All $($SourceGroup.Count) group pair(s) processed." "SUCCESS"
