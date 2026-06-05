<#
.SYNOPSIS
    Option 1: Flattens a source Entra ID group and synchronises the result to a target group.

.DESCRIPTION
    Uses the Microsoft Graph API transitive members endpoint to obtain a fully flattened list of
    users from the source group (regardless of nesting depth), then creates or updates the target
    group so that its direct membership exactly matches that flattened list.

    Intended to be run manually or on a schedule (e.g. Windows Task Scheduler, Azure Automation).

    Authentication uses the OAuth2 client credentials flow (App Registration).

.PARAMETER SourceGroup
    One or more source group display names to flatten. Accepts an array.

.PARAMETER TargetGroup
    One or more corresponding target group display names (must match SourceGroup order).

.PARAMETER TenantId
    Azure AD tenant ID. Overrides config.json if provided.

.PARAMETER ClientId
    App Registration client ID. Overrides config.json if provided.

.PARAMETER ClientSecret
    App Registration client secret. Overrides config.json if provided.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER SyncTo
    Where to sync the flattened membership. Accepts one or both of: "Entra", "Atlassian".
    Defaults to "Entra" if not specified.
    Use "Atlassian" alone if Entra ID SCIM provisioning is already handling Entra group sync
    (to avoid race conditions). Use "Entra","Atlassian" to sync to both simultaneously.
    Can also be set persistently via the "syncTo" array in config.json.

.PARAMETER WhatIf
    Dry-run mode: shows what would be added/removed without making any changes.

.EXAMPLE
    # Default -- syncs to Entra ID only
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

.EXAMPLE
    # Sync to Atlassian Cloud only (use when Entra SCIM provisioning is active)
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -SyncTo Atlassian

.EXAMPLE
    # Sync to both Entra ID and Atlassian Cloud
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -SyncTo Entra,Atlassian

.EXAMPLE
    # Multiple groups, Atlassian only
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering","All-Finance" -TargetGroup "EngineeringFlat","FinanceFlat" -SyncTo Atlassian

.EXAMPLE
    # Dry run to preview changes without applying them
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string[]]$SourceGroup,

    [Parameter(Mandatory)]
    [string[]]$TargetGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,

    [string]$ConfigPath = "",

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

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $colour = switch ($Level) {
        "INFO"    { "Cyan"    }
        "SUCCESS" { "Green"   }
        "WARN"    { "Yellow"  }
        "ERROR"   { "Red"     }
        default   { "White"   }
    }
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $colour
}

# ── Load libraries ────────────────────────────────────────────────────────────

. (Join-Path $PSScriptRoot "lib/GraphAuth.ps1")
. (Join-Path $PSScriptRoot "lib/GraphGroups.ps1")
. (Join-Path $PSScriptRoot "lib/AtlassianScim.ps1")

# ── Load config ───────────────────────────────────────────────────────────────

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading config from '$ConfigPath'"
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ForEach-Object { $h = @{}; $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h }
} else {
    Write-Log "No config.json found at '$ConfigPath' -- relying on parameters only." "WARN"
}

# Parameter values take precedence over config file
if (-not $TenantId)      { $TenantId     = $config["tenantId"]     }
if (-not $ClientId)      { $ClientId     = $config["clientId"]     }
if (-not $ClientSecret)  { $ClientSecret = $config["clientSecret"] }
$prefix = $config["targetGroupPrefix"]

foreach ($required in @("TenantId","ClientId","ClientSecret")) {
    if (-not (Get-Variable $required -ValueOnly -ErrorAction SilentlyContinue)) {
        throw "Missing required value: '$required'. Provide it as a parameter or in config.json."
    }
}

# ── Resolve SyncTo destinations ───────────────────────────────────────────────

# -SyncTo parameter overrides config; if neither provided, default to Entra
$effectiveSyncTo = if ($SyncTo.Count -gt 0) {
    $SyncTo
} elseif ($config["syncTo"]) {
    @($config["syncTo"])
} else {
    @("Entra")
}

$syncToEntra      = $effectiveSyncTo -contains "Entra"
$syncToAtlassian  = $effectiveSyncTo -contains "Atlassian"

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

# ── Authenticate (once, shared across all group pairs) ───────────────────────

Write-Log "Authenticating to Microsoft Graph (tenant: $TenantId)"
$headers = Get-GraphHeaders -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
Write-Log "Authentication successful." "SUCCESS"

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
    Write-Log "Source group resolved: '$($sourceGroupObj.displayName)' (ID: $($sourceGroupObj.id))" "SUCCESS"

    # ── Get flattened source members ──────────────────────────────────────────

    Write-Log "Fetching transitive (flattened) members of source group '$srcName'..."
    $flatMembers   = @(Get-GroupTransitiveMembers -Headers $headers -GroupId $sourceGroupObj.id)
    $uniqueMembers = @($flatMembers | Sort-Object id -Unique)
    Write-Log "Found $($uniqueMembers.Count) unique user(s) in flattened source group." "SUCCESS"

    if ($VerbosePreference -ne 'SilentlyContinue') {
        $uniqueMembers | ForEach-Object { Write-Verbose "  - $($_.userPrincipalName) ($($_.id))" }
    }

    # ── Apply changes to Entra ID ─────────────────────────────────────────────

    if ($syncToEntra) {
        # Resolve or create target group in Entra ID
        Write-Log "Resolving target group: '$tgtName'"
        $targetGroupObj  = Get-GroupByName -Headers $headers -DisplayName $tgtName
        $targetGroupIsNew = $false

        if (-not $targetGroupObj) {
            Write-Log "Target group '$tgtName' not found." "WARN"
            if ($PSCmdlet.ShouldProcess($tgtName, "Create new Entra ID security group")) {
                $targetGroupObj   = New-EntraGroup -Headers $headers -DisplayName $tgtName `
                    -Description "Flattened membership of '$srcName', managed by Sync-FlattenedGroup.ps1"
                $targetGroupIsNew = $true
            } else {
                Write-Log "[WhatIf] Would create group '$tgtName'" "WARN"
            }
        } else {
            Write-Log "Target group resolved: '$($targetGroupObj.displayName)' (ID: $($targetGroupObj.id))" "SUCCESS"
        }

        if ($targetGroupObj) {
            if ($targetGroupIsNew) {
                Write-Log "Target group '$tgtName' was just created -- skipping member fetch (empty by definition)."
                $currentUserIds = @()
            } else {
                Write-Log "Fetching current members of target group '$tgtName'..."
                $currentMembers = @(Get-GroupMembers -Headers $headers -GroupId $targetGroupObj.id)
                $currentUserIds = @($currentMembers | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' } | ForEach-Object { $_.id })
                Write-Log "Target group currently has $($currentUserIds.Count) user member(s)."
            }

            $desiredIds = @($uniqueMembers | ForEach-Object { $_.id })
            $desiredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($id in $desiredIds)     { $null = $desiredSet.Add($id) }
            $actualSet  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($id in $currentUserIds) { $null = $actualSet.Add($id) }

            $toAdd    = @($desiredSet | Where-Object { -not $actualSet.Contains($_) })
            $toRemove = @($actualSet  | Where-Object { -not $desiredSet.Contains($_) })

            Write-Log "[Entra] Diff -- To add: $($toAdd.Count)  |  To remove: $($toRemove.Count)  |  Unchanged: $($actualSet.Count - $toRemove.Count)"

            if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
                Write-Log "[Entra] Target group '$tgtName' is already up to date." "SUCCESS"
            } else {
                if ($PSCmdlet.ShouldProcess($tgtName, "Sync Entra group membership (add $($toAdd.Count), remove $($toRemove.Count))")) {
                    if ($toAdd.Count -gt 0) {
                        Write-Log "[Entra] Adding $($toAdd.Count) member(s) to '$tgtName'..."
                        Add-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $toAdd
                    }
                    if ($toRemove.Count -gt 0) {
                        Write-Log "[Entra] Removing $($toRemove.Count) member(s) from '$tgtName'..."
                        Remove-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $toRemove
                    }
                    Write-Log "[Entra] Membership sync complete for '$tgtName'." "SUCCESS"
                } else {
                    Write-Log "[Entra] [WhatIf] Would add $($toAdd.Count) member(s) and remove $($toRemove.Count) member(s)." "WARN"
                }
            }
        }
    }

    # ── Apply changes to Atlassian Cloud (SCIM) ───────────────────────────────

    if ($syncToAtlassian) {
        Write-Log "[Atlassian] Syncing '$tgtName' to Atlassian Cloud via SCIM..."
        $whatIfBool = -not $PSCmdlet.ShouldProcess($tgtName, "Sync Atlassian SCIM group")
        Sync-ScimGroupMembership `
            -GroupDisplayName  $tgtName `
            -DesiredEntraUsers $uniqueMembers `
            -WhatIf            $whatIfBool
    }
}

Write-Log "All $($SourceGroup.Count) group pair(s) processed." "SUCCESS"
