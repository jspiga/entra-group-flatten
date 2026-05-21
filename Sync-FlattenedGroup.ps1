<#
.SYNOPSIS
    Version 1 (MVP): Flattens a source Entra ID group and synchronises the result to a target group.

.DESCRIPTION
    Uses the Microsoft Graph API transitive members endpoint to obtain a fully flattened list of
    users from the source group (regardless of nesting depth), then creates or updates the target
    group so that its direct membership exactly matches that flattened list.

    Intended to be run manually or on a schedule (e.g. Windows Task Scheduler, Azure Automation).

    Authentication uses the OAuth2 client credentials flow (App Registration).

.PARAMETER SourceGroup
    Display name of the source group to flatten.

.PARAMETER TargetGroup
    Display name of the target group to create/update. If a targetGroupPrefix is configured and
    the name does not already start with that prefix, the prefix is applied automatically.

.PARAMETER TenantId
    Azure AD tenant ID. Overrides config.json if provided.

.PARAMETER ClientId
    App Registration client ID. Overrides config.json if provided.

.PARAMETER ClientSecret
    App Registration client secret. Overrides config.json if provided.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER WhatIf
    Dry-run mode: shows what would be added/removed without making any changes.

.EXAMPLE
    # Using config.json for credentials, with a target group name that gets auto-prefixed
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

.EXAMPLE
    # Fully explicit, override all config values
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_EngineeringFlat" `
        -TenantId "..." -ClientId "..." -ClientSecret "..."

.EXAMPLE
    # Dry run to preview changes without applying them
    .\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$SourceGroup,

    [Parameter(Mandatory)]
    [string]$TargetGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,

    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
Write-Log "Source group resolved: '$($sourceGroupObj.displayName)' (ID: $($sourceGroupObj.id))" "SUCCESS"

# ── Resolve or create target group ────────────────────────────────────────────

Write-Log "Resolving target group: '$TargetGroup'"
$targetGroupObj = Get-GroupByName -Headers $headers -DisplayName $TargetGroup

if (-not $targetGroupObj) {
    Write-Log "Target group '$TargetGroup' not found." "WARN"
    if ($PSCmdlet.ShouldProcess($TargetGroup, "Create new Entra ID security group")) {
        $targetGroupObj = New-EntraGroup -Headers $headers -DisplayName $TargetGroup `
            -Description "Flattened membership of '$SourceGroup', managed by Sync-FlattenedGroup.ps1"
    } else {
        Write-Log "[WhatIf] Would create group '$TargetGroup'" "WARN"
        exit 0
    }
} else {
    Write-Log "Target group resolved: '$($targetGroupObj.displayName)' (ID: $($targetGroupObj.id))" "SUCCESS"
}

# ── Get flattened source members ──────────────────────────────────────────────

Write-Log "Fetching transitive (flattened) members of source group '$SourceGroup'..."
$flatMembers = @(Get-GroupTransitiveMembers -Headers $headers -GroupId $sourceGroupObj.id)

# De-duplicate by ID (should already be unique from the API, but be safe)
$uniqueMembers = @($flatMembers | Sort-Object id -Unique)
Write-Log "Found $($uniqueMembers.Count) unique user(s) in flattened source group." "SUCCESS"

if ($VerbosePreference -ne 'SilentlyContinue') {
    $uniqueMembers | ForEach-Object { Write-Verbose "  - $($_.userPrincipalName) ($($_.id))" }
}

# ── Get current target group members ─────────────────────────────────────────

Write-Log "Fetching current members of target group '$TargetGroup'..."
$currentMembers = @(Get-GroupMembers -Headers $headers -GroupId $targetGroupObj.id)
$currentUserIds = @($currentMembers | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' } | ForEach-Object { $_.id })
Write-Log "Target group currently has $($currentUserIds.Count) user member(s)."

# ── Compute diff ──────────────────────────────────────────────────────────────

$desiredIds = @($uniqueMembers | ForEach-Object { $_.id })

$desiredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($id in $desiredIds)     { $null = $desiredSet.Add($id) }
$actualSet  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($id in $currentUserIds) { $null = $actualSet.Add($id) }

$toAdd    = @($desiredSet | Where-Object { -not $actualSet.Contains($_) })
$toRemove = @($actualSet  | Where-Object { -not $desiredSet.Contains($_) })

Write-Log "Diff -- To add: $($toAdd.Count)  |  To remove: $($toRemove.Count)  |  Unchanged: $($actualSet.Count - $toRemove.Count)"

# ── Apply changes ─────────────────────────────────────────────────────────────

if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
    Write-Log "Target group is already up to date -- no changes required." "SUCCESS"
} else {
    if ($PSCmdlet.ShouldProcess($TargetGroup, "Sync group membership (add $($toAdd.Count), remove $($toRemove.Count))")) {
        if ($toAdd.Count -gt 0) {
            Write-Log "Adding $($toAdd.Count) member(s) to '$TargetGroup'..."
            Add-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $toAdd
        }
        if ($toRemove.Count -gt 0) {
            Write-Log "Removing $($toRemove.Count) member(s) from '$TargetGroup'..."
            Remove-GroupMembers -Headers $headers -GroupId $targetGroupObj.id -MemberIds $toRemove
        }
        Write-Log "Membership sync complete." "SUCCESS"
    } else {
        Write-Log "[WhatIf] Would add $($toAdd.Count) member(s) and remove $($toRemove.Count) member(s)." "WARN"
    }
}
