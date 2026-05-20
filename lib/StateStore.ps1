<#
.SYNOPSIS
    Shared state persistence helper for delta query tracking (Version 3).

.DESCRIPTION
    Persists and retrieves a JSON state file that stores:
      - Delta links for each tracked group
      - Snapshot of current members for each tracked group
      - Snapshot of current members of the target group

    Exports:
        Read-State    - Loads state from disk (returns empty structure if missing).
        Write-State   - Saves state to disk.
        Get-DeltaLink - Retrieves the stored delta link for a given group ID.
        Set-DeltaLink - Updates the stored delta link for a given group ID.
        Get-TrackedMembers  - Returns the tracked member set for a given group ID.
        Set-TrackedMembers  - Replaces the tracked member set for a given group ID.
#>

function Read-State {
    <#
    .SYNOPSIS
        Reads the state file from disk. Returns an empty state object if the file does not exist.

    .PARAMETER Path
        Full path to the JSON state file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        Write-Verbose "[StateStore] Reading state from '$Path'"
        try {
            $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
            $parsed = $raw | ConvertFrom-Json
            $h = @{}
            $parsed.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
            return $h
        }
        catch {
            Write-Warning "[StateStore] Failed to parse state file '$Path': $($_.Exception.Message). Starting with empty state."
        }
    }
    else {
        Write-Verbose "[StateStore] State file '$Path' not found -- initialising empty state."
    }

    return @{
        deltaLinks     = @{}
        trackedMembers = @{}
        targetMembers  = @()
    }
}

function Write-State {
    <#
    .SYNOPSIS
        Writes the state object to disk as JSON.

    .PARAMETER Path
        Full path to the JSON state file.

    .PARAMETER State
        The state hashtable to persist.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [hashtable]$State
    )

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $json = $State | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
    Write-Verbose "[StateStore] State written to '$Path'"
}

function Get-DeltaLink {
    <#
    .SYNOPSIS
        Returns the stored delta link for the given group ID, or $null if not present.

    .PARAMETER State
        The state hashtable (as returned by Read-State).

    .PARAMETER GroupId
        The Entra ID object ID of the group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if ($State.deltaLinks -and $State.deltaLinks.ContainsKey($GroupId)) {
        return $State.deltaLinks[$GroupId]
    }
    return $null
}

function Set-DeltaLink {
    <#
    .SYNOPSIS
        Stores or updates the delta link for a given group ID in the state object.

    .PARAMETER State
        The state hashtable to update (mutated in place).

    .PARAMETER GroupId
        The Entra ID object ID of the group.

    .PARAMETER DeltaLink
        The delta link URL returned by the Graph delta query.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [string]$DeltaLink = $null
    )

    if (-not $State.deltaLinks) { $State.deltaLinks = @{} }
    if ($DeltaLink) {
        $State.deltaLinks[$GroupId] = $DeltaLink
    } else {
        # No delta link (e.g. group doesn't support delta queries) -- remove any stored link
        if ($State.deltaLinks.ContainsKey($GroupId)) {
            $State.deltaLinks.Remove($GroupId)
        }
    }
}

function Get-TrackedMembers {
    <#
    .SYNOPSIS
        Returns the tracked member array for the given group ID. Returns empty array if not tracked.

    .PARAMETER State
        The state hashtable.

    .PARAMETER GroupId
        The Entra ID object ID of the group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    if ($State.trackedMembers -and $State.trackedMembers.ContainsKey($GroupId)) {
        return @($State.trackedMembers[$GroupId])
    }
    return @()
}

function Set-TrackedMembers {
    <#
    .SYNOPSIS
        Replaces the tracked member list for a given group ID in the state object.

    .PARAMETER State
        The state hashtable to update (mutated in place).

    .PARAMETER GroupId
        The Entra ID object ID of the group.

    .PARAMETER Members
        Array of member objects (each with at least 'id' and 'userPrincipalName').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$State,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [AllowEmptyCollection()]
        [array]$Members = @()
    )

    if (-not $State.trackedMembers) { $State.trackedMembers = @{} }
    $State.trackedMembers[$GroupId] = $Members
}
