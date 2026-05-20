<#
.SYNOPSIS
    Shared Microsoft Graph API helper functions for Entra ID group operations.

.DESCRIPTION
    All functions accept a pre-built Headers hashtable (from Get-GraphHeaders in GraphAuth.ps1).
    Rate limiting (HTTP 429) is handled transparently with exponential backoff.

    Exports:
        Invoke-GraphRequest          - Core wrapper with retry/backoff logic.
        Resolve-TargetGroupName      - Applies configured prefix to target group name if needed.
        Get-GroupByName              - Looks up a group by display name; returns $null if not found.
        New-EntraGroup               - Creates a new security group.
        Get-GroupTransitiveMembers   - Returns all transitive (flattened) user members of a group.
        Get-GroupDirectMembers       - Returns direct members of a group (used for nested group discovery).
        Get-GroupMembers             - Returns current direct members of a group.
        Add-GroupMembers             - Adds members to a group (batched, up to 20 per request).
        Remove-GroupMembers          - Removes members from a group (one at a time, per Graph API).
        Sync-GroupMembership         - Diffs desired vs actual membership and adds/removes as needed.
        Get-NestedGroupIds           - Returns all nested group IDs within a source group (for webhook subscriptions).
        Get-GroupMembersDelta        - Executes a delta query for a group's members.
#>

$script:GraphBaseUrl = "https://graph.microsoft.com/v1.0"

#region Core HTTP helper

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
        Calls a Microsoft Graph API endpoint with automatic retry on 429 (rate limit) and 503.

    .PARAMETER Headers
        Authorization headers hashtable (from Get-GraphHeaders).

    .PARAMETER Method
        HTTP method: GET, POST, PATCH, DELETE.

    .PARAMETER Uri
        Full Graph API URI.

    .PARAMETER Body
        Optional request body (will be serialised to JSON if a hashtable/object is passed).

    .PARAMETER MaxRetries
        Maximum number of retry attempts on transient errors. Default: 5.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [ValidateSet("GET","POST","PATCH","DELETE")]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [object]$Body = $null,

        [int]$MaxRetries = 5
    )

    $attempt = 0
    $bodyJson = $null
    if ($null -ne $Body) {
        $bodyJson = ($Body | ConvertTo-Json -Depth 20 -Compress)
    }

    while ($true) {
        $attempt++
        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                Headers     = $Headers
                ErrorAction = "Stop"
            }
            if ($bodyJson) {
                $params["Body"] = $bodyJson
            }

            $response = Invoke-RestMethod @params
            return $response
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $retryAfter = 10
            if ($statusCode -in @(429, 503)) {
                # Try to read Retry-After header
                try {
                    $ra = $_.Exception.Response.Headers["Retry-After"]
                    if ($ra) { $retryAfter = [int]$ra }
                }
                catch {}

                if ($attempt -le $MaxRetries) {
                    $wait = [Math]::Max($retryAfter, [Math]::Pow(2, $attempt))
                    Write-Warning "[GraphGroups] HTTP $statusCode on $Method $Uri — retrying in ${wait}s (attempt $attempt/$MaxRetries)"
                    Start-Sleep -Seconds $wait
                    continue
                }
            }

            # Non-retryable or exhausted retries
            $msg = $_.Exception.Message
            throw "[GraphGroups] Graph API error (HTTP $statusCode) on $Method $Uri : $msg"
        }
    }
}

#endregion

#region Pagination helper

function Get-GraphPagedResults {
    <#
    .SYNOPSIS
        Follows @odata.nextLink pagination and returns all results as a flat array.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$Uri
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-GraphRequest -Headers $Headers -Method GET -Uri $nextUri
        if ($response.value) {
            foreach ($item in $response.value) { $results.Add($item) }
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return $results.ToArray()
}

#endregion

#region Prefix helper

function Resolve-TargetGroupName {
    <#
    .SYNOPSIS
        Applies the configured target group prefix to the supplied name if not already present.

    .PARAMETER Name
        The target group name as supplied by the user.

    .PARAMETER Prefix
        The prefix string from config (may be null/empty — in which case Name is returned unchanged).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string]$Prefix = ""
    )

    if ([string]::IsNullOrWhiteSpace($Prefix)) {
        return $Name
    }

    if ($Name.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Verbose "[GraphGroups] Target group name '$Name' already has prefix '$Prefix' — no change."
        return $Name
    }

    $resolved = "$Prefix$Name"
    Write-Host "[INFO] Target group name normalised to: '$resolved' (prefix '$Prefix' applied)" -ForegroundColor Cyan
    return $resolved
}

#endregion

#region Group lookup / creation

function Get-GroupByName {
    <#
    .SYNOPSIS
        Looks up an Entra ID group by its exact display name.
        Returns the group object, or $null if not found.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER DisplayName
        Exact display name to search for.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    $encodedName = [Uri]::EscapeDataString($DisplayName)
    $uri = "$script:GraphBaseUrl/groups?`$filter=displayName eq '$encodedName'&`$select=id,displayName,mailNickname,groupTypes,securityEnabled"

    Write-Verbose "[GraphGroups] Looking up group: '$DisplayName'"
    $response = Invoke-GraphRequest -Headers $Headers -Method GET -Uri $uri

    if ($response.value -and $response.value.Count -gt 0) {
        return $response.value[0]
    }
    return $null
}

function New-EntraGroup {
    <#
    .SYNOPSIS
        Creates a new mail-disabled security group in Entra ID.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER DisplayName
        Display name for the new group.

    .PARAMETER Description
        Optional description for the group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$DisplayName,

        [string]$Description = "Flattened group managed by Sync-FlattenedGroup"
    )

    # Mail nickname: strip spaces and special chars, max 64 chars
    $mailNickname = ($DisplayName -replace '[^A-Za-z0-9_-]', '') 
    if ($mailNickname.Length -gt 64) { $mailNickname = $mailNickname.Substring(0, 64) }
    if ([string]::IsNullOrWhiteSpace($mailNickname)) { $mailNickname = "flattenedgroup" }

    $body = @{
        displayName     = $DisplayName
        description     = $Description
        mailNickname    = $mailNickname
        mailEnabled     = $false
        securityEnabled = $true
        groupTypes      = @()
    }

    Write-Host "[INFO] Creating new security group: '$DisplayName'" -ForegroundColor Cyan
    $newGroup = Invoke-GraphRequest -Headers $Headers -Method POST -Uri "$script:GraphBaseUrl/groups" -Body $body
    Write-Host "[INFO] Group created with ID: $($newGroup.id)" -ForegroundColor Green
    return $newGroup
}

#endregion

#region Member retrieval

function Get-GroupTransitiveMembers {
    <#
    .SYNOPSIS
        Returns all transitive (recursively flattened) user members of a group.
        Only returns objects of type #microsoft.graph.user (groups are excluded).

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the source group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    Write-Verbose "[GraphGroups] Getting transitive members of group '$GroupId'"
    $uri = "$script:GraphBaseUrl/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=id,displayName,userPrincipalName"
    $members = Get-GraphPagedResults -Headers $Headers -Uri $uri
    Write-Verbose "[GraphGroups] Found $($members.Count) transitive user members"
    return $members
}

function Get-GroupMembers {
    <#
    .SYNOPSIS
        Returns direct members of a group (all types).

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    Write-Verbose "[GraphGroups] Getting direct members of group '$GroupId'"
    $uri = "$script:GraphBaseUrl/groups/$GroupId/members?`$select=id,displayName,userPrincipalName,`@odata.type"
    return Get-GraphPagedResults -Headers $Headers -Uri $uri
}

function Get-NestedGroupIds {
    <#
    .SYNOPSIS
        Returns the IDs of all groups nested within the given group (transitively).
        Used by Version 2 to identify which groups to subscribe to for change notifications.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the source group.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId
    )

    Write-Verbose "[GraphGroups] Getting nested group IDs within group '$GroupId'"
    $uri = "$script:GraphBaseUrl/groups/$GroupId/transitiveMembers/microsoft.graph.group?`$select=id,displayName"
    $nestedGroups = Get-GraphPagedResults -Headers $Headers -Uri $uri

    # Always include the root group itself
    $ids = @($GroupId) + ($nestedGroups | ForEach-Object { $_.id })
    Write-Verbose "[GraphGroups] Found $($ids.Count) group(s) of interest (including root)"
    return $ids
}

#endregion

#region Member add / remove

function Add-GroupMembers {
    <#
    .SYNOPSIS
        Adds an array of users to a group. Uses batch requests (up to 20 members per call).

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the target group.

    .PARAMETER MemberIds
        Array of user object IDs to add.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string[]]$MemberIds
    )

    if ($MemberIds.Count -eq 0) { return }

    # Graph supports up to 20 members per PATCH members/$ref add
    $batchSize = 20
    $total = $MemberIds.Count
    $added = 0

    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $batch = $MemberIds[$i..([Math]::Min($i + $batchSize - 1, $total - 1))]
        $body = @{
            "members@odata.bind" = @($batch | ForEach-Object { "$script:GraphBaseUrl/directoryObjects/$_" })
        }
        Invoke-GraphRequest -Headers $Headers -Method PATCH -Uri "$script:GraphBaseUrl/groups/$GroupId" -Body $body | Out-Null
        $added += $batch.Count
        Write-Verbose "[GraphGroups] Added $added/$total members to group '$GroupId'"
    }

    Write-Host "[INFO] Added $total member(s) to group '$GroupId'" -ForegroundColor Green
}

function Remove-GroupMembers {
    <#
    .SYNOPSIS
        Removes an array of users from a group. Graph requires one DELETE per member reference.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the target group.

    .PARAMETER MemberIds
        Array of user object IDs to remove.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string[]]$MemberIds
    )

    if ($MemberIds.Count -eq 0) { return }

    $total = $MemberIds.Count
    $removed = 0

    foreach ($memberId in $MemberIds) {
        $uri = "$script:GraphBaseUrl/groups/$GroupId/members/$memberId/`$ref"
        Invoke-GraphRequest -Headers $Headers -Method DELETE -Uri $uri | Out-Null
        $removed++
        Write-Verbose "[GraphGroups] Removed member $removed/$total from group '$GroupId'"
    }

    Write-Host "[INFO] Removed $total member(s) from group '$GroupId'" -ForegroundColor Yellow
}

#endregion

#region Membership sync

function Sync-GroupMembership {
    <#
    .SYNOPSIS
        Computes the diff between desired and actual group membership and applies adds/removes.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the target group.

    .PARAMETER DesiredMemberIds
        Array of user object IDs that SHOULD be in the group.

    .PARAMETER ActualMemberIds
        Array of user object IDs currently IN the group (pass $null to fetch live).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string[]]$DesiredMemberIds,

        [string[]]$ActualMemberIds = $null
    )

    if ($null -eq $ActualMemberIds) {
        $currentMembers = Get-GroupMembers -Headers $Headers -GroupId $GroupId
        $ActualMemberIds = @($currentMembers | ForEach-Object { $_.id })
    }

    $desiredSet = [System.Collections.Generic.HashSet[string]]::new($DesiredMemberIds, [System.StringComparer]::OrdinalIgnoreCase)
    $actualSet  = [System.Collections.Generic.HashSet[string]]::new($ActualMemberIds,  [System.StringComparer]::OrdinalIgnoreCase)

    $toAdd    = @($desiredSet | Where-Object { -not $actualSet.Contains($_) })
    $toRemove = @($actualSet  | Where-Object { -not $desiredSet.Contains($_) })

    Write-Host "[INFO] Sync summary — Add: $($toAdd.Count)  Remove: $($toRemove.Count)  Unchanged: $($actualSet.Count - $toRemove.Count)" -ForegroundColor Cyan

    if ($toAdd.Count -gt 0)    { Add-GroupMembers    -Headers $Headers -GroupId $GroupId -MemberIds $toAdd }
    if ($toRemove.Count -gt 0) { Remove-GroupMembers -Headers $Headers -GroupId $GroupId -MemberIds $toRemove }

    if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
        Write-Host "[INFO] Target group membership is already up to date — no changes needed." -ForegroundColor Green
    }
}

#endregion

#region Delta query

function Get-GroupMembersDelta {
    <#
    .SYNOPSIS
        Executes a delta query for a group's members.
        On first run (no stored delta link), performs a full sync and returns the delta link.
        On subsequent runs, uses the stored delta link to get only changes since last run.

    .PARAMETER Headers
        Authorization headers.

    .PARAMETER GroupId
        The Entra ID object ID of the group.

    .PARAMETER DeltaLink
        The delta link from the previous run. Pass $null for initial full sync.

    .OUTPUTS
        A hashtable with keys:
            Changes   - Array of member change objects (each has 'id' and '@removed' if deleted)
            DeltaLink - The new delta link to store for next run
            FullSet   - On first run only: full member array (null on incremental runs)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [string]$DeltaLink = $null
    )

    $startUri = $DeltaLink ?? "$script:GraphBaseUrl/groups/$GroupId/members/delta?`$select=id,displayName,userPrincipalName"

    $allChanges = [System.Collections.Generic.List[object]]::new()
    $nextUri    = $startUri
    $deltaLink  = $null

    while ($nextUri) {
        $response = Invoke-GraphRequest -Headers $Headers -Method GET -Uri $nextUri

        if ($response.value) {
            foreach ($item in $response.value) { $allChanges.Add($item) }
        }

        if ($response.'@odata.deltaLink') {
            $deltaLink = $response.'@odata.deltaLink'
            $nextUri   = $null
        }
        elseif ($response.'@odata.nextLink') {
            $nextUri = $response.'@odata.nextLink'
        }
        else {
            $nextUri = $null
        }
    }

    return @{
        Changes   = $allChanges.ToArray()
        DeltaLink = $deltaLink
    }
}

#endregion
