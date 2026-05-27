<#
.SYNOPSIS
    Atlassian User Provisioning (SCIM 2.0) helper functions.

.DESCRIPTION
    Provides PowerShell functions for interacting with the Atlassian Cloud User Provisioning
    REST API (SCIM 2.0). Used by Sync-FlattenedGroup.ps1 and Invoke-DeltaSync.ps1 when the
    -SyncToAtlassian flag is supplied.

    Base URL: https://api.atlassian.com/scim/directory/{directoryId}
    Auth:     Bearer token (SCIM API key from Atlassian admin)

    Key identifier notes:
      - Entra ID members have a userPrincipalName (UPN) which maps to the SCIM 'userName'
      - Atlassian SCIM group operations require the Atlassian 'id' (UUID), NOT the UPN
      - This module bridges the two by looking up each UPN via GET /Users?filter=userName eq "upn"
#>

$script:ScimBaseUrl = $null
$script:ScimHeaders = $null

# ── Initialisation ─────────────────────────────────────────────────────────────

Add-Type -AssemblyName System.Web

function Initialize-ScimClient {
    <#
    .SYNOPSIS
        Initialises the SCIM client with the directory ID and API key.
    .PARAMETER DirectoryId
        The Atlassian SCIM directory ID (found in the Directory base URL).
    .PARAMETER ApiKey
        The SCIM API key generated in Atlassian admin.
    #>
    param(
        [Parameter(Mandatory)][string]$DirectoryId,
        [Parameter(Mandatory)][string]$ApiKey
    )
    $script:ScimBaseUrl = "https://api.atlassian.com/scim/directory/$DirectoryId"
    $script:ScimHeaders = @{
        'Authorization' = "Bearer $ApiKey"
        'Content-Type'  = 'application/scim+json'
        'Accept'        = 'application/scim+json'
    }
    Write-Verbose "[AtlassianScim] Initialised SCIM client for directory '$DirectoryId'"
}

# ── Internal request helper ───────────────────────────────────────────────────

function Invoke-ScimRequest {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body = $null,
        [int]$MaxRetries = 5
    )

    if (-not $script:ScimBaseUrl) {
        throw "[AtlassianScim] SCIM client not initialised. Call Initialize-ScimClient first."
    }

    $attempt = 0
    while ($true) {
        try {
            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $script:ScimHeaders
                UseBasicParsing = $true
            }
            if ($Body) {
                $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
            }
            $response = Invoke-WebRequest @params
            if ($response.Content) {
                # Content may be a byte array when Content-Type is application/scim+json
                # Decode explicitly as UTF-8 string before parsing JSON
                if ($response.Content -is [byte[]]) {
                    $contentStr = [System.Text.Encoding]::UTF8.GetString($response.Content)
                } else {
                    $contentStr = $response.Content
                }
                return $contentStr | ConvertFrom-Json
            }
            return $null
        }
        catch {
            $statusCode = 0
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            # Retry on 429 (rate limit) or 503 (service unavailable)
            if ($statusCode -in @(429, 503) -and $attempt -lt $MaxRetries) {
                $attempt++
                $wait = [Math]::Pow(2, $attempt)
                Write-Host "[INFO] SCIM rate limited (HTTP $statusCode) -- retrying in ${wait}s (attempt $attempt/$MaxRetries)..." -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
            } else {
                $msg = $_.Exception.Message
                throw "[AtlassianScim] SCIM API error (HTTP $statusCode) on $Method $Uri : $msg"
            }
        }
    }
}

# ── Group operations ──────────────────────────────────────────────────────────

function Get-ScimGroupByName {
    <#
    .SYNOPSIS
        Finds a SCIM group by exact displayName. Returns $null if not found.
    #>
    param([Parameter(Mandatory)][string]$DisplayName)

    # Use simple query string encoding -- some SCIM implementations are sensitive to encoding
    $filter  = "displayName eq `"$DisplayName`""
    $encoded = [System.Web.HttpUtility]::UrlEncode($filter)
    $uri     = "$script:ScimBaseUrl/Groups?filter=$encoded&count=1"

    try {
        $result  = Invoke-ScimRequest -Method GET -Uri $uri
        $resProp = $result.PSObject.Properties['Resources']
        if ($resProp -and $resProp.Value -and $resProp.Value.Count -gt 0) {
            return $resProp.Value[0]
        }
    } catch {
        Write-Verbose "[AtlassianScim] Get-ScimGroupByName filter query failed: $_ -- falling back to list search"
    }

    # Fallback: list all groups and match by displayName (handles encoding/filter issues)
    $uri2   = "$script:ScimBaseUrl/Groups?count=200"
    $result2 = Invoke-ScimRequest -Method GET -Uri $uri2
    $resProp2 = $result2.PSObject.Properties['Resources']
    if ($resProp2 -and $resProp2.Value) {
        $match = $resProp2.Value | Where-Object {
            $dnProp = $_.PSObject.Properties['displayName']
            $dnProp -and $dnProp.Value -eq $DisplayName
        }
        if ($match) { return @($match)[0] }
    }
    return $null
}

function New-ScimGroup {
    <#
    .SYNOPSIS
        Creates a new SCIM group. Returns the created group object.
    #>
    param([Parameter(Mandatory)][string]$DisplayName)

    Write-Host "[INFO] Creating new Atlassian SCIM group: '$DisplayName'" -ForegroundColor Cyan
    $body = @{
        schemas     = @("urn:ietf:params:scim:schemas:core:2.0:Group")
        displayName = $DisplayName
    }
    try {
        $group  = Invoke-ScimRequest -Method POST -Uri "$script:ScimBaseUrl/Groups" -Body $body
    } catch {
        # 409 Conflict means the group already exists -- fetch and return it instead
        if ($_ -match 'HTTP 409') {
            Write-Host "[INFO] SCIM group '$DisplayName' already exists (409 Conflict) -- fetching existing group." -ForegroundColor Yellow
            $existing = Get-ScimGroupByName -DisplayName $DisplayName
            if ($existing) { return $existing }
            throw "[AtlassianScim] Group '$DisplayName' returned 409 but could not be found by name."
        }
        throw
    }
    $idProp  = $group.PSObject.Properties['id']
    $groupId = if ($idProp) { $idProp.Value } else { '(unknown)' }
    Write-Host "[INFO] SCIM group created with ID: $groupId" -ForegroundColor Green
    return $group
}

function Get-ScimGroupMembers {
    <#
    .SYNOPSIS
        Returns the current member Atlassian IDs for a SCIM group.
    #>
    param([Parameter(Mandatory)][string]$GroupId)

    $group   = Invoke-ScimRequest -Method GET -Uri "$script:ScimBaseUrl/Groups/$GroupId"
    $members = $group.PSObject.Properties['members']
    if ($members -and $members.Value) {
        return @($members.Value | ForEach-Object {
            $valProp = $_.PSObject.Properties['value']
            if ($valProp) { $valProp.Value }
        } | Where-Object { $_ })
    }
    return @()
}

function Add-ScimGroupMembers {
    <#
    .SYNOPSIS
        Adds Atlassian user IDs to a SCIM group. Batched in groups of 50.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$AtlassianUserIds
    )

    if ($AtlassianUserIds.Count -eq 0) { return }

    $batchSize = 50
    $total     = $AtlassianUserIds.Count
    $added     = 0

    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $batch = $AtlassianUserIds[$i..([Math]::Min($i + $batchSize - 1, $total - 1))]
        $body  = @{
            schemas    = @("urn:ietf:params:scim:api:messages:2.0:PatchOp")
            Operations = @(
                @{
                    op    = "add"
                    path  = "members"
                    value = @($batch | ForEach-Object { @{ value = $_ } })
                }
            )
        }
        Invoke-ScimRequest -Method PATCH -Uri "$script:ScimBaseUrl/Groups/$GroupId" -Body $body | Out-Null
        $added += $batch.Count
        Write-Verbose "[AtlassianScim] Added $added/$total members to SCIM group '$GroupId'"
    }
    Write-Host "[INFO] Added $total member(s) to Atlassian SCIM group '$GroupId'" -ForegroundColor Green
}

function Remove-ScimGroupMembers {
    <#
    .SYNOPSIS
        Removes Atlassian user IDs from a SCIM group. Batched in groups of 50.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$AtlassianUserIds
    )

    if ($AtlassianUserIds.Count -eq 0) { return }

    $batchSize = 50
    $total     = $AtlassianUserIds.Count
    $removed   = 0

    for ($i = 0; $i -lt $total; $i += $batchSize) {
        $batch = $AtlassianUserIds[$i..([Math]::Min($i + $batchSize - 1, $total - 1))]
        $body  = @{
            schemas    = @("urn:ietf:params:scim:api:messages:2.0:PatchOp")
            Operations = @(
                @{
                    op    = "remove"
                    path  = "members"
                    value = @($batch | ForEach-Object { @{ value = $_ } })
                }
            )
        }
        Invoke-ScimRequest -Method PATCH -Uri "$script:ScimBaseUrl/Groups/$GroupId" -Body $body | Out-Null
        $removed += $batch.Count
        Write-Verbose "[AtlassianScim] Removed $removed/$total members from SCIM group '$GroupId'"
    }
    Write-Host "[INFO] Removed $total member(s) from Atlassian SCIM group '$GroupId'" -ForegroundColor Green
}

# ── User lookup and ID resolution ─────────────────────────────────────────────

function Get-ScimUserByUpn {
    <#
    .SYNOPSIS
        Looks up an Atlassian SCIM user by UPN (userPrincipalName / email).
        Returns the Atlassian user object, or $null if not found.
    #>
    param([Parameter(Mandatory)][string]$Upn)

    $encoded = [System.Uri]::EscapeDataString("userName eq `"$Upn`"")
    $uri     = "$script:ScimBaseUrl/Users?filter=$encoded&count=1"
    $result  = Invoke-ScimRequest -Method GET -Uri $uri

    $resProp = $result.PSObject.Properties['Resources']
    if ($resProp -and $resProp.Value -and $resProp.Value.Count -gt 0) {
        return $resProp.Value[0]
    }
    return $null
}

function Get-ScimObjectId {
    <#
    .SYNOPSIS
        Safely retrieves the 'id' property from a SCIM object under StrictMode.
    #>
    param([Parameter(Mandatory)][object]$ScimObject)
    $idProp = $ScimObject.PSObject.Properties['id']
    if ($idProp) { return $idProp.Value }
    return $null
}

function Resolve-ScimUserIds {
    <#
    .SYNOPSIS
        Resolves a list of Entra user objects (with userPrincipalName) to Atlassian SCIM user IDs.
        Returns a hashtable: UPN -> Atlassian ID. Users not found in Atlassian are logged and skipped.
    .PARAMETER EntraUsers
        Array of Entra user objects with at least a 'userPrincipalName' property.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EntraUsers
    )

    $idMap   = @{}   # UPN -> Atlassian SCIM ID
    $notFound = @()

    foreach ($user in $EntraUsers) {
        $upn = $null
        $upnProp = $user.PSObject.Properties['userPrincipalName']
        if ($upnProp) { $upn = $upnProp.Value }

        if (-not $upn) {
            Write-Host "[WARN] Entra user '$($user.id)' has no userPrincipalName -- skipping." -ForegroundColor Yellow
            continue
        }

        if ($idMap.ContainsKey($upn)) { continue }   # already resolved

        $scimUser = Get-ScimUserByUpn -Upn $upn
        if ($scimUser) {
            $atlasId = Get-ScimObjectId -ScimObject $scimUser
            if ($atlasId) {
                $idMap[$upn] = $atlasId
                Write-Verbose "[AtlassianScim] Resolved UPN '$upn' -> Atlassian ID '$atlasId'"
            } else {
                Write-Host "[WARN] Atlassian user found for '$upn' but has no 'id' -- skipping." -ForegroundColor Yellow
            }
        } else {
            $notFound += $upn
            Write-Host "[WARN] User '$upn' not found in Atlassian SCIM directory -- skipping." -ForegroundColor Yellow
        }
    }  # end foreach user

    if ($notFound.Count -gt 0) {
        Write-Host "[WARN] $($notFound.Count) user(s) not found in Atlassian directory and will not be synced." -ForegroundColor Yellow
    }

    return $idMap
}

# ── Top-level sync function ───────────────────────────────────────────────────

function Sync-ScimGroupMembership {
    <#
    .SYNOPSIS
        Syncs a flat list of Entra users into an Atlassian SCIM group.
        Creates the group if it doesn't exist. Diffs current vs desired membership.

    .PARAMETER GroupDisplayName
        The Atlassian SCIM group displayName to sync to.

    .PARAMETER DesiredEntraUsers
        Array of Entra user objects (must have userPrincipalName) representing desired members.

    .PARAMETER WhatIf
        If true, logs what would change without making any API write calls.

    .OUTPUTS
        None. Writes log messages to host.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupDisplayName,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DesiredEntraUsers,
        [bool]$WhatIf = $false
    )

    # ── Resolve or create the target SCIM group ───────────────────────────────

    $scimGroup = Get-ScimGroupByName -DisplayName $GroupDisplayName
    $groupIsNew = $false

    if (-not $scimGroup) {
        Write-Host "[INFO] Atlassian SCIM group '$GroupDisplayName' not found." -ForegroundColor Yellow
        if ($WhatIf) {
            Write-Host "[WhatIf] Would create Atlassian SCIM group '$GroupDisplayName'" -ForegroundColor Yellow
            return
        }
        $scimGroup  = New-ScimGroup -DisplayName $GroupDisplayName
        $groupIsNew = $true
    } else {
        $scimGroupId = $scimGroup.PSObject.Properties['id']
        Write-Host "[INFO] Atlassian SCIM group '$GroupDisplayName' found (ID: $(if ($scimGroupId) { $scimGroupId.Value } else { 'unknown' }))" -ForegroundColor Cyan
    }

    # ── Resolve Entra UPNs to Atlassian SCIM IDs ─────────────────────────────

    $desiredEntraCount = @($DesiredEntraUsers).Count
    Write-Host "[INFO] Resolving $desiredEntraCount Entra user(s) to Atlassian SCIM IDs..." -ForegroundColor Cyan
    $idMap      = Resolve-ScimUserIds -EntraUsers $DesiredEntraUsers
    $desiredIds = @($idMap.Values | ForEach-Object { $_ })
    Write-Host "[INFO] Resolved $($desiredIds.Count)/$desiredEntraCount user(s) to Atlassian IDs." -ForegroundColor Cyan

    # ── Get current SCIM group membership ─────────────────────────────────────

    $scimGroupIdVal = $scimGroup.PSObject.Properties['id']
    $scimGroupId    = if ($scimGroupIdVal) { $scimGroupIdVal.Value } else { throw "[AtlassianScim] Created/found SCIM group has no 'id' property." }
    if ($groupIsNew) {
        $currentIds = @()
    } else {
        $currentIds = @(Get-ScimGroupMembers -GroupId $scimGroupId)
    }
    $currentIdCount = $currentIds.Count
    Write-Host "[INFO] Current Atlassian group membership: $currentIdCount member(s)." -ForegroundColor Cyan

    # ── Compute diff ──────────────────────────────────────────────────────────

    $desiredIds  = @($desiredIds)
    $currentIds  = @($currentIds)

    $desiredSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in $desiredIds)  { if ($id) { $null = $desiredSet.Add($id) } }
    $currentSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in $currentIds)  { if ($id) { $null = $currentSet.Add($id) } }

    $toAdd    = @($desiredSet | Where-Object { -not $currentSet.Contains($_) })
    $toRemove = @($currentSet | Where-Object { -not $desiredSet.Contains($_) })

    $unchangedCount = $currentSet.Count - $toRemove.Count
    Write-Host "[INFO] Atlassian diff -- To add: $($toAdd.Count)  |  To remove: $($toRemove.Count)  |  Unchanged: $unchangedCount" -ForegroundColor Cyan

    # ── Apply changes ─────────────────────────────────────────────────────────

    if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
        Write-Host "[INFO] Atlassian group '$GroupDisplayName' is already up to date." -ForegroundColor Green
        return
    }

    if ($WhatIf) {
        Write-Host "[WhatIf] Would add $($toAdd.Count) and remove $($toRemove.Count) member(s) in Atlassian group '$GroupDisplayName'" -ForegroundColor Yellow
        return
    }

    if ($toAdd.Count -gt 0) {
        Write-Host "[INFO] Adding $($toAdd.Count) member(s) to Atlassian group '$GroupDisplayName'..." -ForegroundColor Cyan
        Add-ScimGroupMembers -GroupId $scimGroupId -AtlassianUserIds $toAdd
    }
    if ($toRemove.Count -gt 0) {
        Write-Host "[INFO] Removing $($toRemove.Count) member(s) from Atlassian group '$GroupDisplayName'..." -ForegroundColor Cyan
        Remove-ScimGroupMembers -GroupId $scimGroupId -AtlassianUserIds $toRemove
    }
    Write-Host "[INFO] Atlassian group '$GroupDisplayName' sync complete." -ForegroundColor Green
}
