<#
.SYNOPSIS
    Starts an HTTP webhook listener that receives Microsoft Graph change notifications and
    triggers Invoke-DeltaSync.ps1 for the affected source group.

.DESCRIPTION
    Listens on the configured port for Graph change notification POSTs. On receipt:
      1. Validates the clientState against the expected value in subscriptions.json.
      2. Looks up the changed group ID in group-registry.json to find the source/target group.
      3. Triggers Invoke-DeltaSync.ps1 as a background job with per-source-group debouncing.
      4. Responds 202 Accepted to Graph immediately.

    Also handles the Graph subscription validation handshake (GET with validationToken).

    Scaling: A single tenant-wide Graph subscription covers all monitored groups. The listener
    uses group-registry.json (built by Register-ChangeNotification.ps1) to route each
    notification to the correct sync job.

.PARAMETER Port
    TCP port to listen on. Default: 8080.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER SubscriptionStorePath
    Path to subscriptions.json. Defaults to ./state/subscriptions.json.

.PARAMETER RegistryPath
    Path to group-registry.json. Defaults to ./state/group-registry.json.

.PARAMETER DeltaSyncScriptPath
    Path to Invoke-DeltaSync.ps1. Defaults to ./Invoke-DeltaSync.ps1.

.PARAMETER DebounceSecs
    Minimum seconds between sync jobs for the same source group. Default: 30.

.PARAMETER LogPath
    Optional path to a log file. If specified, all log output is also written to this file.

.EXAMPLE
    .\Start-WebhookListener.ps1 -Port 8080 -LogPath ".\logs\webhook.log"
#>

[CmdletBinding()]
param(
    [int]$Port                     = 8080,
    [string]$ConfigPath            = "",
    [string]$SubscriptionStorePath = "",
    [string]$RegistryPath          = "",
    [string]$DeltaSyncScriptPath   = "",
    [int]$DebounceSecs             = 30,
    [string]$LogPath               = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# $PSScriptRoot is empty when invoked via 'powershell -File script.ps1' in some PS 5.1 contexts.
# Fall back to the script file's directory via $MyInvocation.
if (-not $PSScriptRoot) {
    $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Resolve default paths now that $PSScriptRoot is guaranteed to be set
if (-not $ConfigPath)            { $ConfigPath            = Join-Path $PSScriptRoot "config.json"              }
if (-not $SubscriptionStorePath) { $SubscriptionStorePath = Join-Path $PSScriptRoot "state/subscriptions.json" }
if (-not $RegistryPath)          { $RegistryPath          = Join-Path $PSScriptRoot "state/group-registry.json"}
if (-not $DeltaSyncScriptPath)   { $DeltaSyncScriptPath   = Join-Path $PSScriptRoot "Invoke-DeltaSync.ps1"     }

# ── Logging ───────────────────────────────────────────────────────────────────

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
    $line = "[$ts] [$Level] $Message"
    Write-Host $line -ForegroundColor $colour
    if ($LogPath) {
        try {
            Add-Content -Path $LogPath -Value $line -Encoding UTF8
        } catch { <# best-effort #> }
    }
}

# ── Ensure log directory exists ───────────────────────────────────────────────

if ($LogPath) {
    $logDir = Split-Path $LogPath -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
}

# ── Load subscription store ───────────────────────────────────────────────────

function Load-SubscriptionStore {
    $store = @{ clientState = "entra-group-flatten-tenant" }
    if (Test-Path $SubscriptionStorePath) {
        $raw = Get-Content $SubscriptionStorePath -Raw | ConvertFrom-Json
        $raw.PSObject.Properties | ForEach-Object { $store[$_.Name] = $_.Value }
    }
    return $store
}

# ── Load group registry ───────────────────────────────────────────────────────

function Load-Registry {
    # Use a case-insensitive hashtable for entries -- UUIDs from Graph notifications
    # may differ in case from those stored in the registry JSON.
    $result = @{
        Entries     = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
        ExcludedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    if (Test-Path $RegistryPath) {
        $raw = Get-Content $RegistryPath -Raw | ConvertFrom-Json
        $raw.PSObject.Properties | ForEach-Object {
            if ($_.Name -eq '__excludedTargetGroupIds__') {
                foreach ($id in @($_.Value)) { $null = $result.ExcludedIds.Add($id) }
            } else {
                $groupId = $_.Name
                # Wrap in @() to normalise: PS 5.1 ConvertFrom-Json collapses single-element
                # JSON arrays to a bare PSCustomObject rather than an array.
                $entries = @()
                foreach ($item in @($_.Value)) {
                    $e = @{}
                    $item.PSObject.Properties | ForEach-Object { $e[$_.Name] = $_.Value }
                    $entries += $e
                }
                $result.Entries[$groupId] = $entries
            }
        }
        Write-Log "Registry loaded: $($result.Entries.Count) group entries, $($result.ExcludedIds.Count) excluded IDs."
    }
    return $result
}

# ── State: per-source-group debounce tracking ─────────────────────────────────

# $pendingSync  - hashtable: sourceGroupName -> last-notification datetime
# $activeJobs   - hashtable: sourceGroupName -> PS Job
$pendingSync = @{}
$activeJobs  = @{}

# ── Helper: flush completed jobs and log their output ────────────────────────

function Flush-CompletedJobs {
    $toRemove = @()
    foreach ($srcName in @($activeJobs.Keys)) {
        $job = $activeJobs[$srcName]
        if ($job.State -in @('Completed','Failed','Stopped')) {
            # Capture stdout
            $out = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($out) { $out | ForEach-Object { Write-Log "  [JOB:$srcName] $_" } }
            # Capture errors
            if ($job.ChildJobs -and $job.ChildJobs[0].Error.Count -gt 0) {
                $job.ChildJobs[0].Error | ForEach-Object {
                    Write-Log "  [JOB:$srcName] ERROR: $($_.ToString())" "ERROR"
                }
            }
            if ($job.State -eq 'Completed') {
                Write-Log "Sync job for '$srcName' completed successfully." "SUCCESS"
            } else {
                Write-Log "Sync job for '$srcName' ended with state: $($job.State)" "WARN"
            }
            Remove-Job -Job $job -Force
            $toRemove += $srcName
        }
    }
    foreach ($k in $toRemove) { $activeJobs.Remove($k) }
}

# ── Helper: trigger a sync job for a source group ────────────────────────────

function Start-SyncJob {
    param(
        [string]$SourceGroupName,
        [string]$TargetGroupName,
        [string]$AbsConfigPath,
        [string]$AbsDeltaSyncScript,
        [string[]]$SyncToDestinations = @("Entra")
    )

    # If a job is already running for this group, skip
    if ($activeJobs.ContainsKey($SourceGroupName)) {
        $existing = $activeJobs[$SourceGroupName]
        if ($existing.State -eq 'Running') {
            Write-Log "Sync job for '$SourceGroupName' already running -- skipping duplicate trigger." "WARN"
            return
        }
    }

    Write-Log "Starting sync job: '$SourceGroupName' -> '$TargetGroupName' (SyncTo: $($SyncToDestinations -join ','))"
    $job = Start-Job -ScriptBlock {
        param($script, $src, $tgt, $cfg, $syncTo)
        # Redirect all streams (including Write-Host / Information stream 6) to output
        & $script -SourceGroup $src -TargetGroup $tgt -ConfigPath $cfg -SyncTo $syncTo *>&1
    } -ArgumentList $AbsDeltaSyncScript, $SourceGroupName, $TargetGroupName, $AbsConfigPath, $SyncToDestinations

    $activeJobs[$SourceGroupName] = $job
}

# ── Startup ───────────────────────────────────────────────────────────────────

Write-Log "Starting webhook listener on port $Port"
Write-Log "Config path:       $ConfigPath"
Write-Log "Registry path:     $RegistryPath"
Write-Log "Delta sync script: $DeltaSyncScriptPath"
Write-Log "Debounce:          ${DebounceSecs}s per source group"
if ($LogPath) { Write-Log "Log file:          $LogPath" }

# Resolve to absolute paths before passing to background jobs
$absConfigPath      = [System.IO.Path]::GetFullPath($ConfigPath)
$absDeltaSyncScript = [System.IO.Path]::GetFullPath($DeltaSyncScriptPath)

# Initial load
$subStore      = Load-SubscriptionStore
$registryData  = Load-Registry
$registry      = $registryData.Entries
$excludedIds   = $registryData.ExcludedIds
Write-Log "Subscription store loaded (clientState: $($subStore['clientState']))."
Write-Log "Group registry loaded: $($registry.Count) group(s) monitored."

# Read SyncTo destinations from config
$configSyncTo = @("Entra")   # default
if (Test-Path $ConfigPath) {
    $cfgRaw   = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $syncToProp = $cfgRaw.PSObject.Properties['syncTo']
    if ($syncToProp -and $syncToProp.Value) {
        $configSyncTo = @($syncToProp.Value)
    }
}
Write-Log "Sync destinations (from config): $($configSyncTo -join ', ')"

if ($registry.Count -eq 0) {
    Write-Log "Group registry is empty. Run Register-ChangeNotification.ps1 to populate the group registry." "WARN"
}

# ── Start HTTP listener ───────────────────────────────────────────────────────

$listener = [System.Net.HttpListener]::new()

# Determine the best prefix to use and attempt to start the listener.
# Priority: http://+:port/ (public, requires URL ACL or admin) -> http://localhost:port/ (loopback only)
$isAdmin    = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$prefixFull = "http://+:$Port/"
$prefixLocal = "http://localhost:$Port/"

$prefix = $null

# Try public prefix first (works if admin OR if URL ACL registered via Register-UrlAcl.ps1)
$listener.Prefixes.Add($prefixFull)
try {
    $listener.Start()
    $prefix = $prefixFull
    Write-Log "Listener started on $prefix (public binding)" "SUCCESS"
}
catch {
    Write-Log "Could not bind to $prefixFull -- falling back to localhost only." "WARN"
    Write-Log "To enable public binding, run once as Administrator: .\Register-UrlAcl.ps1 -Port $Port" "WARN"
    # Note: do NOT call $listener.Prefixes.Clear() here -- after a failed Start(),
    # the Prefixes collection may be null on PS 5.1 / .NET 4.x, causing a NullReferenceException.
    # Instead, create a fresh HttpListener for the localhost binding.

    # Try localhost prefix
    $listener2 = [System.Net.HttpListener]::new()
    $listener2.Prefixes.Add($prefixLocal)
    try {
        $listener2.Start()
        $listener = $listener2
        $prefix   = $prefixLocal
        Write-Log "Listener started on $prefix (localhost only -- not reachable from external hosts)" "WARN"
    }
    catch {
        Write-Log "Failed to start listener on port ${Port}: $_" "ERROR"
        Write-Log "Run .\Register-UrlAcl.ps1 -Port $Port as Administrator to fix this." "ERROR"
        throw
    }
}

Write-Log "Press Ctrl+C to stop."

# ── Main loop ─────────────────────────────────────────────────────────────────
# Use BeginGetContext / AsyncWaitHandle so that GetContext() never blocks the
# main thread and no background [System.Threading.Thread] is needed.
# A background thread in PS 5.1 can crash the entire powershell.exe process if
# it throws an unhandled exception -- AsyncWaitHandle avoids this entirely.

$asyncResult = $listener.BeginGetContext($null, $null)
Write-Log "Request handler ready (async)." "SUCCESS"

$loopIteration = 0
try {
    while ($listener.IsListening) {
        $loopIteration++
        if ($loopIteration % 300 -eq 0) {
            Write-Log "Listener alive -- iteration $loopIteration, pending syncs: $($pendingSync.Count), active jobs: $($activeJobs.Count)"
        }

        # Housekeeping: flush completed jobs
        Flush-CompletedJobs

        # Check if a sync job has signalled that the registry needs reloading
        $reloadFlag = Join-Path $PSScriptRoot "state\.registry-reload"
        if (Test-Path $reloadFlag) {
            Remove-Item $reloadFlag -Force -ErrorAction SilentlyContinue
            $registryData = Load-Registry
            $registry     = $registryData.Entries
            $excludedIds  = $registryData.ExcludedIds
            Write-Log "Group registry reloaded ($($registry.Count) entries, $($excludedIds.Count) excluded target IDs)." "INFO"
        }

        # Check for pending debounced syncs that are now due
        $now = Get-Date
        foreach ($srcName in @($pendingSync.Keys)) {
            $pending = $pendingSync[$srcName]
            if (($now - $pending.Time).TotalSeconds -ge $DebounceSecs) {
                $pendingSync.Remove($srcName)
                Start-SyncJob -SourceGroupName $srcName `
                              -TargetGroupName $pending.TargetGroup `
                              -AbsConfigPath $absConfigPath `
                              -AbsDeltaSyncScript $absDeltaSyncScript `
                              -SyncToDestinations $configSyncTo
            }
        }

        # Process any completed async context requests
        while ($asyncResult.IsCompleted) {
            $context = $null
            try {
                $context = $listener.EndGetContext($asyncResult)
            } catch [System.Net.HttpListenerException] {
                # Listener was stopped normally
                break
            } catch {
                Write-Log "Error getting request context: $_" "ERROR"
                break
            }
            # Immediately queue the next async receive before processing this context
            $asyncResult = $listener.BeginGetContext($null, $null)

            $request  = $context.Request
            $response = $context.Response

            Write-Log "$($request.HttpMethod) $($request.RawUrl)"

            try {
                # ── Graph subscription validation handshake ───────────────────────────
                # Graph sends a POST (not GET) to the notification URL with
                # ?validationToken=<token> in the query string and an empty body.
                # We must respond 200 with Content-Type: text/plain and the raw token.
                if ($request.QueryString["validationToken"]) {
                    $token = $request.QueryString["validationToken"]
                    Write-Log "Subscription validation handshake received ($($request.HttpMethod)). Echoing token." "SUCCESS"
                    $response.ContentType     = "text/plain"
                    $response.StatusCode      = 200
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($token)
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $response.OutputStream.Flush()
                    $response.Close()
                    continue
                }

                # ── Graph change notification (POST) ──────────────────────────────────
                if ($request.HttpMethod -eq "POST") {

                    # Read body BEFORE closing response
                    $bodyText = ""
                    if ($request.HasEntityBody) {
                        $reader   = [System.IO.StreamReader]::new($request.InputStream, $request.ContentEncoding)
                        $bodyText = $reader.ReadToEnd()
                        $reader.Close()
                    }

                    # Respond 202 Accepted immediately (Graph requires fast response)
                    $response.StatusCode      = 202
                    $response.ContentLength64 = 0
                    $response.Close()

                    if ([string]::IsNullOrWhiteSpace($bodyText)) {
                        Write-Log "POST received with empty body -- ignoring." "WARN"
                        continue
                    }

                    $payload = $null
                    try {
                        $payload = $bodyText | ConvertFrom-Json
                    } catch {
                        Write-Log "Failed to parse notification JSON: $_" "WARN"
                        continue
                    }

                    $valueProp = $payload.PSObject.Properties['value']
                    if (-not $valueProp -or -not $valueProp.Value) {
                        Write-Log "Notification payload has no 'value' array -- ignoring." "WARN"
                        continue
                    }

                    # Reload subscription store in case it was updated
                    $subStore = Load-SubscriptionStore
                    $expectedClientState = $subStore["clientState"]

                    foreach ($notification in $valueProp.Value) {

                        # Validate clientState
                        $csProp = $notification.PSObject.Properties['clientState']
                        $cs     = if ($csProp) { $csProp.Value } else { "" }

                        if ($cs -ne $expectedClientState) {
                            Write-Log "Ignoring notification with unexpected clientState: '$cs'" "WARN"
                            continue
                        }

                        # Extract changed group ID from resourceData.id or resource path
                        $groupId = $null
                        $rdProp  = $notification.PSObject.Properties['resourceData']
                        if ($rdProp -and $rdProp.Value) {
                            $idProp = $rdProp.Value.PSObject.Properties['id']
                            if ($idProp) { $groupId = $idProp.Value }
                        }
                        if (-not $groupId) {
                            $resProp = $notification.PSObject.Properties['resource']
                            if ($resProp -and $resProp.Value -match 'groups/([0-9a-f-]{36})') {
                                $groupId = $Matches[1]
                            }
                        }

                        if (-not $groupId) {
                            Write-Log "Could not extract group ID from notification -- ignoring." "WARN"
                            continue
                        }

                        Write-Log "Notification received for group ID: $groupId"

                        # Ignore notifications for target groups (our own writes would cause loops)
                        if ($excludedIds.Contains($groupId)) {
                            Write-Log "Group '$groupId' is an excluded target group -- ignoring to prevent circular sync." "INFO"
                            continue
                        }

                        $entries = $registry[$groupId]
                        if (-not $entries -or $entries.Count -eq 0) {
                            Write-Log "Group '$groupId' not in registry -- ignoring (not a monitored source group or descendant)." "INFO"
                            continue
                        }

                        # Queue a sync for every source group that contains this group ID
                        # (a child group may belong to multiple source hierarchies)
                        foreach ($entry in @($entries)) {
                            $srcName = $entry["sourceGroup"]
                            $tgtName = $entry["targetGroup"]
                            Write-Log "Mapped to source group: '$srcName' -> target: '$tgtName'"

                            # Store both the timestamp and the target group name so the debounce
                            # loop doesn't need to re-lookup the registry (which is GUID-keyed).
                            $pendingSync[$srcName] = @{ Time = Get-Date; TargetGroup = $tgtName }
                            Write-Log "Sync for '$srcName' queued (debounce: ${DebounceSecs}s)."
                        }
                    }

                    continue
                }

                # ── Anything else: 405 Method Not Allowed ────────────────────────────
                $response.StatusCode = 405
                $response.Close()

            }
            catch {
                Write-Log "Error processing request: $_" "ERROR"
                try { $response.StatusCode = 500; $response.Close() } catch { }
            }
        }

        # Brief sleep to avoid busy-waiting when queue is empty
        Start-Sleep -Milliseconds 100
    }
}
catch [System.OperationCanceledException] {
    Write-Log "Listener cancelled." "WARN"
}
catch {
    Write-Log "Listener error: $_" "ERROR"
}
finally {
    Write-Log "Shutting down..."
    Flush-CompletedJobs
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
    Write-Log "Listener stopped." "SUCCESS"
}
