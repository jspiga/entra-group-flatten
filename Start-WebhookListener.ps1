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
    [int]$Port                   = 8080,
    [string]$ConfigPath          = (Join-Path $PSScriptRoot "config.json"),
    [string]$SubscriptionStorePath = (Join-Path $PSScriptRoot "state/subscriptions.json"),
    [string]$RegistryPath        = (Join-Path $PSScriptRoot "state/group-registry.json"),
    [string]$DeltaSyncScriptPath = (Join-Path $PSScriptRoot "Invoke-DeltaSync.ps1"),
    [int]$DebounceSecs           = 30,
    [string]$LogPath             = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
    $reg = @{}
    if (Test-Path $RegistryPath) {
        $raw = Get-Content $RegistryPath -Raw | ConvertFrom-Json
        $raw.PSObject.Properties | ForEach-Object {
            $entry = @{}
            $_.Value.PSObject.Properties | ForEach-Object { $entry[$_.Name] = $_.Value }
            $reg[$_.Name] = $entry
        }
    }
    return $reg
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
        [string]$AbsDeltaSyncScript
    )

    # If a job is already running for this group, skip
    if ($activeJobs.ContainsKey($SourceGroupName)) {
        $existing = $activeJobs[$SourceGroupName]
        if ($existing.State -eq 'Running') {
            Write-Log "Sync job for '$SourceGroupName' already running -- skipping duplicate trigger." "WARN"
            return
        }
    }

    Write-Log "Starting sync job: '$SourceGroupName' -> '$TargetGroupName'"
    $job = Start-Job -ScriptBlock {
        param($script, $src, $tgt, $cfg)
        & $script -SourceGroup $src -TargetGroup $tgt -ConfigPath $cfg 2>&1
    } -ArgumentList $AbsDeltaSyncScript, $SourceGroupName, $TargetGroupName, $AbsConfigPath

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
$subStore = Load-SubscriptionStore
$registry = Load-Registry
Write-Log "Subscription store loaded (clientState: $($subStore['clientState']))."
Write-Log "Group registry loaded: $($registry.Count) group(s) monitored."

if ($registry.Count -eq 0) {
    Write-Log "WARNING: Group registry is empty. Run Register-ChangeNotification.ps1 first." "WARN"
}

# ── Start HTTP listener ───────────────────────────────────────────────────────

$listener = [System.Net.HttpListener]::new()

# http://+:port/ requires admin privileges (URL ACL). Fall back to localhost if not elevated.
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$prefix  = if ($isAdmin) { "http://+:$Port/" } else { "http://localhost:$Port/" }

if (-not $isAdmin) {
    Write-Log "Not running as Administrator -- listening on localhost only. Use 'netsh http add urlacl url=http://+:$Port/ user=$env:USERNAME' to enable public binding." "WARN"
}

$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
    Write-Log "Listener started on $prefix -- press Ctrl+C to stop." "SUCCESS"
}
catch {
    Write-Log "Failed to start listener on port ${Port}: $_" "ERROR"
    throw
}

# ── Main loop ─────────────────────────────────────────────────────────────────

try {
    while ($listener.IsListening) {

        # Non-blocking context check with 500ms timeout so we can flush jobs
        $asyncResult = $listener.BeginGetContext($null, $null)
        $signalled   = $asyncResult.AsyncWaitHandle.WaitOne(500)

        # Housekeeping: flush completed jobs regardless of whether a request came in
        Flush-CompletedJobs

        # Check for pending debounced syncs that are now due
        $now = Get-Date
        foreach ($srcName in @($pendingSync.Keys)) {
            $pendingTime = $pendingSync[$srcName]
            if (($now - $pendingTime).TotalSeconds -ge $DebounceSecs) {
                $pendingSync.Remove($srcName)
                # Re-load registry to get target group
                $regEntry = $registry[$srcName]   # keyed by srcName here
                if ($regEntry) {
                    Start-SyncJob -SourceGroupName $srcName `
                                  -TargetGroupName $regEntry["targetGroup"] `
                                  -AbsConfigPath $absConfigPath `
                                  -AbsDeltaSyncScript $absDeltaSyncScript
                }
            }
        }

        if (-not $signalled) { continue }

        # Retrieve the context
        $context  = $listener.EndGetContext($asyncResult)
        $request  = $context.Request
        $response = $context.Response

        Write-Log "$($request.HttpMethod) $($request.RawUrl)"

        try {
            # ── Graph subscription validation handshake (GET with validationToken) ──
            if ($request.HttpMethod -eq "GET" -and $request.QueryString["validationToken"]) {
                $token = $request.QueryString["validationToken"]
                Write-Log "Subscription validation handshake received. Echoing token." "SUCCESS"
                $response.ContentType    = "text/plain"
                $response.StatusCode     = 200
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
                $response.StatusCode = 202
                $response.ContentLength64 = 0
                $response.Close()

                # Parse and process notifications
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

                # Reload registry and store in case they were updated
                $registry = Load-Registry
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

                    # Extract the changed resource ID (group ID)
                    # Graph sends resourceData.id or resource path like "groups/{id}"
                    $groupId = $null
                    $rdProp  = $notification.PSObject.Properties['resourceData']
                    if ($rdProp -and $rdProp.Value) {
                        $idProp  = $rdProp.Value.PSObject.Properties['id']
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

                    # Look up in registry
                    $entry = $registry[$groupId]
                    if (-not $entry) {
                        Write-Log "Group '$groupId' not in registry -- not monitored, ignoring." "INFO"
                        continue
                    }

                    $srcName = $entry["sourceGroup"]
                    $tgtName = $entry["targetGroup"]
                    Write-Log "Mapped to source group: '$srcName' -> target: '$tgtName'"

                    # Per-source-group debounce: record pending sync time
                    # If already pending, refresh the timer (coalesces rapid notifications)
                    $pendingSync[$srcName] = Get-Date
                    Write-Log "Sync for '$srcName' queued (debounce: ${DebounceSecs}s)."
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
