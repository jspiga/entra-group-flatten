<#
.SYNOPSIS
    Version 2: Starts a local HTTP listener that receives Microsoft Graph change notifications
    and triggers Sync-FlattenedGroup.ps1 (or Invoke-DeltaSync.ps1 for V3) on receipt.

.DESCRIPTION
    Microsoft Graph sends an HTTPS POST to your notification URL when a subscribed group changes.
    This script runs a simple HttpListener on the specified port to receive those notifications.

    On first contact with a new subscription, Graph sends a validation request (GET with
    validationToken query param) which must be echoed back as plain text within 10 seconds.
    This script handles that handshake automatically.

    After validation, incoming notifications are processed asynchronously:
      - The clientState is verified against the expected pattern.
      - Sync-FlattenedGroup.ps1 is invoked for the affected source/target group pair.
      - To use V3 delta sync instead, set -UseDeltaSync.

    NOTE: For production use this script should run behind a reverse proxy (e.g. nginx, IIS,
    Azure Application Gateway) that terminates TLS. Graph requires HTTPS for notification URLs.
    For local testing, use a tunnel tool such as ngrok.

.PARAMETER ListenPrefix
    The HTTP prefix to listen on. Default: http://+:8080/notify/
    For HTTPS, configure a cert binding via netsh and use https://+:443/notify/.

.PARAMETER SourceGroup
    Display name of the source group (passed to the sync script).

.PARAMETER TargetGroup
    Display name of the target group (passed to the sync script).

.PARAMETER TenantId
    Azure AD tenant ID. Overrides config.json if provided.

.PARAMETER ClientId
    App Registration client ID. Overrides config.json if provided.

.PARAMETER ClientSecret
    App Registration client secret. Overrides config.json if provided.

.PARAMETER ConfigPath
    Path to config.json. Defaults to ./config.json relative to the script location.

.PARAMETER UseDeltaSync
    If specified, triggers Invoke-DeltaSync.ps1 (V3) instead of Sync-FlattenedGroup.ps1 (V1).

.EXAMPLE
    # Start listener and trigger V1 sync on notifications
    .\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

.EXAMPLE
    # Start listener and trigger V3 delta sync on notifications
    .\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -UseDeltaSync
#>

[CmdletBinding()]
param(
    [string]$ListenPrefix = "http://+:8080/notify/",

    [Parameter(Mandatory)]
    [string]$SourceGroup,

    [Parameter(Mandatory)]
    [string]$TargetGroup,

    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret,

    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json"),

    [switch]$UseDeltaSync
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

# ── Load config ───────────────────────────────────────────────────────────────

. (Join-Path $PSScriptRoot "lib/GraphAuth.ps1")
. (Join-Path $PSScriptRoot "lib/GraphGroups.ps1")

$config = @{}
if (Test-Path $ConfigPath) {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json | ForEach-Object { $h = @{}; $_.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h }
}

if (-not $TenantId)     { $TenantId     = $config["tenantId"]     }
if (-not $ClientId)     { $ClientId     = $config["clientId"]     }
if (-not $ClientSecret) { $ClientSecret = $config["clientSecret"] }
$prefix = $config["targetGroupPrefix"]

# Resolve target group name (apply prefix if needed)
$TargetGroup = Resolve-TargetGroupName -Name $TargetGroup -Prefix $prefix

# Determine which sync script to invoke on notification
$syncScript = if ($UseDeltaSync) {
    Join-Path $PSScriptRoot "Invoke-DeltaSync.ps1"
} else {
    Join-Path $PSScriptRoot "Sync-FlattenedGroup.ps1"
}

# ── Start HTTP listener ───────────────────────────────────────────────────────

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($ListenPrefix)

try {
    $listener.Start()
    Write-Log "Webhook listener started on $ListenPrefix" "SUCCESS"
    Write-Log "Press Ctrl+C to stop."

    # Track last sync time to debounce rapid successive notifications
    $lastSyncTime = [datetime]::MinValue
    $debounceSecs = 10

    while ($listener.IsListening) {
        # GetContext() blocks until a request arrives
        $contextTask = $listener.GetContextAsync()

        # Poll with a short timeout so Ctrl+C is responsive
        while (-not $contextTask.IsCompleted) {
            Start-Sleep -Milliseconds 200
        }

        $context  = $contextTask.Result
        $request  = $context.Request
        $response = $context.Response

        try {
            # ── Graph validation handshake ────────────────────────────────────
            # When a new subscription is created, Graph sends a GET with ?validationToken=...
            if ($request.HttpMethod -eq "GET" -and $request.QueryString["validationToken"]) {
                $token = $request.QueryString["validationToken"]
                Write-Log "Received validation request -- echoing token."
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($token)
                $response.ContentType   = "text/plain"
                $response.StatusCode    = 200
                $response.ContentLength64 = $bytes.Length
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
                Write-Log "Validation handshake complete." "SUCCESS"
            }
            # ── Incoming change notification ──────────────────────────────────
            elseif ($request.HttpMethod -eq "POST") {
                # Graph expects a 202 Accepted response within a few seconds
                $response.StatusCode = 202
                $response.Close()

                # Read notification body (don't block response)
                $reader  = [System.IO.StreamReader]::new($request.InputStream)
                $body    = $reader.ReadToEnd()
                $reader.Dispose()

                Write-Log "Change notification received."
                Write-Verbose "Notification body: $body"

                # Parse and verify clientState
                try {
                    $notification = $body | ConvertFrom-Json
                    $validNotification = $false

                    foreach ($item in $notification.value) {
                        $cs = $item.clientState
                        if ($cs -and $cs.StartsWith("entra-group-flatten-")) {
                            $validNotification = $true
                            break
                        }
                    }

                    if (-not $validNotification) {
                        Write-Log "Notification ignored -- clientState does not match expected pattern." "WARN"
                        continue
                    }
                }
                catch {
                    Write-Log "Failed to parse notification body: $_" "WARN"
                    continue
                }

                # Debounce: skip if a sync ran very recently
                $now = [datetime]::UtcNow
                if (($now - $lastSyncTime).TotalSeconds -lt $debounceSecs) {
                    Write-Log "Debounce: skipping sync (last sync was $([int]($now - $lastSyncTime).TotalSeconds)s ago)" "WARN"
                    continue
                }
                $lastSyncTime = $now

                # Invoke sync script as a background job so the listener loop stays responsive
                Write-Log "Triggering sync: $syncScript"
                $jobArgs = @{
                    FilePath         = $syncScript
                    ArgumentList     = @(
                        "-SourceGroup", $SourceGroup,
                        "-TargetGroup", $TargetGroup,
                        "-TenantId",    $TenantId,
                        "-ClientId",    $ClientId,
                        "-ClientSecret",$ClientSecret
                    )
                }
                $job = Start-Job -ScriptBlock {
                    param($fp, $args)
                    & $fp @args
                } -ArgumentList $jobArgs.FilePath, $jobArgs.ArgumentList

                Write-Log "Sync job started (Job ID: $($job.Id))" "SUCCESS"

                # Clean up completed jobs (non-blocking)
                Get-Job -State Completed | ForEach-Object {
                    $output = Receive-Job $_
                    if ($output) { Write-Log "Job $($_.Id) output: $output" }
                    Remove-Job $_
                }
            }
            else {
                $response.StatusCode = 405
            }
        }
        catch {
            Write-Log "Error handling request: $_" "ERROR"
            try { $response.StatusCode = 500 } catch {}
        }
        finally {
            try { $response.Close() } catch {}
        }
    }
}
catch [System.OperationCanceledException] {
    Write-Log "Listener cancelled." "WARN"
}
catch {
    Write-Log "Fatal listener error: $_" "ERROR"
}
finally {
    if ($listener.IsListening) { $listener.Stop() }
    $listener.Close()
    Write-Log "Webhook listener stopped."
}
