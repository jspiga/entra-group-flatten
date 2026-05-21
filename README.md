# Entra ID Group Flattening — PowerShell Script Suite

Flattens a nested Entra ID (Azure AD) security group and keeps a target group synchronised with the resulting flat list of users. Built around the [Microsoft Graph REST API](https://learn.microsoft.com/en-us/graph/overview).

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [App Registration Setup](#app-registration-setup)
- [Configuration](#configuration)
- [Option 1 — Manual / Scheduled Sync](#option-1--manual--scheduled-sync)
- [Option 2 — Delta Query Optimised Sync](#option-2--delta-query-optimised-sync)
- [Option 3 — Webhook-Driven Sync](#option-3--webhook-driven-sync)
- [Target Group Naming & Prefix](#target-group-naming--prefix)
- [File Reference](#file-reference)
- [Required API Permissions](#required-api-permissions)
- [Notes & Caveats](#notes--caveats)

---

## Prerequisites

- PowerShell 7.2 or later (Windows PowerShell 5.1 is supported but not recommended)
- An Azure App Registration with the permissions listed in [Required API Permissions](#required-api-permissions)
- Network access to `login.microsoftonline.com` and `graph.microsoft.com`
- For Option 3: a publicly accessible HTTPS endpoint for Graph webhook delivery (e.g. ngrok for local testing, or an Azure Function / reverse proxy in production)

---

## App Registration Setup

1. In the [Azure Portal](https://portal.azure.com), go to **Azure Active Directory → App registrations → New registration**.
2. Give it a name (e.g. `EntraGroupFlattenScript`), select **Accounts in this organisational directory only**, and click **Register**.
3. Note the **Application (client) ID** and **Directory (tenant) ID**.
4. Go to **Certificates & secrets → New client secret**. Copy the secret value immediately.
5. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions** and add the permissions listed below. Click **Grant admin consent**.

---

## Configuration

Copy `config.example.json` to `config.json` and fill in your values:

```json
{
  "tenantId": "your-tenant-id",
  "clientId": "your-app-registration-client-id",
  "clientSecret": "your-client-secret",
  "targetGroupPrefix": "FLAT_",
  "webhookNotificationUrl": "https://your-endpoint.example.com/notify",
  "stateFilePath": "./state/group-state.json",
  "groupRegistryPath": "./state/group-registry.json",
  "subscriptionStorePath": "./state/subscriptions.json"
}
```

| Field | Required | Description |
|---|---|---|
| `tenantId` | Yes | Azure AD tenant ID (GUID or domain) |
| `clientId` | Yes | App Registration client ID |
| `clientSecret` | Yes | Client secret value |
| `targetGroupPrefix` | No | If set, automatically prepended to target group names that don't already have it (e.g. `FLAT_`) |
| `webhookNotificationUrl` | Option 3 only | Public HTTPS URL that Graph will POST change notifications to |
| `stateFilePath` | Option 2 only | Path to the JSON state file for delta link persistence (relative to script root or absolute) |
| `groupRegistryPath` | Option 3 only | Path to the group registry JSON (maps group IDs to source groups). Default: `./state/group-registry.json` |
| `subscriptionStorePath` | Option 3 only | Path to the subscription store JSON. Default: `./state/subscriptions.json` |

> **Security note:** `config.json` contains sensitive credentials. Do not commit it to source control. Add it to `.gitignore`. Consider using environment variables or a secrets manager in production.

---

## Option 1 — Manual / Scheduled Sync

**Script:** `Sync-FlattenedGroup.ps1`

Calls `GET /groups/{id}/transitiveMembers` to obtain a fully flattened list of users from the source group, then adds/removes members in the target group to match.

### Usage

```powershell
# Minimal — reads credentials from config.json
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering"

# Override credentials inline
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering" `
    -TenantId "<tid>" -ClientId "<cid>" -ClientSecret "<secret>"

# Dry run — shows what would change without making any API write calls
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering" -WhatIf

# Verbose output (shows individual member names)
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering" -Verbose
```

### Scheduling

To run on a schedule with Windows Task Scheduler:

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument '-NonInteractive -ExecutionPolicy ByPass -File "C:\scripts\entra-group-flatten\Sync-FlattenedGroup.ps1" -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering"'
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "EntraGroupFlatten" -Action $action -Trigger $trigger -RunLevel Highest
```

---

## Option 2 — Delta Query Optimised Sync

**Script:** `Invoke-DeltaSync.ps1`

Builds on Option 1 by using [Microsoft Graph delta queries](https://learn.microsoft.com/en-us/graph/delta-query-overview) to track only changes since the last run. This dramatically reduces API call volume for large groups with infrequent changes.

### How it works

| Run | Behaviour |
|---|---|
| First run | Full member fetch for each group + stores delta links and member snapshots |
| Subsequent runs | Delta query per group -> applies changes to tracked snapshot -> aggregates -> diffs against tracked target -> only calls Graph to update target if a real change is detected |

### Usage

```powershell
# Normal run (uses delta links if available, full sync on first run)
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering"

# Force a full re-sync (discards all stored delta state)
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering" -ForceFullSync

# Custom state file location
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "FLAT_All-Engineering" `
    -StateFilePath "C:\data\entra-state.json"
```

Best used with **Option 3** as the trigger -- the webhook listener automatically calls `Invoke-DeltaSync.ps1` when a change notification is received.

---

## Option 3 — Webhook-Driven Sync

**Scripts:** `Register-ChangeNotification.ps1` + `Start-WebhookListener.ps1`

Instead of polling on a schedule, Version 2 uses a **single tenant-wide Microsoft Graph change notification subscription** covering all groups. This scales to any number of source groups without hitting Graph subscription limits. When any monitored group changes, Graph POSTs a notification to your endpoint, the listener looks up the affected group in the registry, and triggers `Invoke-DeltaSync.ps1` for the relevant source group.

### Architecture

```
Graph (one subscription on /groups)
    |
    v
Start-WebhookListener.ps1
    |-- loads state/group-registry.json  (groupId -> sourceGroup mapping)
    |-- validates clientState
    |-- routes notification to correct Invoke-DeltaSync.ps1 job
    |-- per-source-group debounce (coalesces rapid changes)
```

### Step 1 — Register groups and create subscription

Run once per source group (or all at once). Pass multiple groups as an array:

```powershell
# Single group
.\Register-ChangeNotification.ps1 `
    -SourceGroup "All-Engineering" `
    -TargetGroup "FLAT_All-Engineering" `
    -NotificationUrl "https://your-endpoint.example.com/notify"

# Multiple groups (scales to 100+ with a single subscription)
.\Register-ChangeNotification.ps1 `
    -SourceGroup "All-Engineering","All-Finance","All-HR" `
    -TargetGroup "FLAT_All-Engineering","FLAT_All-Finance","FLAT_All-HR" `
    -NotificationUrl "https://your-endpoint.example.com/notify"
```

This creates/updates:
- `state/group-registry.json` -- maps every group ID in every hierarchy to its source group name
- `state/subscriptions.json` -- stores the single subscription ID for renewal

Graph subscriptions on `/groups` expire after a maximum of 3 days. Schedule this script to run every 2 days:

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument '-NonInteractive -ExecutionPolicy ByPass -File "C:\scripts\entra-group-flatten\Register-ChangeNotification.ps1" -SourceGroup "All-Engineering","All-Finance" -TargetGroup "FLAT_All-Engineering","FLAT_All-Finance"'
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Days 2) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "EntraGroupFlattenRenew" -Action $action -Trigger $trigger -RunLevel Highest
```

### Step 2 — Start the webhook listener

```powershell
# Basic
.\Start-WebhookListener.ps1

# Custom port with file logging
.\Start-WebhookListener.ps1 -Port 9090 -LogPath ".\logs\webhook.log"

# All options
.\Start-WebhookListener.ps1 -Port 8080 -DebounceSecs 60 -LogPath ".\logs\webhook.log"
```

The listener automatically reads `state/group-registry.json` to route notifications. To monitor additional groups, re-run `Register-ChangeNotification.ps1` (the listener reloads the registry on each notification).

> For local development, use [ngrok](https://ngrok.com/) to expose the listener publicly:
> ```powershell
> ngrok http 8080
> # Copy the https://... forwarding URL and set webhookNotificationUrl in config.json
> ```

---

## Target Group Naming & Prefix

If `targetGroupPrefix` is set in `config.json` (e.g. `"FLAT_"`), the scripts will automatically prepend it to any `-TargetGroup` value that doesn't already start with the prefix:

```
-TargetGroup "All-Engineering"        ->  resolved to "FLAT_All-Engineering"
-TargetGroup "FLAT_All-Engineering"   ->  unchanged (prefix already present)
```

A message is printed to confirm the normalisation:
```
[INFO] Target group name normalised to: 'FLAT_All-Engineering' (prefix 'FLAT_' applied)
```

If no prefix is configured, the supplied name is used as-is.

---

## File Reference

```
entra-group-flatten/
├── config.example.json              # Template — copy to config.json and fill in
├── config.json                      # Your credentials (do not commit to source control)
│
├── Sync-FlattenedGroup.ps1          # Option 1: Full sync (manual / scheduled)
├── Invoke-DeltaSync.ps1             # Option 2: Delta query optimised sync
├── Register-ChangeNotification.ps1  # Option 3: Create/renew single Graph subscription + build registry
├── Start-WebhookListener.ps1        # Option 3: HTTP listener -- receives notifications, triggers sync
│
├── lib/
│   ├── GraphAuth.ps1                # Shared: OAuth2 client credentials token acquisition + caching
│   ├── GraphGroups.ps1              # Shared: All Graph group operations (get, create, add, remove, delta)
│   └── StateStore.ps1               # Option 2/3: JSON state persistence for delta links + member snapshots
│
└── state/                           # Auto-created at runtime
    ├── group-state.json             # Option 2: Delta link + member snapshot store
    ├── group-registry.json          # Option 3: Group ID -> source group mapping (built by Register-ChangeNotification.ps1)
    └── subscriptions.json           # Option 3: Active subscription ID store
```

---

## Required API Permissions

These are **Application** permissions (not Delegated) — required for the client credentials flow.

| Permission | Reason |
|---|---|
| `Group.Read.All` | Look up groups by name, read members, read transitive members |
| `Group.ReadWrite.All` | Create target group if missing, add/remove members |
| `GroupMember.ReadWrite.All` | Add/remove direct members of a group |
| `Directory.Read.All` | Read directory objects (used by transitive member queries) |

> `Group.ReadWrite.All` implies `Group.Read.All`. You may be able to combine these depending on your organisation's consent policies.

For Option 3 (webhooks), no additional API permissions are needed -- the subscription is created using the same token. However, your notification URL must be publicly reachable over HTTPS.

---

## Notes & Caveats

- **Rate limiting:** All Graph API calls use automatic exponential backoff on HTTP 429 responses. For very large groups (10,000+ members), consider using Option 2 to minimise API call volume.
- **SCIM sync delay:** Even after the target group is updated in Entra ID, downstream SCIM provisioning to applications (e.g. Atlassian Cloud) may take 20--40 minutes to propagate.
- **Webhook HTTPS requirement:** Microsoft Graph will only deliver notifications to HTTPS endpoints with a valid TLS certificate. For production, place the listener behind a reverse proxy or deploy as an Azure Function.
- **Subscription expiry:** Graph change notification subscriptions for group resources expire after at most 3 days. Run `Register-ChangeNotification.ps1` on a schedule (every 2 days) to keep them active.
- **Target group type:** New target groups are created as mail-disabled security groups (`securityEnabled: true`, `mailEnabled: false`). This is the correct type for Atlassian Cloud SCIM group provisioning.
- **State file security (Option 2/3):** The state files contain user object IDs and subscription metadata. Store them in a location with appropriate access controls.
