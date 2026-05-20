# Entra ID Group Flattening — PowerShell Script Suite

Flattens a nested Entra ID (Azure AD) security group and keeps a target group synchronised with the resulting flat list of users. Built around the [Microsoft Graph REST API](https://learn.microsoft.com/en-us/graph/overview).

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [App Registration Setup](#app-registration-setup)
- [Configuration](#configuration)
- [Version 1 — MVP (Manual / Scheduled)](#version-1--mvp-manual--scheduled)
- [Version 2 — Webhook-Driven Sync](#version-2--webhook-driven-sync)
- [Version 3 — Delta Query Optimised Sync](#version-3--delta-query-optimised-sync)
- [Target Group Naming & Prefix](#target-group-naming--prefix)
- [File Reference](#file-reference)
- [Required API Permissions](#required-api-permissions)
- [Notes & Caveats](#notes--caveats)

---

## Prerequisites

- PowerShell 7.2 or later (Windows PowerShell 5.1 is supported but not recommended)
- An Azure App Registration with the permissions listed in [Required API Permissions](#required-api-permissions)
- Network access to `login.microsoftonline.com` and `graph.microsoft.com`
- For Version 2: a publicly accessible HTTPS endpoint for Graph webhook delivery (e.g. ngrok for local testing, or an Azure Function / reverse proxy in production)

---

## App Registration Setup

1. In the [Azure Portal](https://portal.azure.com), go to **Microsoft Entra ID → Manage → App registrations → New registration**.
2. Give it a name (e.g. `EntraGroupFlattenScript`), select **Accounts in this organisational directory only** or **Single tenant only - Default Directory**, and click **Register**.
3. Note the **Application (client) ID** and **Directory (tenant) ID**.
4. Go to **Manage → Certificates & secrets → New client secret** and create a new secret. Copy the secret value immediately.
5. Go to **Manage → API permissions → Add a permission → Microsoft Graph → Application permissions** and add the permissions listed below. Click **Grant admin consent**.

---

## Configuration

Copy `config.example.json` to `config.json` and fill in your values:

```json
{
  "tenantId": "your-tenant-id",
  "clientId": "your-app-registration-client-id",
  "clientSecret": "your-client-secret",
  "targetGroupPrefix": "ATL_",
  "webhookNotificationUrl": "https://your-endpoint.example.com/notify",
  "stateFilePath": "./state/group-state.json"
}
```

| Field | Required | Description |
|---|---|---|
| `tenantId` | Yes | Azure AD tenant ID (GUID or domain) |
| `clientId` | Yes | App Registration client ID |
| `clientSecret` | Yes | Client secret value |
| `targetGroupPrefix` | No | If set, automatically prepended to target group names that don't already have it (e.g. `ATL_`) |
| `webhookNotificationUrl` | V2 only | Public HTTPS URL that Graph will POST change notifications to |
| `stateFilePath` | V3 only | Path to the JSON state file for delta link persistence (relative to script root or absolute) |

> **Security note:** `config.json` contains sensitive credentials. Do not commit it to source control. Add it to `.gitignore`. Consider using environment variables or a secrets manager in production.

---

## Version 1 — MVP (Manual / Scheduled)

**Script:** `Sync-FlattenedGroup.ps1`

Calls `GET /groups/{id}/transitiveMembers` to obtain a fully flattened list of users from the source group, then adds/removes members in the target group to match.

### Usage

```powershell
# Minimal — reads credentials from config.json
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

# Override credentials inline
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" `
    -TenantId "<tid>" -ClientId "<cid>" -ClientSecret "<secret>"

# Dry run — shows what would change without making any API write calls
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -WhatIf

# Verbose output (shows individual member names)
.\Sync-FlattenedGroup.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -Verbose
```

### Scheduling

To run on a schedule with Windows Task Scheduler:

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument '-NonInteractive -File "C:\scripts\entra-group-flatten\Sync-FlattenedGroup.ps1" -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"'
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "EntraGroupFlatten" -Action $action -Trigger $trigger -RunLevel Highest
```

---

## Version 2 — Webhook-Driven Sync

**Scripts:** `Register-ChangeNotification.ps1` + `Start-WebhookListener.ps1`

Instead of polling on a schedule, Version 2 subscribes to Microsoft Graph change notifications on the source group and all its nested sub-groups. When any group changes, Graph POSTs a notification to your endpoint and the sync runs automatically.

### Step 1 — Start the webhook listener

```powershell
# Triggers V1 (full transitive sync) on each notification
.\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

# Triggers V3 (delta sync) on each notification instead
.\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -UseDeltaSync

# Custom listen port/path
.\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" `
    -ListenPrefix "http://+:9090/graphnotify/"
```

> For local development, use [ngrok](https://ngrok.com/) to expose the listener publicly:
> ```powershell
> ngrok http 8080
> # Copy the https://... forwarding URL → set webhookNotificationUrl in config.json with /notify/ appended
> ```

### Step 2 — Register subscriptions

```powershell
.\Register-ChangeNotification.ps1 -SourceGroup "All-Engineering"
```

Graph change notification subscriptions expire after a maximum of 3 days for group resources. Run this script on a schedule (every 2 days) to renew them:

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument '-NonInteractive -File "C:\scripts\entra-group-flatten\Register-ChangeNotification.ps1" -SourceGroup "All-Engineering"'
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Days 2) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "EntraGroupFlattenSubscribe" -Action $action -Trigger $trigger -RunLevel Highest
```

---

## Version 3 — Delta Query Optimised Sync

**Script:** `Invoke-DeltaSync.ps1`

Builds on Version 2 by using [Microsoft Graph delta queries](https://learn.microsoft.com/en-us/graph/delta-query-overview) to track only changes since the last run. This dramatically reduces API call volume for large groups with infrequent changes.

### How it works

| Run | Behaviour |
|---|---|
| First run | Full member fetch for each group + stores delta links and member snapshots |
| Subsequent runs | Delta query per group → applies changes to tracked snapshot → aggregates → diffs against tracked target → only calls Graph to update target if a real change is detected |

### Usage

```powershell
# Normal run (uses delta links if available, full sync on first run)
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat"

# Force a full re-sync (discards all stored delta state)
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -ForceFullSync

# Custom state file location
.\Invoke-DeltaSync.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" `
    -StateFilePath "C:\data\entra-state.json"
```

Best used with **Version 2** as the trigger:

```powershell
.\Start-WebhookListener.ps1 -SourceGroup "All-Engineering" -TargetGroup "EngineeringFlat" -UseDeltaSync
```

---

## Target Group Naming & Prefix

If `targetGroupPrefix` is set in `config.json` (e.g. `"ATL_"`), the scripts will automatically prepend it to any `-TargetGroup` value that doesn't already start with the prefix:

```
-TargetGroup "EngineeringFlat"   →  resolved to "ATL_EngineeringFlat"
-TargetGroup "ATL_EngineeringFlat"  →  unchanged (prefix already present)
```

A message is printed to confirm the normalisation:
```
[INFO] Target group name normalised to: 'ATL_EngineeringFlat' (prefix 'ATL_' applied)
```

If no prefix is configured, the supplied name is used as-is.

---

## File Reference

```
entra-group-flatten/
├── config.example.json              # Template — copy to config.json and fill in
├── config.json                      # Your credentials (do not commit to source control)
│
├── Sync-FlattenedGroup.ps1          # V1: MVP full sync (manual / scheduled)
├── Register-ChangeNotification.ps1  # V2: Create/renew Graph change notification subscriptions
├── Start-WebhookListener.ps1        # V2: HTTP listener — receives notifications, triggers sync
├── Invoke-DeltaSync.ps1             # V3: Delta query optimised sync
│
├── lib/
│   ├── GraphAuth.ps1                # Shared: OAuth2 client credentials token acquisition + caching
│   ├── GraphGroups.ps1              # Shared: All Graph group operations (get, create, add, remove, delta)
│   └── StateStore.ps1               # V3: JSON state persistence for delta links + member snapshots
│
└── state/                           # Auto-created at runtime
    ├── group-state.json             # V3: Delta link + member snapshot store
    └── subscriptions.json           # V2: Active subscription ID store
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

For Version 2 (webhooks), no additional API permissions are needed — the subscription is created using the same token. However, your notification URL must be publicly reachable over HTTPS.

---

## Notes & Caveats

- **Rate limiting:** All Graph API calls use automatic exponential backoff on HTTP 429 responses. For very large groups (10,000+ members), consider running Version 3 to minimise API call volume.
- **SCIM sync delay:** Even after the target group is updated in Entra ID, downstream SCIM provisioning to applications (e.g. Atlassian Cloud) may take 20–40 minutes to propagate.
- **Webhook HTTPS requirement:** Microsoft Graph will only deliver notifications to HTTPS endpoints with a valid TLS certificate. For production, place the listener behind a reverse proxy or deploy as an Azure Function.
- **Subscription expiry:** Graph change notification subscriptions for group resources expire after at most 3 days. Run `Register-ChangeNotification.ps1` on a schedule (every 2 days) to keep them active.
- **Target group type:** New target groups are created as mail-disabled security groups (`securityEnabled: true`, `mailEnabled: false`). This is the correct type for Atlassian Cloud SCIM group provisioning.
- **State file security (V3):** The state file contains user object IDs. Store it in a location with appropriate access controls.
