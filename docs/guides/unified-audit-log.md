# Unified audit log

Search the Microsoft Purview unified audit log for resource-level events: mail sends, file downloads, sharing, permission changes, configuration changes, etc.

## Menu

**Slot 14 → Audit & Reporting → Unified audit log search**

Or directly:

```powershell
Search-UAL -UserId alice@contoso.com -From (Get-Date).AddDays(-1) -Operations @('FileDownloaded','MailItemsAccessed')
```

## UAL vs sign-in log

| Question | Tool |
|---|---|
| Who signed in, when, from where, with what risk score? | [`sign-in-lookup.md`](sign-in-lookup.md) (`Search-SignIns`) |
| What did they DO once signed in? | This doc (`Search-UAL`) |

UAL covers Exchange + SharePoint + Teams + Power Apps + every M365 workload that emits to the unified audit pipeline. Sign-ins are NOT in UAL — they're in `Search-SignIns`.

## Prereqs

- UAL must be enabled in the tenant. Compliance Center → Audit → "Start recording user and admin activities" — if you see that button instead of search results, UAL is off.
- Role: `Audit Logs` (compliance) or `View-Only Audit Logs`.
- Retention: 90 days by default; up to 1 year with E5 or audit add-on.

## Filter parameters

| Parameter | Notes |
|---|---|
| `-UserId` | UPN. Multiple users not supported in one call. |
| `-From` / `-To` | `[DateTime]` UTC. Default last 24h. |
| `-Operations` | Array of operation names. Wildcards supported. |
| `-IP` | Source IP filter. |

## Operation groups

The tool defines a few convenience groups (in `UnifiedAuditLog.ps1`):

```powershell
$script:UALOperationGroups = @(
    @{ Group='Mail';       Ops=@('Send','MailboxLogin','AddFolderPermissions','RemoveFolderPermissions','UpdateInboxRules','Set-Mailbox','New-InboxRule','Set-MailboxAutoReplyConfiguration') },
    @{ Group='SharePoint'; Ops=@('FileDownloaded','FileUploaded','FileAccessed','FileDeleted','SharingSet','AnonymousLinkCreated','SecureLinkCreated','PermissionLevelModified') },
    @{ Group='Teams';      Ops=@('MemberAdded','MemberRemoved','OwnerAdded','OwnerRemoved','ChannelAdded','ChannelDeleted','MessageSent') },
    @{ Group='Identity';   Ops=@('Add user.','Change user password.','Update user.','Disable account.','Enable account.','Add member to role.','Remove member from role.') },
    @{ Group='Apps';       Ops=@('Consent to application.','Add app role assignment grant to user.') }
)
```

Pass `-Operations @('SharePoint')` and the tool expands to every SharePoint op. Or pass exact op names.

## Common queries

### Files Alice downloaded in the last 24h

```powershell
Search-UAL -UserId alice@contoso.com -From (Get-Date).AddDays(-1) -Operations @('FileDownloaded') | Format-Table CreationDate, Operation, ObjectId
```

### Inbox-rule changes across the tenant in the last week

```powershell
Search-UAL -From (Get-Date).AddDays(-7) -Operations @('New-InboxRule','Set-InboxRule','Remove-InboxRule')
```

### Anyone consenting to a new app

```powershell
Search-UAL -From (Get-Date).AddDays(-30) -Operations @('Consent to application.')
```

### Anonymous link creation

```powershell
Search-UAL -From (Get-Date).AddDays(-7) -Operations @('AnonymousLinkCreated')
```

## Output shape

UAL rows are JSON-rich and vary by operation. Common fields:

| Field | Meaning |
|---|---|
| `CreationDate` | UTC timestamp. |
| `UserIds` | Comma-separated UPNs. |
| `Operation` | Op name. |
| `ResultStatus` | `Succeeded`, `Failed`. |
| `Workload` | `Exchange`, `SharePoint`, `MicrosoftTeams`, `AzureActiveDirectory`. |
| `ObjectId` | Resource id (file URL, group id, etc.). |
| `AuditData` | Operation-specific JSON. Cast / inspect per-operation. |

The tool pipes through `Search-UnifiedAuditLog` from the EXO module, so the row shape matches that cmdlet's output.

## Used by

- [`incident-response.md`](incident-response.md) step 8 (Audit24h captures all operations for the user).
- `Detect-MassFileDownload` — looks at `FileDownloaded` counts.
- `Detect-SuspiciousInboxRule` — reads the rule directly via Graph; UAL is the historical record of when it was created.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `The term 'Search-UnifiedAuditLog' is not recognized` | EXO module not loaded or not connected. Run `Connect-ExchangeOnline` first. The tool's `Connect-ForTask 'Audit'` does this. |
| `UAL not enabled for this tenant` | Enable in Compliance Center → Audit. May take up to 24h before events start landing. |
| `Returns empty for a recent event` | UAL has a 30-minute to 4-hour ingestion lag. Wait and retry. |
| `Too many results, narrow your query` | Use a tighter date range or specific operations. |

## See also

- [`sign-in-lookup.md`](sign-in-lookup.md) — the sibling for authentication events.
- [`incident-response.md`](incident-response.md) step 8 — automated 24h UAL slice per incident.
- [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) — `Detect-MassFileDownload` consumes UAL.
