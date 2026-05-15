# Audit & undo

Every mutation in the tool writes an audit entry; most entries are reversible via the undo system. This guide covers reading the audit log, filtering it, and reversing operations.

## Menu

**Slot 14 → Audit & Reporting**

```
  1. Open audit log viewer (interactive)
  2. Filter audit log
  3. Export filtered entries to CSV
  4. Export filtered entries to HTML
  5. List recent undoable operations
  6. Undo a specific operation (by entry id)
  7. View undo state (already-reversed entries)
```

## Where logs live

| File | Contents |
|---|---|
| `%LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>-<tenant>.log` | One JSONL line per `Invoke-Action`. |
| `%LOCALAPPDATA%\M365Manager\audit\mark-<ts>.log` | AI assistant per-command audit (separate from session log). |
| `<stateDir>\undo-state.json` | Map of already-reversed entry IDs → reversal metadata. |

On non-Windows, replace `%LOCALAPPDATA%\M365Manager` with `~/.m365manager`. The directory is `chmod 700` on POSIX.

## Reading entries

Every line is one complete JSON object. Full schema at [`../reference/audit-format.md`](../reference/audit-format.md). The key fields:

| Field | Use |
|---|---|
| `entryId` | UUID. Pass to `Invoke-Undo -EntryId X`. |
| `ts` | ISO 8601 UTC. |
| `event` | `PROPOSE` (PREVIEW) / `EXEC` / `OK` / `ERROR` / `UNDO`. |
| `actionType` | Stable filter key. E.g. `BlockSignIn`, `AssignLicense`, `Incident:RevokeAuthMethods`. |
| `target` | Structured operand hashtable. E.g. `{ userUpn, userId, groupId, groupName }`. |
| `result` | `success` / `failure` / `preview` / `info`. |
| `reverse` | `{type, description, target}` for reversible ops; `null` otherwise. |
| `noUndoReason` | Set on irreversible ops with the why. |
| `tenant` | Phase 6: structured `{name, id, domain, mode}`. Older entries have a string. |

## Interactive viewer

**Slot 14 → option 1.**

```
+-- AUDIT LOG VIEWER -- session-2026-05-14_173208-1234-contoso.log -----------+
| Filter: (none)                                                              |
| 247 entries  page 1/5                                                       |
+-----------------------------------------------------------------------------+
| Time     Event    ActionType         Target                       Result   |
|---------|--------|-------------------|------------------------------|---------|
| 17:42:11 EXEC     BlockSignIn        alice@contoso.com             success  |
| 17:42:11 EXEC     RevokeSessions     alice@contoso.com             success  |
| 17:42:12 EXEC     Incident:RevokeAu  alice@contoso.com             success  |
| ...                                                                         |
+-----------------------------------------------------------------------------+

  [n]ext page  [p]rev page  [f]ilter  [d]etail  [e]xport  [u]ndo  [q]uit
```

`[d]etail` prints the full JSON of the selected row. `[u]ndo` runs `Invoke-Undo` on the row if it's reversible.

## Filtering

Filter via:

- **Mode**: `LIVE` / `PREVIEW`
- **EventType**: `EXEC`, `OK`, `ERROR`, etc.
- **ActionType**: exact / wildcard match (`BlockSignIn`, `Incident:*`)
- **Result**: `success`, `failure`, `preview`, `info`
- **User**: UPN substring against `target.userUpn` / `target.userPrincipalName`
- **Target**: substring against the full `target` hashtable (catches group names, file names, etc.)
- **Tenant**: name / id / domain
- **Date range**: `From` / `To` DateTime

Programmatic equivalent:

```powershell
$entries = Read-AuditEntries
Filter-AuditEntries -Entries $entries -Filter @{
    Result = 'failure'
    From   = (Get-Date).AddDays(-7)
} | Format-Table
```

## Export

CSV (one row per entry, flattened):

```powershell
$entries | Export-AuditEntriesCsv -Path .\audit-2026-05-14.csv
```

HTML (interactive table with filter / sort):

```powershell
$entries | Export-AuditEntriesHtml -Path .\audit-2026-05-14.html
```

The HTML report renders standalone (no external dependencies) — safe to email or attach to a compliance ticket.

## Undo system

The dispatch table lives in `Undo.ps1`. Today it covers 24 reverse types — the full list at [`../reference/undo-handlers.md`](../reference/undo-handlers.md). Common ones:

| Forward | Reverse |
|---|---|
| `AssignLicense` | `RemoveLicense` |
| `RemoveLicense` | `AssignLicense` |
| `AddToGroup` | `RemoveFromGroup` |
| `RemoveFromGroup` | `AddToGroup` |
| `BlockSignIn` | `UnblockSignIn` |
| `SetForwarding` | `ClearForwarding` |
| `ClearForwarding` | `SetForwarding` |
| `AddSiteOwner` | `RemoveSiteOwner` |
| `GrantOneDriveAccess` | `RevokeOneDriveAccess` |
| Phase 7 `Incident:DisableInboxRule` | `EnableInboxRule` |

### Running an undo

```powershell
# By entry id from the audit log:
Invoke-Undo -EntryId a1b2c3d4-...

# Interactive: pick from the last 20 undoable ops:
Show-RecentUndoable | Out-Host        # lists with index
Invoke-Undo -Index 3                   # run by index
```

`Invoke-Undo` writes its own audit entry (`event=UNDO`, `actionType=Undo-<original>`) and marks the original in `<stateDir>\undo-state.json` so it can't be reversed twice.

### Irreversible operations

Some operations have no curated reverse — they get `noUndoReason` set:

| Operation | Why irreversible |
|---|---|
| User deletion | Restorable from `/directory/deletedItems` for 30 days, but not via `Invoke-Undo`. Audit entry calls this out. |
| Session revocation | Sessions are gone; user must re-auth on next access (which is the desired state). |
| MFA method revocation | No "re-grant" API; user must re-enroll. |
| Password change | Can't restore the prior password (we don't keep it). |
| Compliance purge | Messages are permanently removed from the tenant. |

`Invoke-Undo` on these entries prints the `noUndoReason` and exits without acting.

## Audit retention

The tool does NOT auto-prune audit logs. Operators are expected to integrate with their org's retention policy. For long-term retention, copy `audit/*.log` to a write-once archive monthly.

The `health-checks/health-audit-rotation.ps1` (planned) will surface log files older than the configured retention; for now this is a manual operation.

## Per-tenant logs

Phase 6 stamps each log file with a tenant slug. Switch tenants and the next entries go to a new file. AuditViewer's filter UI distinguishes by `tenant.name` / `tenant.id` / `tenant.domain`. See [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md).

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Could not parse line N: Unexpected character` | A non-JSON line slipped in (rare; usually from a process that wrote partial output during a crash). Skip the line — the viewer continues. |
| `Invoke-Undo: handler not registered for reverseType 'X'` | The audit entry references a reverse type that's no longer registered. Update the handler table or accept the manual reversal. |
| `Already reversed in undo-state.json` | The original entry was undone in a prior session. Audit log retains both PROPOSE / EXEC and UNDO trails. |

## See also

- [`../reference/audit-format.md`](../reference/audit-format.md) — JSONL field reference.
- [`../reference/undo-handlers.md`](../reference/undo-handlers.md) — dispatch table.
- [`../concepts/security-model.md`](../concepts/security-model.md) — audit threat model.
- [`incident-response.md`](incident-response.md) — `Undo-Incident` walks the entire incident's reversible steps.
