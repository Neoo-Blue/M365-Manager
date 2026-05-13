# Audit log format (Phase 2 onward)

The session audit log under `%LOCALAPPDATA%\M365Manager\audit\session-*.log` (Windows) or `~/.m365manager/audit/session-*.log` (POSIX) is **JSON-per-line** (NDJSON / JSONL). One complete JSON object per line, no surrounding array. External tooling can parse it with any streaming JSON reader. The AI assistant's `mark-*.log` stays human-readable text — `AuditViewer.ps1`'s parser handles both transparently.

## Field reference

| Field          | Type     | Always? | Meaning |
|----------------|----------|---------|---------|
| `ts`           | string   | yes     | ISO 8601 UTC timestamp, millisecond precision (`2026-05-12T16:23:45.123Z`). |
| `entryId`      | string   | yes     | UUID. Same id is shared across all entries for one logical operation (e.g. PROPOSE then EXEC). |
| `mode`         | string   | yes     | `LIVE` or `PREVIEW`. |
| `event`        | string   | yes     | One of: `SESSION_START`, `PREVIEW`, `EXEC`, `REJECT`, `TODO`, `CLEAR`. |
| `description`  | string   | yes     | Human-readable short string (the `-Description` passed to `Invoke-Action`). |
| `actionType`   | string   | usually | Stable short tag for filtering / undo dispatch (e.g. `AssignLicense`, `AddToGroup`, `BlockSignIn`). Null on legacy lines and `SESSION_START`. |
| `target`       | object   | usually | Free-form hashtable of operands. Keys like `userUpn`, `userId`, `groupId`, `groupName`, `skuId`, `skuPart`, `dlIdentity`, `mailbox`, `identity` are stable contract for downstream tooling. |
| `result`       | string   | yes     | `success`, `failure`, `preview`, `info`. |
| `error`        | string   | when failure | Exception message captured by `Invoke-Action`. |
| `tenant`       | object   | when known | **Phase 6:** structured hashtable `{ name; id; domain; mode }` from `$script:SessionState`. `mode` is `Direct`, `Partner`, or `Profile`. Legacy entries (Phase 1-5) carry this field as a single string; `AuditViewer.ps1`'s filter handles both shapes via `Filter.Tenant` substring match. |
| `session`      | int      | yes     | OS PID of the M365 Manager process that wrote the line. |
| `reverse`      | object   | optional | `{ type, description, target }` describing the inverse operation. `null` when the operation has no curated reverse. |
| `noUndoReason` | string   | optional | Human-readable explanation of why this entry is non-reversible. Set on destructive operations (user / mailbox / group deletion, session revocation, MFA method revocation). |

## Reversibility contract

When `reverse` is non-null:

- `reverse.type` MUST match a key in `$script:UndoHandlers` (see `Undo.ps1`).
- `reverse.target` MUST contain every key the handler expects (e.g. `RemoveFromGroup` needs `userId` + `groupId`).
- `Invoke-Undo -EntryId X` dispatches to `$script:UndoHandlers[reverse.type]` with `reverse.target`. The reversal itself becomes a new audit entry (with its own `entryId`), and `audit/undo-state.json` records that the original entry has been reversed so it won't appear again in `Show-RecentUndoable`.

When `noUndoReason` is set, `reverse` is forced to `null`.

`result` must be `success` for an entry to be undoable — failed and preview entries are skipped by `Get-UndoableEntries`.

## Example lines

A reversible success (license assignment from a template):

```json
{"ts":"2026-05-12T16:23:45.123Z","entryId":"a1b2c3d4-...","mode":"LIVE","event":"EXEC","description":"Assign license 'SPE_E3' to user 4f3a-...","actionType":"AssignLicense","target":{"userId":"4f3a-...","skuId":"05e9...","skuPart":"SPE_E3"},"result":"success","error":null,"tenant":"contoso.onmicrosoft.com","session":17432,"reverse":{"type":"RemoveLicense","description":"Remove license 'SPE_E3' from user 4f3a-...","target":{"userId":"4f3a-...","skuId":"05e9...","skuPart":"SPE_E3"}},"noUndoReason":null}
```

A non-reversible destructive op:

```json
{"ts":"2026-05-12T16:30:11.005Z","entryId":"b2c3d4e5-...","mode":"LIVE","event":"EXEC","description":"DELETE security group 'SG-Old-Marketing'","actionType":"DeleteSecurityGroup","target":{"groupId":"g-id","groupName":"SG-Old-Marketing"},"result":"success","error":null,"tenant":"contoso.onmicrosoft.com","session":17432,"reverse":null,"noUndoReason":"Security group deletion is irreversible (cannot reconstruct membership history)."}
```

A preview entry (dry-run):

```json
{"ts":"2026-05-12T16:31:00.118Z","entryId":"c3d4e5f6-...","mode":"PREVIEW","event":"PREVIEW","description":"Block sign-in for jane@contoso.com","actionType":"BlockSignIn","target":{"userId":"4f3a-...","userUpn":"jane@contoso.com"},"result":"preview","error":null,"tenant":"contoso.onmicrosoft.com","session":17432,"reverse":{"type":"UnblockSignIn","description":"Unblock sign-in for jane@contoso.com","target":{"userId":"4f3a-...","userUpn":"jane@contoso.com"}},"noUndoReason":null}
```

A failure (Graph 403):

```json
{"ts":"2026-05-12T16:32:42.310Z","entryId":"d4e5f6a7-...","mode":"LIVE","event":"EXEC","description":"Add user 4f3a-... to security group 'SG-Engineering'","actionType":"AddToGroup","target":{"userId":"4f3a-...","groupId":"g-id","groupName":"SG-Engineering"},"result":"failure","error":"Insufficient privileges to complete the operation.","tenant":"contoso.onmicrosoft.com","session":17432,"reverse":null,"noUndoReason":null}
```

## Legacy human-readable lines

Pre-Phase-2 session logs (and all AI `mark-*.log` entries) look like:

```
[2026-05-12 09:14:03.124] [PREVIEW] [MODE=PREVIEW] Create user jane.smith@contoso.com
```

`ConvertFrom-AuditLine` recognizes both shapes and produces the same normalized PSCustomObject. `source` field is set to `jsonl`, `legacy-session`, or `ai-mark` so callers can tell them apart.
