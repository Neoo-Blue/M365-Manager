# Persistent chat sessions (Phase 5)

Every chat is auto-saved on `/quit` (unless `/ephemeral` is on) so
multi-day investigations don't lose context, and you can hand a
saved session to a coworker via `/export`. The on-disk blob is
DPAPI-encrypted on Windows so a stolen laptop disk doesn't leak the
operator's privacy map.

## Layout

```
<stateDir>/chat-sessions/
  index.json                    -- unencrypted; id / title / model /
                                   tenant / cost / msg-count
  20260512-093011-a1b2c3d4.session  -- DPAPI-encrypted JSON blob
  ...
```

`<stateDir>` resolves to `%LOCALAPPDATA%\M365Manager` on Windows or
`~/.m365manager` (chmod 700) on POSIX.

`.session` blob (after decryption):

```json
{
  "schemaVersion": 1,
  "id": "20260512-093011-a1b2c3d4",
  "title": "audit guest cleanup contoso",
  "createdUtc": "2026-05-12T13:30:11Z",
  "lastUpdatedUtc": "2026-05-12T13:42:55Z",
  "provider": "Anthropic",
  "model": "claude-opus-4-7",
  "tenant": "contoso.onmicrosoft.com",
  "costRolledUpUsd": 0.4133,
  "privacyMap": {
    "ByValue":  { "alice@contoso.com": "<UPN_1>", ... },
    "ByToken":  { "<UPN_1>": "alice@contoso.com", ... },
    "Counters": { "UPN": 3, "GUID": 7, ... }
  },
  "history": [ { "role": "user", "content": "..." }, ... ]
}
```

`index.json` mirrors metadata but never contains chat content or the
privacy map, so `/list` is a single file read.

## Commands

| Command                       | What it does                                                          |
|-------------------------------|-----------------------------------------------------------------------|
| `/list`                       | Compact table of saved sessions, newest first. Marks current.        |
| `/load <id-or-prefix>`        | Loads by id exact match, then by title prefix.                       |
| `/save [title]`               | Persists the current chat. Title auto-derived from first user msg if omitted. |
| `/rename <id-or-prefix> <new title>` | Renames in place (re-encrypts).                                |
| `/delete <id-or-prefix>`      | Deletes blob + index entry.                                          |
| `/ephemeral`                  | Marks the current chat no-save; suppresses auto-save on `/quit`.     |
| `/export <id> [path]`         | Writes a **redacted** plaintext JSON (placeholders, not real values) for sharing. |

`/quit` calls `Save-AISession` unless `/ephemeral` was toggled on.

## Encryption

- **Windows**: `[System.Security.Cryptography.ProtectedData]::Protect`
  with `DataProtectionScope.CurrentUser`. Only the current user
  account on the current machine can decrypt. Domain roaming profiles
  carry the key with the user.
- **POSIX / failure**: falls back to plaintext with a warning.
  Operators who care about confidentiality on Mac/Linux should
  combine `/ephemeral` with no-disk telemetry or use full-disk
  encryption.

The privacy map is stored INSIDE the encrypted blob so a leaked
`.session` file isn't directly readable. The unencrypted `index.json`
only has metadata (tenant, msg count, USD spend).

## /export

`/export` writes a JSON file with `exportType: "redacted"` and pushes
every string through `Convert-ToSafePayload`. Real UPNs, GUIDs,
tenant IDs, secrets, and names become placeholders (`<UPN_1>`,
`<GUID_3>`, etc.). The privacy map is INTENTIONALLY excluded so the
exported file is safe to email or attach to a Jira ticket.

## Caveats

- **No cross-tenant restore**: the privacy map is per-tenant. Loading
  a session under a different tenant will produce wrong restorations
  for any placeholder whose real value doesn't exist there. The
  schema doesn't prevent this; operators must `/clear` first.
- **Schema version**: `schemaVersion: 1` is the only version. Future
  versions will offer a migration path.
- **No autosave throttle**: every `/quit` saves regardless of
  dirtiness. The dirty flag exists for a future "save before reaching
  context limit" autosave but isn't wired to that path yet.
