# AI chat slash commands

Every slash command supported inside the AI assistant (menu slot 99). Inventory generated from `AIAssistant.ps1`.

## Setup & state

| Command | Behavior |
|---|---|
| `/help` | Print this command list. |
| `/about` | Diagnostic snapshot: provider/model, plan mode, cost totals, audit / session / cost dirs, plus current preview state. |
| `/config` | Re-run the AI provider setup wizard (provider, model, endpoint, API key). |
| `/models` | List available models for the current provider (where the provider exposes a list endpoint). |
| `/api` | Show / set the API endpoint URL for the current provider. |
| `/context` | Print the current conversation context (token count, message count, tenant). |
| `/clear` | Reset the conversation history + privacy map. Same as starting fresh. |
| `/privacy` | Configure PII redaction settings (`ExternalRedaction` / `RedactInAuditLog` / `ExternalPayloadCapBytes` / `TrustedProviders`). |

## Mode control

| Command | Behavior |
|---|---|
| `/dryrun` | Toggle PREVIEW mode for the current session. Status banner flips color. |
| `/plan` | Force plan mode for the next prompt — the AI must submit_plan before any tool. |
| `/noplan` | Disable plan mode for the next prompt (use when you know it's a single tool call). |

## Tools

| Command | Behavior |
|---|---|
| `/tools` | List every tool the AI can call, grouped by area. |
| `/tags` | Filter `/tools` by destructive / non-destructive / meta. |

## Sessions

| Command | Behavior |
|---|---|
| `/list` | List saved chat sessions for the current tenant. |
| `/load <id-or-prefix>` | Resume a saved session. Loads history + privacy map. |
| `/save [title]` | Persist the current chat. Title defaults to the first user message. |
| `/rename <id-or-prefix> <new-title>` | Rename a saved session. |
| `/delete <id-or-prefix>` | Delete a saved session + its DPAPI blob. |
| `/ephemeral` | Mark the current chat no-save (skipped at `/quit`). |
| `/export <id> [path]` | Write a redacted JSON safe to share (placeholders, no real values). |

## Cost

| Command | Behavior |
|---|---|
| `/cost` | Current session cost summary (calls, tokens, USD by provider). |
| `/costs` | Historical rollup (last-7d + last-30d + per-tenant + budget status). |

## Tenants

| Command | Behavior |
|---|---|
| `/tenants` | List registered tenants. |
| `/tenant <name>` | Switch to a registered tenant. Re-issues OAuth or app-only re-auth. |

## Incident response

| Command | Behavior |
|---|---|
| `/incident <upn> [Low\|Medium\|High\|Critical]` | Trigger the compromised-account playbook via the AI planner. Severity defaults to High. Forces plan mode for the next prompt. |

## Quit

| Command | Behavior |
|---|---|
| `/quit` (alias `/exit`, `/back`) | Exit the assistant. Auto-saves the current chat unless `/ephemeral` was set; writes a `/quit` audit line. |

## Command syntax

- Commands start with `/` and are case-insensitive.
- Multi-argument commands use space separation: `/incident alice@contoso.com Critical`.
- Arguments are NOT quoted in the shell sense — the parser splits on whitespace. Filenames with spaces should use the menu-driven equivalent.
- `id-or-prefix` accepts a full session id (`20260514-173208-a1b2c3d4`) or a unique title prefix.

## Where chat history lives

`<stateDir>\chat-sessions\<id>.session` (DPAPI-encrypted on Windows; plaintext + warning on POSIX). One file per saved session. The unencrypted `<stateDir>\chat-sessions\index.json` is the lookup map for `/list`.

See [`../guides/ai-sessions.md`](../guides/ai-sessions.md) for the session-management walkthrough and [`../concepts/ai-assistant.md`](../concepts/ai-assistant.md) for the conceptual model.

## See also

- [`tool-catalog.md`](tool-catalog.md) — every AI tool the model can call.
- [`menu-map.md`](menu-map.md) — non-AI menu navigation.
- [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md) — using the AI day-to-day.
