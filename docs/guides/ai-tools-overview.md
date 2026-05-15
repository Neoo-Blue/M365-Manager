# AI tool calling (Phase 5)

Phase 5 replaced the legacy `RUN: <cmd>` regex extractor with native
provider tool calling. The model proposes a tool by name, the operator
approves (Y / A / N / Q), and the dispatcher runs the corresponding
PowerShell function. The regex path is still loaded as a fallback for
Ollama models that lack tool support, but it now emits a deprecation
warning on first use.

## Where tools live

```
ai-tools/
  _meta.json        -- submit_plan, ask_operator (special / always loaded)
  users.json        -- Get-User, Set-Mailbox-*, etc.
  groups.json
  licenses.json
  mailboxes.json
  audit.json
  mfa.json
  lifecycle.json
  sharepoint.json
  teams.json
  onedrive.json
  guests.json
```

Each entry follows JSON Schema:

```json
{
  "name": "Remove-Guest",
  "description": "Full guest teardown: revoke shares, remove from groups + teams, then DELETE /users/{id}. 30-day restore window via /directory/deletedItems.",
  "parameters": {
    "type": "object",
    "properties": {
      "UPN":    { "type": "string" },
      "Reason": { "type": "string" }
    },
    "required": ["UPN","Reason"]
  },
  "destructive": true,
  "wrapInInvokeAction": true,
  "reverseTool": null
}
```

Fields:

| Field                 | Purpose                                                                    |
|-----------------------|----------------------------------------------------------------------------|
| `destructive`         | Drives the `[DESTRUCTIVE]` marker in the prompt and plan-mode logic.       |
| `wrapInInvokeAction`  | When true, the dispatcher routes through `Invoke-Action` so audit/preview/undo work. |
| `reverseTool`         | Name of the tool that undoes this one (informational; planner uses it too).|
| `isMeta`              | Set on `submit_plan` / `ask_operator`; rejected as plan steps.             |

## How a turn flows

1. `Get-AIToolCatalog` loads + caches every JSON file in `ai-tools/`.
2. `Invoke-AIChatToolingTurn` builds a provider-shaped payload
   (`tool_use` blocks for Anthropic, `tool_calls` for OpenAI / Azure,
   Ollama function-calling).
3. Provider response is normalized to
   `@{ Text; ToolUses; AssistantContent; StopReason; Usage; Error }`.
4. Each `tool_use` becomes a Y/A/N/Q prompt
   (`Mark wants to call: [DESTRUCTIVE] <name>` + arg list).
5. Approved calls run via `Invoke-AIToolCall`. Unknown / wrong-shape
   inputs are caught by `Test-AIToolInput` before any code executes.
6. The result is JSON-encoded, truncated to 4 KB, and sent back as a
   `tool_result` block. The model keeps going until `StopReason =
   end_turn` or the 8-hop ceiling is reached.

## Privacy boundary

Outbound payloads (system prompt + history + tool definitions + tool
results) pass through `Convert-ToSafePayload` -- the same redaction
layer that protects regular chat. Inbound `tool_use` inputs and
assistant text pass through `Restore-FromSafePayload` before the
operator sees them or the dispatcher runs anything. The privacy map
is per-session and never leaves the local machine; `/clear` drops it.

## Adding a tool

1. Write the PowerShell function (must work standalone -- AI calls
   are not interactive).
2. Add a JSON entry in the appropriate `ai-tools/<area>.json` file.
3. If the function should respect PREVIEW / audit / undo, wrap it via
   `Invoke-Action` and set `wrapInInvokeAction: true`.
4. Register the dispatch in `Invoke-AIToolImpl` in `AIToolDispatch.ps1`
   (the switch-table that maps tool name -> function call).
5. Add a Pester test under `tests/AIToolDispatch.Tests.ps1`.

## Catalog browsing

```
You: /tools
```

Prints the full loaded catalog grouped by file, with the
`[DESTRUCTIVE]` marker on each. Useful for sanity-checking before a
risky chat.

## Deprecation timeline for the regex path

The `RUN: <PowerShell>` extractor stays available so old Ollama models
keep working, but it now logs `[deprecated] Using regex RUN: extractor
because native tool support was not detected` on first use. The plan
is to remove it once `gpt-oss` / `llama-3.1` / `qwen2.5` models are
ubiquitous on Ollama.
