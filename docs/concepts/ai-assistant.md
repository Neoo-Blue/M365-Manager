# The AI assistant ("Mark") — conceptual

How M365 Manager's AI integration is shaped. This is the concept doc; the task-oriented "how do I use feature X" walkthroughs live in [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md), [`../guides/ai-planning.md`](../guides/ai-planning.md), [`../guides/ai-sessions.md`](../guides/ai-sessions.md), and [`../guides/ai-costs.md`](../guides/ai-costs.md).

## What Mark is, and isn't

Mark is a thin REPL over a tool catalog. The operator types natural language; the model proposes calls to specific catalog'd PowerShell functions; the operator approves each call individually; the function runs through the same `Invoke-Action` path as menu-driven mutations.

Mark is NOT:

- An agent with autonomous execution. Every tool call needs operator approval.
- A general LLM proxy. Mark only invokes catalog'd tools — there's no shell-out, no arbitrary code execution.
- A long-running daemon. The REPL is single-process, single-session.

## The four pillars

### 1. Native tool calling

The Phase 5 design uses provider-native tool calling (Anthropic `tool_use` / OpenAI `tool_calls` / Ollama function-calling), not the legacy regex `RUN:` extractor. Tools are defined in `ai-tools/*.json` — one JSON file per area (`users.json`, `incident.json`, `licenses.json`, etc.). Each entry declares:

```jsonc
{
  "name":        "Remove-Guest",
  "description": "Full guest teardown: revoke shares, remove from groups, delete user.",
  "parameters":  { ... JSON Schema for the input shape ... },
  "destructive": true,
  "wrapInInvokeAction": true,
  "reverseTool": null,
  "isMeta":      false
}
```

The dispatcher (`AIToolDispatch.ps1`) validates the AI's proposed input against the schema, then routes to either an explicit case branch (for tools that need parameter remapping or special handling) or to a generic `& $ToolName @splat` default branch. The default branch honors `wrapInInvokeAction` so destructive cmdlets get a proper audit entry with `actionType=AI:<name>` and `noUndoReason`.

The regex `RUN:` path still exists for Ollama models that lack tool support, but it emits a one-time deprecation warning when it fires.

See [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md) for the catalog browsing surface.

### 2. Multi-step plans

For requests that need 3+ tool calls (configurable via `AI.AutoPlanThreshold`), Mark submits a plan first via the special meta-tool `submit_plan` (defined in `ai-tools/_meta.json`). The plan is a structured list of steps with dependencies, destructive flags, and per-step parameters; the operator approves the whole thing as `[A]pprove all` / `[S]tep-by-step` / `[E]dit` / `[R]eject` before any step runs.

`stepByStep` (and the no-approveAll-allowed flag on `Invoke-CompromisedAccountResponse`) keeps every individual call gated. The executor walks the dependency graph topologically; on failure, `failureMode='ask'` lets the operator continue / abort / ask Mark to revise.

See [`../guides/ai-planning.md`](../guides/ai-planning.md).

### 3. Persistent encrypted sessions

`/save` writes the current chat history (and privacy map) to a DPAPI-encrypted blob under `<stateDir>/chat-sessions/<id>.session`. `/list` shows saved sessions; `/load <id-or-prefix>` resumes one. Auto-save fires on `/quit` unless `/ephemeral` was toggled.

`/export <id> [path]` writes a *redacted* JSON safe to share — every UPN / GUID / secret in the history is replaced with its tokenized placeholder. The privacy map is intentionally excluded from the export.

See [`../guides/ai-sessions.md`](../guides/ai-sessions.md).

### 4. Cost tracking

Every model call lands a JSONL row in `<stateDir>/ai-costs.jsonl` with provider / model / input tokens / output tokens / dollar cost / tenant. The `/cost` and `/costs` commands roll up running session totals + last-7-day + last-30-day spends. Cost calculation uses `templates/ai-prices.json` — operators can edit this to match their negotiated rates.

`AI.MonthlyBudgetUsd` triggers a warning when crossed; the budget check itself is currently advisory (no hard cap on spend). See [`../guides/ai-costs.md`](../guides/ai-costs.md).

## Privacy guarantee

Every outbound payload passes through `Convert-ToSafePayload` before hitting the provider. Tokenization rules vary by setting:

- **Always tokenized** (regardless of provider): JWTs, `sk-...` API keys, cert thumbprints.
- **Tokenized for external providers** (default): UPNs, GUIDs, tenant IDs, captured display names.
- **Restored on response**: the operator sees real values; the dispatcher dispatches against real values.

The tokenization map is per-session, per-tenant, and dropped at `/clear`. If you `/save` a session and `/load` it later under a different tenant, the privacy map travels with the session blob — meaning a saved chat from tenant A is rehydrated with tenant-A real values even after switching to tenant B.

See [`security-model.md`](security-model.md) for the redaction implementation.

## The /incident special case

The Phase 7 incident-response playbook gets a dedicated chat command:

```
/incident alice@contoso.com Critical
```

Internally this synthesizes a forced-plan-mode prompt ("Run a Critical-severity compromised-account response against alice@contoso.com. Submit a plan first using submit_plan...") and lets the plan approval gate handle the rest. `Invoke-CompromisedAccountResponse` is tagged `requiresExplicitApproval` in the catalog, which:

- Forces a yellow `EXPLICIT APPROVAL REQUIRED` banner around its tool-use prompt.
- Drops `[A]pprove all` from the prompt options for that single call.
- Downgrades `approveAll` to `stepByStep` for any plan that includes the tool.

This is the only place in the tool where the AI's batch-approve path is forcibly disabled. The flag is reusable by future tools that warrant the same treatment — set `requiresExplicitApproval: true` in the catalog entry.

## When NOT to use Mark

- **Bulk operations.** The bulk CSV flows (`Invoke-BulkOnboard`, `Invoke-BulkOffboard`, `Invoke-BulkIncidentResponse`) are faster + more predictable than asking the AI to iterate. Use the CSVs.
- **Anything time-sensitive.** Cloud-LLM latency is 1-3 seconds per turn. For real-time troubleshooting, drive the menu.
- **Compliance-critical work.** The audit trail is identical (every AI-driven call goes through `Invoke-Action`), but the narrative path is harder to reproduce. For audit-prep work, prefer scripted runs that can be replayed.
- **When you're not paying attention.** Mark surfaces destructive proposals clearly, but it doesn't think for you. Don't `/incident` a UPN you haven't verified.

## Provider trust model

Default ships against Anthropic; OpenAI / Azure OpenAI / Ollama / Custom HTTP endpoints all work via the same `Invoke-AIChat` abstraction. Defaults:

| Provider | Default retention | Caveats |
|---|---|---|
| Anthropic | Zero retention (enterprise API key) | Verify against your contract; consumer keys have different terms. |
| OpenAI | 30-day abuse-monitoring retention | Zero Data Retention agreement needed for stricter handling. |
| Azure OpenAI | Governed by your subscription | In-tenant deployments stay inside your compliance boundary. |
| Ollama (localhost) | n/a — never leaves the machine | The default `TrustedProviders` list treats `localhost`-prefixed endpoints as non-PII-redacted (still scrubs secrets). |

`Privacy.TrustedProviders` is the operator's escape hatch for trusted external endpoints. Add `"azure-openai"` if your Azure deployment is in the same tenant you're administering — PII goes raw to that provider, secrets still tokenized.

## See also

- [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md) — using the AI, catalog browsing, chat commands.
- [`../guides/ai-planning.md`](../guides/ai-planning.md) — plan approval flow.
- [`../guides/ai-sessions.md`](../guides/ai-sessions.md) — persistent sessions, export.
- [`../guides/ai-costs.md`](../guides/ai-costs.md) — cost tracking + budgets.
- [`security-model.md`](security-model.md) — redaction internals.
- [`../reference/chat-commands.md`](../reference/chat-commands.md) — every slash command reference.
- [`../reference/tool-catalog.md`](../reference/tool-catalog.md) — every catalog'd tool.
