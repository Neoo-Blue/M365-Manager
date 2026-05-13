# AI cost tracking (Phase 5)

Every chat turn lands one event in a JSONL log and updates the
running session + monthly + per-tenant USD totals. Operators see a
one-line footer after each turn and can pull historical totals via
`/cost` and `/costs`.

## Where data lives

```
<stateDir>/ai-cost/
  events-YYYY-MM.jsonl   -- one line per chat call (provider, model,
                            tokens, USD, tenant, reason)
  monthly.json           -- per-month + per-tenant rollup + alert
                            crossings (so the same threshold doesn't
                            fire twice in a billing month)
```

`<stateDir>` resolves to `%LOCALAPPDATA%\M365Manager` on Windows or
`~/.m365manager` on POSIX (chmod 700).

## Price table

`templates/ai-prices.json` defines USD per million tokens, separately
for `input` and `output`:

```json
{
  "Anthropic": {
    "claude-opus-4-7":   { "input": 15.00, "output": 75.00 },
    "claude-sonnet-4-6": { "input":  3.00, "output": 15.00 },
    "claude-haiku-4-5":  { "input":  1.00, "output":  5.00 },
    "claude-opus-*":     { "input": 15.00, "output": 75.00 }
  },
  ...
}
```

Lookup order:

1. **Exact match** on provider + model name.
2. **Family match** on the longest `key*` prefix.
3. **`unknown`** fallback -- recorded at $0 with a marker so /cost
   surfaces "[price unknown]".

Edit the file to match your account's negotiated rates. Ollama is
pinned to $0 by default since local inference has no per-token cost.

## Footer format

```
  cost: in=843 out=1102 $0.0247 | session $0.4133 | month $12.8911
```

- `in` / `out` -- tokens for THIS call.
- `$x.xxxx` -- USD for THIS call.
- `session` -- running session total.
- `month` -- current month total.

Calls that recorded $0 (Ollama, unknown models) print nothing so
local-only chats stay quiet.

## Budget alerts

Add two optional keys to `Config`:

```json
{ "MonthlyBudgetUsd": 50.00, "AlertAtPct": 80 }
```

The tracker fires once per crossing per month at the 50% / `AlertAtPct`
(default 80%) / 100% / 150% marks. Each crossing:

- Writes an `AIBudgetAlert` audit entry with the USD used and the
  threshold.
- Surfaces a `[!]` warning under the cost footer in the chat.
- Stamps the threshold in `monthly.json.alerted` so it doesn't re-fire.

The session shop floor still works -- this is a softer cousin of a
hard budget cap. If you want a hard cap, wire `Test-AIBudgetCap`
into your pre-call hook and refuse to call `Invoke-AIChat` past the
threshold.

## Commands

| Command  | Output                                                                |
|----------|-----------------------------------------------------------------------|
| `/cost`  | Current session totals (calls, tokens, USD) + month-to-date rollup + per-tenant breakdown. |
| `/costs` | Last 6 monthly totals + 7-day rolling total scanned from the current month's event log. |

## Caveats

- **Tool-call cost**: each provider hop is one chat call, so a
  tool-using turn lands multiple events (one per hop). The footer
  shows the running session total so this is visible.
- **Cached tokens**: Anthropic exposes `cache_creation_input_tokens`
  and `cache_read_input_tokens`; we count cache-read as input at full
  price for now (slight over-estimate). The audit log keeps the raw
  usage object so post-hoc analysis is possible.
- **Privacy redaction tokens**: the redaction layer can grow the
  prompt slightly (placeholder tokens are short, but counts can climb
  for long sessions). This is reflected in your real bill.
