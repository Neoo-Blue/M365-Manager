# AI multi-step plans (Phase 5)

For tasks that need three or more tool calls, or any sequence of
destructive operations, Mark builds a structured plan FIRST and the
operator approves the whole thing before any step runs. This stops
the "approve 12 dialogs in a row" pattern from earlier phases and
gives operators a clean checkpoint for risky bulk work.

## Triggers

Plan mode kicks in three ways:

1. `/plan` -- forces plan mode for the next prompt. Mark *must* call
   `submit_plan` first; calling tools directly returns an error.
2. **Auto-trigger** -- if Mark proposes >= `AI.AutoPlanThreshold`
   tool calls on the first hop and didn't already submit a plan,
   the chat layer rejects all of them with `auto_plan_required`
   and the AI is re-prompted in forced plan mode. Default
   threshold is 3.
3. **System-prompt guidance** -- the planning addendum tells Mark
   to use `submit_plan` for any sequence of destructive operations
   even when the count is < 3.

`/noplan` disables planning for one prompt. Useful when you know the
work is a single tool call but the model is being over-cautious.

## Plan schema

The model calls the special meta-tool `submit_plan`:

```json
{
  "goal": "Offboard alice@contoso.com",
  "steps": [
    { "id": 1, "description": "Reset password",                "tool": "Reset-UserPassword",  "params": { "UPN": "alice@contoso.com" }, "destructive": true,  "dependsOn": []  },
    { "id": 2, "description": "Revoke sign-in sessions",       "tool": "Revoke-UserSignIns",  "params": { "UPN": "alice@contoso.com" }, "destructive": true,  "dependsOn": [1] },
    { "id": 3, "description": "Convert mailbox to shared",     "tool": "Convert-MailboxShared", "params": { "UPN": "alice@contoso.com" }, "destructive": true,  "dependsOn": [2] },
    { "id": 4, "description": "Remove all licenses",           "tool": "Remove-UserLicenses", "params": { "UPN": "alice@contoso.com" }, "destructive": true,  "dependsOn": [3] }
  ],
  "estimatedDurationSec": 45,
  "destructiveStepCount": 4,
  "failureMode": "stop"
}
```

| Field                  | Required | Notes                                                        |
|------------------------|----------|--------------------------------------------------------------|
| `goal`                 | yes      | One-sentence summary. Shown at the top of the approval UI.  |
| `steps[].id`           | yes      | Unique integer. Used to express dependencies.                |
| `steps[].description`  | yes      | Human-readable line shown in the approval UI.                |
| `steps[].tool`         | yes      | Must match a real entry in `ai-tools/*.json`. Meta-tools rejected. |
| `steps[].params`       | yes      | Passed verbatim to the dispatcher; validated against schema. |
| `steps[].dependsOn`    | no       | Array of earlier `id`s. Forward refs / cycles rejected.      |
| `steps[].destructive`  | no       | Marks the step `[DESTRUCTIVE]` in the approval UI.           |
| `failureMode`          | no       | `stop` (default) aborts on first failure; `ask` prompts.     |

## Approval UX

```
  AI PLAN: Offboard alice@contoso.com
    Estimated duration : 45 seconds
    Steps total        : 4
    Destructive        : 4

    [ 1] [DESTRUCTIVE] Reset-UserPassword
          Reset password
            UPN: alice@contoso.com
    [ 2] [DESTRUCTIVE] Revoke-UserSignIns (depends on: 1)
          Revoke sign-in sessions
            UPN: alice@contoso.com
    ...

  Plan action: [A]pprove all / [S]tep-by-step / [E]dit / [R]eject:
```

- `[A]` -- runs every step in dependency order, no per-step prompt.
- `[S]` -- prompts `[Y/A/S/Q]` per step.
- `[E]` -- drops the plan JSON into `$EDITOR` (or notepad / nano).
  Re-validated on save.
- `[R]` -- discards. AI gets a `rejected` tool_result and may
  re-plan or pivot.

## Execution + audit

`Invoke-AIPlan` walks `Get-TopologicalStepOrder` and:

- Writes one `AIPlan` audit entry at start (with step count + goal).
- Writes one per-step entry via `Invoke-AIToolCall` (which itself
  emits the standard `AIToolCall` shape).
- Writes one `AIPlanResult` summary entry at the end.
- Honors PREVIEW: tools with `wrapInInvokeAction: true` log
  `[would run]` and skip the actual cmdlet.

## Failure recovery

`failureMode: "stop"` (default) aborts at the first failed step.
`failureMode: "ask"` prompts `[C]ontinue / [A]bort / [R]evise` --
choose Revise and the AI gets a `tool_result` containing the
partial trace and is asked to submit a new plan that picks up
where the previous one left off. The new plan goes through the
same approval gate.

## Non-interactive automation

Set `$env:M365MGR_PLAN_APPROVAL` to one of `approveAll`, `stepByStep`,
`edit`, `reject` before running -- in NonInteractive mode, the planner
reads that instead of prompting. Default if unset is `reject` so a
scheduled run can never silently destroy data.
