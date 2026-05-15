# AI tool catalog

Every tool the AI assistant can invoke. Sourced from `ai-tools/*.json`.

## Catalog entry shape

```jsonc
{
  "name":                "Remove-Guest",
  "description":         "Full guest teardown: revoke shares, remove from groups + teams, delete user.",
  "parameters":          {  /* JSON Schema for the input */ },
  "destructive":         true,
  "wrapInInvokeAction":  true,
  "reverseTool":         null,
  "isMeta":              false,
  "requiresExplicitApproval": false
}
```

Field reference:

| Field | Notes |
|---|---|
| `name` | Stable name. Must match the PowerShell function the dispatcher will call. |
| `description` | One-paragraph explanation. Sent to the model as context. |
| `parameters` | JSON Schema for the input shape. Validated by `Test-AIToolInput` before dispatch. |
| `destructive` | `true` adds the red `[DESTRUCTIVE]` marker in the operator's confirmation prompt. |
| `wrapInInvokeAction` | When `true` AND the function isn't already wrapping internally, the dispatcher's default branch wraps the call so it audits + respects PREVIEW. |
| `reverseTool` | Name of a paired catalog tool that undoes this one. Informational; the planner uses it to schedule cleanups. |
| `isMeta` | `true` for the two meta-tools (`submit_plan`, `ask_operator`); these don't dispatch to real cmdlets. |
| `requiresExplicitApproval` | Phase 7. Forces a yellow banner, drops `[A]pprove-all` from the prompt, downgrades plan-mode `approveAll` to `stepByStep`. |

## Catalog files

```
ai-tools/
├── _meta.json          # submit_plan, ask_operator
├── users.json          # Get-MgUser, Update-MgUser, Set-MgUser, etc.
├── groups.json         # group + member ops
├── licenses.json       # SKU assignment / removal
├── mailboxes.json      # mailbox config (forwarding, OOO, etc.)
├── audit.json          # read-only audit + sign-in lookups
├── mfa.json            # method list / revoke / TAP
├── lifecycle.json      # onboard + offboard helpers
├── sharepoint.json     # site + share ops
├── teams.json          # Teams membership + ownership
├── onedrive.json       # OneDrive handoff
├── guests.json         # guest discovery + removal
└── incident.json       # Phase 7 incident response (4 tools)
```

## Meta-tools (`_meta.json`)

| Name | Purpose |
|---|---|
| `submit_plan` | The AI calls this with a structured plan when the operator typed `/plan` or when 3+ tool calls are needed. Operator approves before any step runs. |
| `ask_operator` | The AI calls this when it needs structured input (a UPN, a date, an SKU). Renders as a console prompt; the operator's response goes back to the model as the next user message. |

## Read-only tools (no `destructive`)

Used by Mark to answer questions without changing state. Examples:

- `Get-User` — pull one user's core fields.
- `Search-Users` — search by display-name / department / job title.
- `Get-LicenseAssignments` — license rollup.
- `Get-UserGroups` — group memberships.
- `Search-SignIns` — read sign-in log.
- `Search-UAL` — read unified audit log.
- `Get-IncidentTimeline` — chronological audit slice for an incident id.
- `Get-IncidentList` — list incidents in the current tenant.

These default `wrapInInvokeAction=false` since there's no state mutation.

## Destructive tools

Marked `destructive=true`. The operator's confirmation prompt prints `[DESTRUCTIVE]` in red. A non-exhaustive list (full catalog in `ai-tools/*.json`):

- `Set-MgUserLicense-Add` / `Set-MgUserLicense-Remove` — license changes.
- `Update-MgUser` — generic user PATCH.
- `Set-Mailbox-Type` — convert mailbox (`Regular` ↔ `Shared`).
- `Set-Mailbox-Forwarding` — set or clear forwarding.
- `Add-MailboxPermission` / `Remove-MailboxPermission` — Full Access / Send-As grants.
- `Remove-MgGroupMember` / `Remove-DistributionGroupMember` — group / DL removal.
- `New-TemporaryAccessPass` — TAP issuance.
- `Revoke-MgUserSignInSession` — session kill.
- `Remove-AllAuthMethods` — Phase 4 MFA wipe.
- `Revoke-OneDriveAccess` — undo a OneDrive handoff grant.
- `Set-TeamOwnership` / `Remove-UserFromTeam` — Teams admin.
- `Remove-Guest` — full guest teardown.
- `Invoke-CompromisedAccountResponse` — the full Phase 7 playbook. `requiresExplicitApproval=true`.

## Reverse-tool pairing

Where applicable, destructive tools declare a `reverseTool`:

| Destructive | Reverse |
|---|---|
| `Set-MgUserLicense-Remove` | `Set-MgUserLicense-Add` |
| `Remove-MgGroupMember` | `Add-MgGroupMember` |
| `Remove-DistributionGroupMember` | `Add-DistributionGroupMember` |
| `Set-Mailbox-Forwarding` (clear) | `Set-Mailbox-Forwarding` (set) |

The planner uses this to optionally add a "if step N fails, run reverseTool against the prior result" recipe. This is opportunistic — not every reverse is meaningful (some are no-ops, some need additional context).

## How a tool runs

1. Operator types a request.
2. Model produces one or more `tool_use` blocks (Anthropic) or `tool_calls` (OpenAI).
3. `AIToolDispatch.Invoke-AIToolCall`:
   - Looks up the catalog entry by name.
   - Validates the input via `Test-AIToolInput` (JSON Schema check).
   - Calls `Test-AICommandAllowed` to verify the command is on the AST allow-list (defense in depth).
4. Operator sees `Mark wants to call: [DESTRUCTIVE] <name>` + the parameters.
5. Operator approves (`Y` / `A` / `N` / `Q`).
6. Dispatcher runs the call:
   - Explicit case branch for tools that need parameter remapping (license adds/removes, mailbox setters).
   - Default branch (`& $ToolName @splat`) for the rest, optionally wrapped in `Invoke-Action`.
7. Result returns to the model as a `tool_result` block.
8. Audit entries lands with `actionType=AI:<name>` (default branch) or the function's own actionType (explicit branch).

## Adding a tool

See [`../developer/adding-an-ai-tool.md`](../developer/adding-an-ai-tool.md) for the full walkthrough. TL;DR:

1. Implement the PowerShell function (or pick an existing cmdlet).
2. Add a JSON entry to the appropriate `ai-tools/<area>.json`.
3. If the function needs parameter remapping or special handling, add an explicit case in `AIToolDispatch.Invoke-AIToolImpl`.
4. Run `Invoke-Pester ./tests/` — catalog tests verify the entry parses + the JSON Schema is valid.

## See also

- [`../concepts/ai-assistant.md`](../concepts/ai-assistant.md) — conceptual model.
- [`../guides/ai-tools-overview.md`](../guides/ai-tools-overview.md) — operator-facing.
- [`../developer/adding-an-ai-tool.md`](../developer/adding-an-ai-tool.md) — contribution guide.
