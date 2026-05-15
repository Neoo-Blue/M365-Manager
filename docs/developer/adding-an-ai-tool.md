# Adding an AI tool

How to make a PowerShell function callable by the AI assistant. The catalog (`ai-tools/*.json`) is the AI's interface to the tool; this guide walks through extending it.

## What you're adding

Two artifacts:

1. **Catalog JSON entry** in `ai-tools/<area>.json` — declarative.
2. **(Optional) Dispatcher case** in `AIToolDispatch.Invoke-AIToolImpl` — only needed if the function needs parameter remapping.

The function itself either exists already (you're surfacing a public M365 Manager cmdlet to the AI) or you're adding it as part of [`adding-a-module.md`](adding-a-module.md).

## Step 1 — pick the catalog file

Tools are grouped by feature area in `ai-tools/`:

```
ai-tools/
├── _meta.json          # submit_plan, ask_operator (don't touch)
├── users.json          # user CRUD + properties
├── groups.json         # group / DL / role-group membership
├── licenses.json       # SKU operations
├── mailboxes.json      # mailbox config
├── audit.json          # read-only audit + sign-in
├── mfa.json            # auth methods
├── lifecycle.json      # onboard / offboard
├── sharepoint.json
├── teams.json
├── onedrive.json
├── guests.json
└── incident.json       # Phase 7
```

If your tool fits an existing area, add to that file. If it's a brand new feature area, create `ai-tools/<area>.json`.

## Step 2 — write the catalog entry

```jsonc
{
  "name":        "Set-MailboxAutoReplyConfiguration",
  "description": "Configure Out-of-Office auto-reply for a mailbox. Supports scheduled start/end + separate internal/external messages.",
  "parameters": {
    "type": "object",
    "properties": {
      "Identity":          { "type": "string",  "description": "UPN or mailbox alias." },
      "AutoReplyState":    { "type": "string",  "enum": ["Disabled","Enabled","Scheduled"], "description": "Default Disabled." },
      "StartTime":         { "type": "string",  "description": "ISO 8601 UTC. Required when state is Scheduled." },
      "EndTime":           { "type": "string",  "description": "ISO 8601 UTC. Required when state is Scheduled." },
      "InternalMessage":   { "type": "string",  "description": "HTML allowed." },
      "ExternalMessage":   { "type": "string" },
      "ExternalAudience":  { "type": "string", "enum": ["None","Known","All"] }
    },
    "required": ["Identity", "AutoReplyState"]
  },
  "destructive":             true,
  "requiresExplicitApproval": false,
  "wrapInInvokeAction":       true,
  "reverseTool":              null
}
```

Field-by-field:

| Field | Purpose |
|---|---|
| `name` | Stable. Must match the PowerShell function the dispatcher routes to. |
| `description` | Sent to the AI as context. Be specific — "Configure Out-of-Office" is better than "Set OOO". |
| `parameters` | JSON Schema. `Test-AIToolInput` validates incoming AI calls against this before dispatch. |
| `destructive` | `true` adds `[DESTRUCTIVE]` red marker to the operator's confirmation. |
| `requiresExplicitApproval` | Phase 7 flag. Forces yellow banner + disables `[A]pprove-all`. Reserve for the highest-blast-radius tools. |
| `wrapInInvokeAction` | When `true`, the dispatcher's default branch wraps the call in `Invoke-Action` so it audits + respects PREVIEW. Set `false` ONLY if the function wraps internally (e.g. `Invoke-CompromisedAccountResponse` does its own per-step wrap). |
| `reverseTool` | Name of a paired catalog tool that undoes this one. Used by the planner for opportunistic reverse pairing. |

## Step 3 — write the JSON Schema for parameters

`Test-AIToolInput` enforces:

- **Required fields present.** If `required: ["Identity"]` and the AI omits `Identity`, the call is rejected before dispatch.
- **Type checks.** `string` / `integer` / `boolean` are enforced. Numeric strings auto-coerce to integers when the schema says `integer`.
- **Enum constraints.** If `enum: ["Disabled","Enabled","Scheduled"]`, anything else is rejected.

Pattern: be strict at the schema level so bad AI calls fail before any cmdlet runs.

For complex inputs, nested `object` schemas are supported:

```jsonc
"parameters": {
  "type": "object",
  "properties": {
    "Settings": {
      "type": "object",
      "properties": {
        "Quota":      { "type": "integer" },
        "Expiry":     { "type": "string" },
        "AutoArchive":{ "type": "boolean" }
      }
    }
  },
  "required": ["Settings"]
}
```

## Step 4 — dispatcher routing

`AIToolDispatch.Invoke-AIToolImpl` does the actual function call. It walks a switch table:

1. **Explicit case** for tools that need parameter remapping. Example:
   ```powershell
   '^Set-Mailbox-Forwarding$' {
       $isClear = -not $splat.ForwardTo
       return (Invoke-Action `
           -Description ("AI: Forwarding for {0}" -f $splat.Identity) `
           -ActionType ("Set-Mailbox-Forwarding") `
           -ReverseType $(if ($isClear) { 'SetForwarding' } else { 'ClearForwarding' }) `
           -Target @{ identity = $splat.Identity; forwardTo = $splat.ForwardTo } `
           -Action {
               if ($isClear) { Set-Mailbox -Identity $splat.Identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop }
               else          { Set-Mailbox -Identity $splat.Identity -ForwardingSmtpAddress ("smtp:" + $fwd) -DeliverToMailboxAndForward $deliver -ErrorAction Stop }
           })
   }
   ```

2. **Default case** for tools where `& $ToolName @splat` works. This is the common case. The default branch honors `wrapInInvokeAction` and wraps when set to `true`.

Choose explicit case when:

- The function takes parameters with different names than the AI catalog (e.g. catalog has `UserId`, function takes `-Identity`).
- The function takes positional args (AI tool catalog only does named).
- Multiple cmdlets need to fire as one logical step.
- You want a richer audit entry with explicit `actionType` + `reverseType`.

Choose default case when:

- The function accepts named parameters matching the catalog schema.
- Either the function wraps internally OR `wrapInInvokeAction=true` is sufficient.

## Step 5 — test the tool

`tests/AIToolDispatch.Tests.ps1` covers catalog parsing + schema validation. Add a test for your specific tool:

```powershell
Describe "Set-MailboxAutoReplyConfiguration catalog entry" {
    It "is registered" {
        $tool = Get-AIToolByName -Name 'Set-MailboxAutoReplyConfiguration'
        $tool | Should -Not -BeNullOrEmpty
        $tool.destructive | Should -BeTrue
    }
    It "rejects missing required Identity" {
        $tool = Get-AIToolByName -Name 'Set-MailboxAutoReplyConfiguration'
        $r = Test-AIToolInput -ToolDef $tool -InputHash @{ AutoReplyState = 'Enabled' }
        $r.Valid | Should -BeFalse
    }
    It "accepts valid input" {
        $tool = Get-AIToolByName -Name 'Set-MailboxAutoReplyConfiguration'
        $r = Test-AIToolInput -ToolDef $tool -InputHash @{ Identity = 'a@b.com'; AutoReplyState = 'Disabled' }
        $r.Valid | Should -BeTrue
    }
}
```

Plus dispatch test if you added an explicit case:

```powershell
Describe "Invoke-AIToolImpl: Set-Mailbox-Forwarding" {
    It "routes through Invoke-Action with the right ActionType" {
        Mock Set-Mailbox { return $true }
        Mock Write-AuditEntry { } -Verifiable
        # ... invoke + verify
    }
}
```

## Step 6 — document the tool

Add a row to [`../reference/tool-catalog.md`](../reference/tool-catalog.md) under the appropriate section.

If the tool surfaces a new workflow, also link from the relevant guide. E.g. if you add a guest-removal tool, the [`../guides/guest-lifecycle.md`](../guides/guest-lifecycle.md) should reference it.

## Patterns + anti-patterns

### Do
- **Make read-only tools generously available.** The AI is better at helping when it can answer "show me X" questions without per-call confirmation.
- **Be specific in `description`.** The AI uses this to decide whether to invoke. Vague descriptions lead to poor tool selection.
- **Pin destructive tools tightly.** `requiresExplicitApproval=true` is for the highest blast radius. Use sparingly.
- **Reverse-pair where possible.** Setting `reverseTool` helps the planner build clean cleanup paths.

### Don't
- **Don't take a free-text `Action` parameter.** "Action": "delete" / "update" / "create" is an anti-pattern — split into separate tools so the schema is meaningful.
- **Don't accept structured PowerShell objects.** JSON Schema can't validate `[PSCustomObject]` arbitrary content; pass primitives or nested JSON objects only.
- **Don't bypass `Invoke-Action`.** Even tools that "feel" safe (e.g. updating a display name) should audit.
- **Don't catalog internal helpers.** If a function is internal-only, leave it out of the catalog — every catalog entry is part of the public surface from the AI's perspective.

## Pre-merge fix history

PR #9 fixed a closure-scoping bug in the redaction layer that the dispatcher relies on. PR #6 fixed a catalog-honoring bug where the default branch was ignoring `wrapInInvokeAction`. Both happened because the AI catalog system has subtle interactions with the rest of the codebase — when you add new tools, run through these specifically:

1. Does the catalog entry validate against the schema in `Test-AIToolInput`?
2. Does the dispatched call get an audit entry visible in `Read-AuditEntries`?
3. Does the operator see the `[DESTRUCTIVE]` marker if marked?
4. In PREVIEW mode, does the call NOT actually execute?

If any of these fails, root-cause before merging.

## See also

- [`architecture.md`](architecture.md) — AI dispatch flow.
- [`../reference/tool-catalog.md`](../reference/tool-catalog.md) — full catalog reference.
- [`../concepts/ai-assistant.md`](../concepts/ai-assistant.md) — AI conceptual model.
- [`adding-a-module.md`](adding-a-module.md) — for new features that need a new module.
