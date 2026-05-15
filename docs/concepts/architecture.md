# Architecture

How M365 Manager is organized — module map, dot-source order, and the threading model for state-mutating operations.

## Single-process, dot-sourced PowerShell

The tool is not a module — it's a collection of `.ps1` files that `Main.ps1` dot-sources into a single PowerShell process. There's no `psd1`, no `Import-Module`, no isolated session-state per "module". This shape is deliberate:

- **Tenant connections are process-global.** `Connect-MgGraph` mutates SDK state visible to every script that calls Graph cmdlets. Trying to isolate per-module session state would fight the SDK rather than work with it.
- **Audit / undo / preview are cross-cutting.** Every feature module's mutations route through `Invoke-Action` (in `Preview.ps1`), which writes the audit log entry, gates on PREVIEW mode, and captures the reverse recipe. A module isolation boundary would make that wiring much harder.

The tradeoff: function-name collisions are possible (mitigated by review + Pester) and there's no per-module versioning.

## Module map

By rough layer:

```
+---------------------------------------------------------------+
|  Entry point                                                  |
|    Main.ps1            menu + main loop                       |
+---------------------------------------------------------------+
|  Foundation -- loaded first, used by everything below         |
|    UI.ps1              colors, banners, menus, prompts        |
|    Auth.ps1            connect/disconnect, tenant picker      |
|    Audit.ps1           Write-AuditEntry + log path resolver   |
|    Preview.ps1         Invoke-Action wrapper + PREVIEW mode   |
|    Templates.ps1       role-template loader                   |
|    Notifications.ps1   Send-Email + Send-TeamsCard            |
|    TenantRegistry.ps1  Phase 6: tenants.json                  |
|    TenantSwitch.ps1    Phase 6: Switch-Tenant                 |
|    TenantOverrides.ps1 Phase 6: Get-EffectiveConfig           |
+---------------------------------------------------------------+
|  Feature modules -- each owns one menu area                   |
|    Onboard, BulkOnboard, Offboard, BulkOffboard               |
|    License, Archive                                            |
|    SecurityGroup, DistributionList, SharedMailbox             |
|    CalendarAccess, UserProfile, GroupManager                  |
|    Reports, eDiscovery                                         |
|    AuditViewer, Undo                                           |
|    SignInLookup, UnifiedAuditLog, MFAManager                  |
|    OneDriveManager, TeamsManager, SharePoint                  |
|    GuestUsers, LicenseOptimizer                                |
|    Scheduler, BreakGlass                                       |
|    MSPReports, MSPDashboard (Phase 6)                          |
|    IncidentResponse, IncidentRegistry, IncidentBulk,          |
|    IncidentTriggers (Phase 7)                                 |
+---------------------------------------------------------------+
|  AI assistant -- Phase 5, loaded last                         |
|    AICostTracker, AISessionStore, AIUx                        |
|    AIToolDispatch (catalog + provider payloads)               |
|    AIPlanner (submit_plan flow)                                |
|    AIAssistant (REPL + chat commands)                          |
+---------------------------------------------------------------+
```

The load order is the order declared in `Main.ps1:71-83`'s `$modules` array. Functions are visible to everything that loads after them; cross-module references in earlier-loaded modules are guarded with `Get-Command -ErrorAction SilentlyContinue` so the dependency graph stays acyclic at load time.

## Threading mutations through Invoke-Action

Every state change in every feature module routes through `Invoke-Action` (`Preview.ps1`). The contract:

```powershell
Invoke-Action `
    -Description "Block sign-in for $upn" `        # human-readable
    -ActionType  "BlockSignIn" `                    # stable filter / undo key
    -Target      @{ userUpn = $upn; userId = $id } # operands (structured)
    -ReverseType 'UnblockSignIn' `                  # for the undo dispatch table
    -ReverseDescription "Re-enable sign-in for $upn" `
    -Action      {                                  # the actual mutation
        Update-MgUser -UserId $id -AccountEnabled $false -ErrorAction Stop
    }
```

`Invoke-Action`:

1. Generates an `entryId` (UUID) for correlation.
2. Writes one **PROPOSE** audit line (PREVIEW mode) or **EXEC** audit line (LIVE).
3. Runs the `-Action` scriptblock only when LIVE.
4. Writes a **OK** / **ERROR** line on success / failure (LIVE only — PREVIEW has nothing to follow up).
5. Stamps the `reverse` recipe into the audit line for later `Invoke-Undo` dispatch.

A non-reversible action sets `-NoUndoReason` instead of `-ReverseType` — the audit line is then explicit about why it can't be undone.

PREVIEW mode runs every step's logging path but skips the actual cmdlet. The result is a tenant-untouched audit log that the operator can review before committing.

## Per-tenant audit log

`Audit.ps1:23` resolves the session log path to:

```
%LOCALAPPDATA%\M365Manager\audit\session-<timestamp>-<pid>-<tenant-slug>.log
```

The tenant slug comes from the active `$script:SessionState.TenantName`. `Switch-Tenant` calls `Reset-AuditLogPath` so the post-switch entries land in a new file. This makes cross-tenant operations easy to bucket — one file per tenant per session.

## Undo system

`Undo.ps1` exposes `Invoke-Undo -EntryId <id>`. It:

1. Loads the audit log entry by id.
2. Reads its `reverse` recipe (`type`, `description`, `target`).
3. Dispatches to `$script:UndoHandlers[reverse.type]` — a hashtable of scriptblocks, one per known reverse type.
4. Runs the handler against the target.
5. Writes its own audit entry for the reversal + stamps the original entry as reversed in `audit/undo-state.json` so it can't be undone twice.

Adding a new reversible action means: (a) emit `-ReverseType X` from the forward action, (b) register `$script:UndoHandlers['X']` in `Undo.ps1`. See [`../developer/adding-a-module.md`](../developer/adding-a-module.md).

## Connection lifecycle

`Connect-ForTask <area>` (in `Auth.ps1`) is the entry-point most feature functions call. It looks up the area's connection requirements (Graph + EXO? Graph alone? +SCC?) and ensures every service is connected before returning `$true`. The result is one-time setup cost; subsequent calls are no-ops.

`Disconnect-AllSessions` runs at `Quit and Disconnect` (or when `Switch-Tenant` fires). It tears down Graph + EXO + SCC + SPO so the next tenant starts clean.

## AI assistant integration

The AI layer (Phase 5) is structurally a separate concern but uses the same audit / preview / undo / redaction primitives. When the operator types `/incident alice@contoso.com Critical`, the synthesized prompt goes to the model with the full tool catalog attached; the model responds with `tool_use` blocks for specific catalog entries; the dispatcher resolves each to a real PowerShell function and invokes it through `Invoke-Action` (or through the function's own internal wrap). Audit lines from AI-driven calls carry an `actionType` like `AI:Update-MgUser` so the audit-viewer can distinguish them.

See [`ai-assistant.md`](ai-assistant.md) for the chat REPL + plan-approval model.

## Health checks

`health-checks/*.ps1` are short standalone scripts that bootstrap their own minimal dependency set via `_bootstrap.ps1`, run one check (license usage, MFA gaps, stale guests, incident triggers, etc.), and write a result JSON via `_writeresult.ps1`. The Scheduler module (`Scheduler.ps1`) registers these as Windows Task Scheduler entries running on cron-style intervals; results land in `<stateDir>/health-results/` for the MSP dashboard to roll up.

## State directory

`<stateDir>` resolves to `%LOCALAPPDATA%\M365Manager\state` on Windows, `~/.m365manager/state` (chmod 700) on POSIX. Contents:

```
<stateDir>/
├── tenants.json                       # tenant registry (Phase 6)
├── secrets/
│   └── tenant-<name>.dat              # DPAPI-encrypted credential manifests
├── tenant-overrides/<tenant>.json     # per-tenant config overrides
├── chat-sessions/<id>.session         # DPAPI-encrypted saved chats (Phase 5)
├── ai-costs.jsonl                     # AI usage cost log (Phase 5)
├── health-results/                    # scheduled health check outputs
├── breakglass-accounts.json           # break-glass registry (Phase 4)
├── scheduler-cred.xml                 # DPAPI scheduler service principal
├── undo-state.json                    # audit entries already reversed
└── <tenant-slug>/
    ├── incidents.jsonl                # incident registry (Phase 7)
    └── incidents/<INC-...>/           # per-incident snapshot + artifacts
```

See [`security-model.md`](security-model.md) for which paths are encrypted vs plaintext + the threat model around each.

## See also

- [`security-model.md`](security-model.md) — DPAPI, redaction, AST allow-list, audit log, undo internals.
- [`multi-tenant.md`](multi-tenant.md) — Phase 6 tenant model.
- [`ai-assistant.md`](ai-assistant.md) — Phase 5 AI model.
- [`../developer/architecture.md`](../developer/architecture.md) — deeper dive for contributors.
- [`../reference/cmdlets.md`](../reference/cmdlets.md) — public function reference grouped by module.
