# Developer architecture deep-dive

The shape of the codebase from a contributor's perspective. Pairs with [`../concepts/architecture.md`](../concepts/architecture.md), which is the operator-facing concept view; this one goes deeper into how the pieces fit + the conventions a contributor should follow.

## Source layout

```
M365-Manager/
├── Launch.bat               # Windows entrypoint (sets env + execpol)
├── Main.ps1                 # Bootstrap + main menu loop
├── *.ps1                    # ~40 feature modules; dot-sourced flat
├── ai-tools/*.json          # AI tool catalog
├── templates/               # role / site / tabletop / price templates
├── health-checks/           # scheduled-task scripts
├── tests/                   # Pester suites
├── docs/                    # documentation
└── ai_config.example.json   # config template
```

Every `.ps1` lives at the repo root. They're dot-sourced in the order declared in `Main.ps1:71-83`. There are no nested subdirectories for feature modules — by design, since:

- Dot-source semantics work best with a flat layout.
- The collision risk is mitigated by per-feature function-name prefixes.
- Discoverability favors flat over deep.

## Dot-source order

```powershell
$modules = @(
    # Foundation (everything below depends on these)
    "UI.ps1","Auth.ps1","Audit.ps1","Preview.ps1","Templates.ps1",
    "Notifications.ps1","TenantRegistry.ps1","TenantSwitch.ps1","TenantOverrides.ps1",

    # Feature modules
    "Onboard.ps1","BulkOnboard.ps1","Offboard.ps1","BulkOffboard.ps1",
    "License.ps1","Archive.ps1","SecurityGroup.ps1","DistributionList.ps1",
    "SharedMailbox.ps1","CalendarAccess.ps1","UserProfile.ps1",
    "Reports.ps1","eDiscovery.ps1","GroupManager.ps1",
    "AuditViewer.ps1","Undo.ps1","SignInLookup.ps1","UnifiedAuditLog.ps1",
    "MFAManager.ps1","OneDriveManager.ps1","TeamsManager.ps1","SharePoint.ps1",
    "GuestUsers.ps1","LicenseOptimizer.ps1","Scheduler.ps1","BreakGlass.ps1",
    "MSPReports.ps1","MSPDashboard.ps1",
    "IncidentResponse.ps1","IncidentRegistry.ps1","IncidentBulk.ps1","IncidentTriggers.ps1",

    # AI layer (loaded last so it can reference everything above)
    "AICostTracker.ps1","AISessionStore.ps1","AIUx.ps1","AIToolDispatch.ps1","AIPlanner.ps1","AIAssistant.ps1"
)
```

A module loaded earlier cannot reference functions defined later. Mitigation: use `Get-Command <Name> -ErrorAction SilentlyContinue` guards when a forward reference is unavoidable.

Adding a new module: append to this array at the appropriate position. Almost always between an existing module and the AI block.

## Function naming conventions

| Pattern | Used for |
|---|---|
| `Verb-Noun` | Public functions. Use approved verbs (`Get`, `Set`, `New`, `Remove`, `Invoke`, `Start`, `Test`, etc.). |
| `Get-X` | Read-only; returns data. |
| `Set-X` / `Remove-X` / `New-X` | Mutations. Route through `Invoke-Action`. |
| `Invoke-X` | Orchestrator; may chain multiple atomic operations. |
| `Start-XMenu` | Submenu entry points called from `Main.ps1`. |
| `Test-X` | Returns `$true`/`$false`. |
| `Read-X` / `Write-X` | Lower-level file I/O helpers. Generally module-internal. |

Internal-only helpers don't follow Verb-Noun rigidly — they often use camelCase function names. The line between "public" (documented in [`../reference/cmdlets.md`](../reference/cmdlets.md)) and "internal" is the cmdlets.md inclusion list.

## The Invoke-Action contract

Every state-mutating operation routes through `Invoke-Action` (`Preview.ps1`). The contract is:

```powershell
Invoke-Action `
    -Description "human-readable" `
    -ActionType  "StableShortTag" `       # for audit filter + undo dispatch
    -Target      @{ structured = "operands" } `
    -ReverseType 'InverseActionType' `    # optional; pairs with $UndoHandlers[X]
    -ReverseDescription "Undo it" `
    -ReverseTarget @{ ... } `            # optional; defaults to -Target
    -NoUndoReason "explanation" `         # optional; mutually exclusive with -Reverse*
    -Critical `                           # optional; rethrows on failure
    -StubReturn $stubObj `                # PREVIEW stand-in
    -Action {                             # the actual mutation
        Set-MgUser -UserId $id -AccountEnabled $false -ErrorAction Stop
    }
```

The wrapper:

1. Generates an `entryId` (UUID).
2. PREVIEW mode: writes a `PROPOSE` audit line; returns `$StubReturn`.
3. LIVE mode: writes `EXEC`; runs the scriptblock; writes `OK` or `ERROR`.
4. Stamps `reverse` recipe if `-ReverseType` is set, OR `noUndoReason` if explicit.

**Every contributor mutation must use this wrapper.** Direct cmdlet calls bypass the audit log and break undo. The pre-merge review (PR #6) caught and fixed a place where the AI dispatcher's default branch was bypassing the wrapper for SDK cmdlets.

## The audit log

JSONL at `%LOCALAPPDATA%\M365Manager\audit\session-<ts>-<pid>-<tenant>.log`. One line per `Invoke-Action` event. Full schema at [`../reference/audit-format.md`](../reference/audit-format.md).

Conventions:
- `actionType` is a stable short string. Don't change them — downstream filtering depends on them.
- `target` is a hashtable. Use snake_case keys consistently (`userUpn`, `groupId`, `skuId`).
- `tenant` is a structured `{name, id, domain, mode}` block from Phase 6 onward.
- `reverse.target` defaults to `target` but can be a subset (e.g. omit redundant fields).

## The undo dispatch table

`$script:UndoHandlers` (in `Undo.ps1`) is a hashtable. Keys are reverse types; values are scriptblocks that take a `[hashtable]$Target` and run the inverse operation.

Adding a new reverse: define the handler in the table, emit `-ReverseType X` from the forward action. See [`../reference/undo-handlers.md`](../reference/undo-handlers.md) and [`adding-a-module.md`](adding-a-module.md).

## Connection lifecycle

`Connect-ForTask <area>` (in `Auth.ps1`) is the entry point for getting connected to whatever services a flow needs. The areas:

| Area | Connects |
|---|---|
| `Onboard` / `Offboard` | Graph + EXO |
| `License` | Graph |
| `Archive` | EXO |
| `Reports` | Graph |
| `eDiscovery` | EXO + SCC |
| `SharePoint` / `Guests` | Graph + SPO |
| `Audit` | Graph + EXO (for UAL) |
| `Incident` | Graph + EXO + SPO + SCC |

Adding a new area: extend the lookup table in `Connect-ForTask`.

`Reset-AllSessions` tears down every connection. Called on `Switch-Tenant` and `Quit`.

## State directory

`<stateDir>` = `%LOCALAPPDATA%\M365Manager\state` on Windows, `~/.m365manager/state` on POSIX. Created with `chmod 700` on POSIX.

Conventions for a new module writing state:
- Use `Get-StateDirectory` to resolve the base path.
- Tenant-scoped artifacts under `<stateDir>/<tenant-slug>/<area>/`.
- Encrypt sensitive content with `Protect-Secret` (DPAPI on Windows, plaintext + warning on POSIX).
- Use JSONL for append-only logs, JSON for state files.

## AI integration

If your new module's public functions should be AI-callable, add a catalog entry (see [`adding-an-ai-tool.md`](adding-an-ai-tool.md)). The contract:

- The function must work with a hashtable splat (`@splat`).
- All parameters should be named (avoid positional).
- Mark `destructive` correctly. Mark `requiresExplicitApproval` if the operation has very high blast radius.
- Set `wrapInInvokeAction` based on whether the function internally wraps OR you want the dispatcher to wrap.

## Testing

`tests/<Module>.Tests.ps1` per module. Pester 5. See [`testing.md`](testing.md) for patterns.

Convention: every public function has at least one Pester test covering the happy path + one common error case. New code without tests gets a `-NoTest` comment + a TODO.

## Style + conventions

- **PowerShell 7 compatibility.** Both 5.1 and 7 should work but 7 is the target for new code.
- **No backticks** for line continuation unless necessary; prefer pipeline-breaking or hashtable splatting.
- **Strings**: double-quoted for interpolation, single-quoted for literal. Use `${var}` delimiter when the next char could be confused as variable-name-continuation (we got bit by `$row:` once — see PR #11).
- **`@()` on function returns** when calling sites expect an array. Function return unwrapping is PowerShell's most common foot-gun.
- **`[DateTime]::MinValue` not `$null`** for `[ref] $dt` variables. PS 7 strict binding requires this.
- **Closure scoping for scriptblocks invoked from .NET callbacks** — capture functions as `${function:Name}` before constructing the scriptblock (PR #9).

These four rules cover the bugs the v1 ship caught during the first Pester pass.

## See also

- [`../concepts/architecture.md`](../concepts/architecture.md) — operator-facing.
- [`adding-a-module.md`](adding-a-module.md) — walkthrough.
- [`adding-an-ai-tool.md`](adding-an-ai-tool.md) — catalog contributions.
- [`adding-a-detector.md`](adding-a-detector.md) — incident-response detector additions.
- [`testing.md`](testing.md) — Pester patterns.
