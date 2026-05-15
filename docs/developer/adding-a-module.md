# Adding a new feature module

End-to-end walkthrough of adding a feature module to M365 Manager. We'll add a fictional "Distribution List Auditor" module that scans for DLs with no members in 180 days. The pattern generalizes to any new feature.

## Plan

| Step | Output |
|---|---|
| 1 | Plan public-facing surface (what menu slot, what functions, what audit shape) |
| 2 | Create the module file with the standard skeleton |
| 3 | Implement the functions, threading mutations through `Invoke-Action` |
| 4 | Wire reverse handlers in `Undo.ps1` for any reversible mutations |
| 5 | Add the module to `Main.ps1` load list at the right position |
| 6 | Add a menu entry (top-level slot OR existing submenu) |
| 7 | Write Pester tests |
| 8 | Add an AI tool catalog entry if appropriate |
| 9 | Document in `docs/guides/<area>.md` + reference updates |

## Step 1 — plan the public surface

Take the time. Once an audit `actionType` ships it's stable forever. Once a function name is public it can't change without a deprecation cycle.

For our example:

- **Public functions:**
  - `Get-StaleDistributionLists -DaysWithoutActivity 180` — read-only finder.
  - `Show-StaleDistributionListReport` — interactive viewer.
  - `Remove-StaleDistributionList -GroupId <id> -Reason <text>` — single removal.
  - `Invoke-BulkStaleDLCleanup -Path <csv>` — bulk path.
- **Audit shapes:**
  - `actionType=RemoveDistributionList`, `reverseType=...` — but DL deletion isn't reversible (memberships are lost). Mark `noUndoReason` instead.
- **Menu slot:** Slot 6 (Distribution List Management) already exists. Add as a submenu option.

## Step 2 — create the module

File: `DistributionListAuditor.ps1`. Standard skeleton:

```powershell
# ============================================================
#  DistributionListAuditor.ps1 -- stale DL detection + cleanup
#
#  Finds DLs with no member activity (sends / receives / member
#  additions) within the last N days. Surfaces in a CSV / HTML
#  report. Optional bulk cleanup with operator confirmation.
# ============================================================

function Get-StaleDistributionLists {
    <#
        Returns DLs with last-activity older than -DaysWithoutActivity.
        Activity = mail sent to the DL OR member added/removed.
        Falls back to "created date" when no UAL data exists.
    #>
    param(
        [int]$DaysWithoutActivity = 180,
        [int]$Max = 500
    )
    if (-not (Connect-ForTask 'DistributionLists')) { return @() }
    # ... implementation
}

function Show-StaleDistributionListReport {
    <# Interactive viewer. #>
    $stale = Get-StaleDistributionLists
    # render
}

function Remove-StaleDistributionList {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$Reason
    )
    Invoke-Action `
        -Description ("Delete distribution list {0} ({1})" -f $GroupId, $Reason) `
        -ActionType 'RemoveDistributionList' `
        -Target @{ groupId = $GroupId; reason = $Reason } `
        -NoUndoReason 'DL deletion is irreversible (membership list is lost; recreate manually).' `
        -Action {
            Remove-DistributionGroup -Identity $GroupId -Confirm:$false -ErrorAction Stop
            $true
        }
}

function Invoke-BulkStaleDLCleanup {
    param([Parameter(Mandatory)][string]$Path, [switch]$WhatIf)
    # Validate-first / per-row / result CSV pattern
}
```

## Step 3 — implement, threading mutations through Invoke-Action

The mutation surface is `Remove-StaleDistributionList`. Every state change goes through `Invoke-Action` (already shown above).

Read functions don't need `Invoke-Action` but DO call `Connect-ForTask` to ensure the right services are connected.

If your module needs a new connection area (e.g. SCC for compliance lookups + Graph + EXO), add it to `Auth.ps1`'s `Connect-ForTask` lookup table.

## Step 4 — wire reverse handlers

For our DL example, deletion is irreversible — use `-NoUndoReason` instead of `-ReverseType` and skip step 4.

For a reversible case (say `Disable-DistributionList`), add to `Undo.ps1`:

```powershell
$script:UndoHandlers['EnableDistributionList'] = {
    param([hashtable]$Target)
    Set-DistributionGroup -Identity $Target.groupId -ModerationEnabled $false -ErrorAction Stop
    $true
}
```

Pair this with the forward action emitting `-ReverseType 'EnableDistributionList'`.

## Step 5 — wire into Main.ps1

Add the file to the `$modules` array at the appropriate position. Foundation modules first, feature modules in the middle, AI layer last. For our example:

```powershell
$modules = @(
    # ... foundation ...
    "DistributionList.ps1",
    "DistributionListAuditor.ps1",     # <-- new, placed next to its peer module
    # ... rest ...
)
```

Don't put it too late (forward references won't resolve) or too early (dependencies on UI / Auth / Audit / Preview won't be loaded yet).

## Step 6 — add a menu entry

Two options:

### Option A — extend an existing submenu

In `DistributionList.ps1`'s `Start-DistributionListManagement`, add a new menu option that routes to `Show-StaleDistributionListReport`.

### Option B — add a top-level slot

If the feature is substantial enough to warrant its own top-level slot (like Phase 7's Incident Response), edit `Main.ps1`:

```powershell
$sel = Show-Menu -Title "Main Menu - Select a Task" -Options @(
    # ... existing ...
    "Distribution List Auditing..."   # new slot
) -BackLabel "Quit and Disconnect" -HiddenOptions @(99)

switch ($sel) {
    # ...
    24 { Start-DistributionListAuditingMenu }   # new dispatch
}
```

Then implement `Start-DistributionListAuditingMenu` in the new module.

## Step 7 — Pester tests

`tests/DistributionListAuditor.Tests.ps1`:

```powershell
BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'DistributionListAuditor.ps1')
}

Describe "Get-StaleDistributionLists" {
    It "returns the documented shape" {
        # Mock the EXO calls + assert on the shape of the output
    }
    It "honors -DaysWithoutActivity threshold" {
        # Run with -DaysWithoutActivity 30 vs 365, assert different counts
    }
}

Describe "Remove-StaleDistributionList" {
    It "writes an audit entry with noUndoReason" {
        # PREVIEW mode + verify the audit log line
    }
    It "uses Invoke-Action" {
        # Verify the call goes through the wrapper (not Remove-DistributionGroup directly)
    }
}
```

See [`testing.md`](testing.md) for the full pattern + mocking guidance.

## Step 8 — AI tool catalog (optional)

If the AI should be able to invoke these tools, add `ai-tools/dl-auditor.json`:

```jsonc
[
  {
    "name":        "Get-StaleDistributionLists",
    "description": "Find distribution lists with no member activity in the last N days. Read-only.",
    "parameters": {
      "type": "object",
      "properties": {
        "DaysWithoutActivity": { "type": "integer", "description": "Default 180" }
      }
    },
    "destructive":        false,
    "wrapInInvokeAction": false,
    "reverseTool":        null
  },
  {
    "name":        "Remove-StaleDistributionList",
    "description": "Delete a distribution list by id. Irreversible (membership history is lost).",
    "parameters": {
      "type": "object",
      "properties": {
        "GroupId": { "type": "string" },
        "Reason":  { "type": "string" }
      },
      "required": ["GroupId", "Reason"]
    },
    "destructive":             true,
    "requiresExplicitApproval": false,
    "wrapInInvokeAction":       false,
    "reverseTool":              null
  }
]
```

The dispatcher's default branch will resolve these by name. If the function needs parameter remapping, add an explicit case in `AIToolDispatch.Invoke-AIToolImpl`. See [`adding-an-ai-tool.md`](adding-an-ai-tool.md).

## Step 9 — documentation

Add:
- `docs/guides/distribution-list-auditing.md` — operator walkthrough.
- Update `docs/reference/cmdlets.md` with the new public functions.
- Update `docs/reference/menu-map.md` if you added a menu slot.
- Update `docs/reference/csv-formats.md` if you added a bulk CSV.

## Conventions checklist

Before opening a PR:

- [ ] Every mutation routes through `Invoke-Action`.
- [ ] Every reversible mutation has a `-ReverseType` + a `$UndoHandlers` entry.
- [ ] Every irreversible mutation has `-NoUndoReason`.
- [ ] `actionType` values are stable, snake-case-free, kebab-or-PascalCase, audit-filter-friendly.
- [ ] `target` hashtable keys are conventional (`userUpn`, `groupId`, `skuId`).
- [ ] Function returns are array-safe (`@(Get-X)` at call sites).
- [ ] `[DateTime]` `[ref]` vars are initialized to `[DateTime]::MinValue`.
- [ ] No `$varName:` interpolation that could confuse the PowerShell tokenizer.
- [ ] Pester tests cover the happy path + one error path per public function.
- [ ] Cross-references between new + existing docs are checked.
- [ ] `Invoke-Pester ./tests/` is green.

## See also

- [`architecture.md`](architecture.md) — codebase shape.
- [`adding-an-ai-tool.md`](adding-an-ai-tool.md) — AI integration.
- [`adding-a-detector.md`](adding-a-detector.md) — incident-response detector.
- [`testing.md`](testing.md) — Pester patterns.
- [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) — what to verify before claiming done.
