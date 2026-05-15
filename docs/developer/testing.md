# Testing

Pester 5 patterns used in M365 Manager. 22 test suites cover the foundation + every feature module — no live Graph / EXO / SPO calls. Tests run against canned data + mocked SDK calls.

## Running tests

```powershell
# Full suite (currently 224 tests):
Invoke-Pester ./tests/

# One suite:
Invoke-Pester ./tests/Privacy.Tests.ps1

# One Describe block:
Invoke-Pester ./tests/IncidentResponse.Tests.ps1 -FullName "Get-IncidentSteps -- severity gating"

# With detailed output:
Invoke-Pester ./tests/ -Output Detailed
```

Tests run in a fresh `pwsh` process; they shouldn't affect each other.

## File structure

`tests/<Module>.Tests.ps1`. One file per module. Some shared fixtures live in `tests/fixtures/`.

Standard skeleton:

```powershell
BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    # Dot-source the dependency chain.
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'YourModule.ps1')

    # Redirect state to a temp dir so tests don't touch real state.
    $script:TempState = Join-Path ([IO.Path]::GetTempPath()) ("test-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempState -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempState
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempState -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Function-or-feature name" {
    It "describes a specific behavior" {
        # Arrange
        # Act
        # Assert
    }
}
```

## Patterns

### Mocking Graph SDK calls

Most Graph cmdlets are mocked with `Mock`:

```powershell
Describe "Get-UserAuthMethods" {
    BeforeEach {
        Mock Invoke-MgGraphRequest -MockWith {
            @{
                value = @(
                    @{ '@odata.type' = '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'; id = 'm1' },
                    @{ '@odata.type' = '#microsoft.graph.phoneAuthenticationMethod'; id = 'p1' }
                )
            }
        }
    }
    It "returns one entry per method" {
        $methods = Get-UserAuthMethods -User 'alice@contoso.com'
        @($methods).Count | Should -Be 2
    }
}
```

For a more complex mock that returns different data based on the URI:

```powershell
Mock Invoke-MgGraphRequest -MockWith {
    param($Method, $Uri)
    if ($Uri -like '*/users/alice*')           { return @{ id='u-001'; userPrincipalName='alice@contoso.com'; accountEnabled=$true } }
    elseif ($Uri -like '*/users/*/memberOf*')  { return @{ value=@(@{ id='g-001'; displayName='Group1' }) } }
    else                                       { throw "Unmocked URI: $Uri" }
}
```

### Mocking with parameter filters

`-ParameterFilter` lets the mock match specific argument combos:

```powershell
Mock Read-AuditEntries -ParameterFilter { -not $Path } -MockWith { @() }
Mock Read-AuditEntries -ParameterFilter { $Path } -MockWith { Get-Content $Path | ConvertFrom-Json }
```

**Watch out for recursive mocks.** PR #13 fixed a case where a mock body called the same function being mocked, recursing until stack overflow. Use `-ParameterFilter` to ensure the inner call doesn't match the same mock — see [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) for the case study.

### Mocking with state

For tests that need shared state across mocks:

```powershell
BeforeAll {
    $script:TempState = @{}
    Mock Read-UndoState -MockWith { $script:TempState }
    Mock Write-UndoState -MockWith { param([hashtable]$State) $script:TempState = $State }
}
It "round-trips state" {
    Write-UndoState -State @{ key = 'value' }
    (Read-UndoState).key | Should -Be 'value'
}
```

### Asserting on audit entries

```powershell
It "writes the right audit entry" {
    $script:CapturedEntries = New-Object System.Collections.ArrayList
    Mock Write-AuditEntry -MockWith {
        param($EventType, $Detail, $ActionType, $Target, $Result)
        [void]$script:CapturedEntries.Add(@{
            EventType = $EventType
            ActionType = $ActionType
            Target = $Target
            Result = $Result
        })
        return [guid]::NewGuid().ToString()
    }
    Remove-StaleDistributionList -GroupId 'g1' -Reason 'stale'
    $script:CapturedEntries[0].ActionType | Should -Be 'RemoveDistributionList'
    $script:CapturedEntries[0].Result | Should -Be 'success'
}
```

### Testing Invoke-Action wrapping

```powershell
It "routes through Invoke-Action" {
    Mock Invoke-Action -MockWith { param($Description, $Action) & $Action; $true } -Verifiable
    Remove-StaleDistributionList -GroupId 'g1' -Reason 'stale'
    Should -Invoke Invoke-Action -Times 1 -ParameterFilter { $ActionType -eq 'RemoveDistributionList' }
}
```

### Testing PREVIEW mode

Set the flag, run the function, verify nothing was actually called:

```powershell
It "doesn't call SDK in PREVIEW mode" {
    Set-PreviewMode -Enabled $true
    Mock Remove-DistributionGroup { } -Verifiable
    Remove-StaleDistributionList -GroupId 'g1' -Reason 'stale'
    Should -Invoke Remove-DistributionGroup -Times 0
    Set-PreviewMode -Enabled $false
}
```

## Common pitfalls

### `[ref]` types with `$null` on PS 7

```powershell
$ts = $null
[DateTime]::TryParse($input, [ref]$ts)  # FAILS on PS 7 -- "Cannot find overload"
```

Initialize to a typed default:

```powershell
$ts = [DateTime]::MinValue
[DateTime]::TryParse($input, [ref]$ts)  # works
```

### Hashtable vs PSCustomObject property access

```powershell
$h = @{ Location = 'US' }
$h.PSObject.Properties.Name -contains 'Location'  # FALSE -- enumerates Hashtable members, not keys
$h.ContainsKey('Location')                          # TRUE
$h.Location                                          # 'US' -- member access works either way
```

When test fixtures use hashtables and production code uses PSCustomObject (or vice versa), property access semantics differ. Build helpers that tolerate both — see `Get-IncidentSignInCountry` in `IncidentTriggers.ps1`.

### Closure scope when calling from .NET callbacks

A scriptblock invoked from a .NET delegate (e.g. `[System.Text.RegularExpressions.MatchEvaluator]`) runs OUTSIDE PowerShell's function-resolution scope. Functions defined at script scope aren't callable from inside the lambda. PR #9 fixed this by capturing the function reference:

```powershell
$tokenFn = ${function:Get-OrCreatePrivacyToken}
$evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    & $tokenFn -Value $m.Value -Type $type
}.GetNewClosure()
```

Tests that exercise these paths should explicitly invoke from a child scope (`& { ... }`) to catch the regression.

### Function return unwrapping

```powershell
function Get-Something { return @($singleItem) }
$x = Get-Something
$x.GetType()     # PSCustomObject -- the array got unwrapped to its single element!
$x += $other     # FAILS -- PSObject doesn't have op_Addition
```

Wrap call sites: `$x = @(Get-Something)`. PR #12 fixed three places this happened in production.

## Test inventory

22 suites, 224 tests:

```
AICostTracker.Tests.ps1
AIPlanner.Tests.ps1
AISessionStore.Tests.ps1
AIToolDispatch.Tests.ps1
AuditViewer.Tests.ps1
BreakGlass.Tests.ps1
GuestUsers.Tests.ps1
IncidentRegistry.Tests.ps1     (Phase 7)
IncidentResponse.Tests.ps1     (Phase 7)
IncidentTriggers.Tests.ps1     (Phase 7)
InvokeAcrossTenants.Tests.ps1
LicenseOptimizer.Tests.ps1
MFAManager.Tests.ps1
Notifications.Tests.ps1
OneDriveManager.Tests.ps1
Privacy.Tests.ps1
Scheduler.Tests.ps1
SharePoint.Tests.ps1
TeamsManager.Tests.ps1
TenantOverrides.Tests.ps1
TenantRegistry.Tests.ps1
Undo.Tests.ps1
```

## Pre-merge smoke checklist

Before claiming a PR is done:

- [ ] `Invoke-Pester ./tests/` is green.
- [ ] New public functions have at least one happy-path + one error-path test.
- [ ] Mutations are verified to route through `Invoke-Action`.
- [ ] PREVIEW mode is verified — calls don't fire in PREVIEW.
- [ ] Hashtable vs PSCustomObject paths both tested if either is plausible at the call site.
- [ ] No `$varName:` string interpolation that the PS tokenizer would mis-parse.
- [ ] `@()` wrap on function returns that callers append to.
- [ ] `[DateTime]::MinValue` init on `[ref]` DateTime vars.

## See also

- [`architecture.md`](architecture.md) — what's where.
- [`../operations/pre-merge-review.md`](../operations/pre-merge-review.md) — historical pre-merge issues caught.
- [`adding-a-module.md`](adding-a-module.md) — when adding a module that needs tests.
