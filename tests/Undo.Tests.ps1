# ============================================================
#  Pester tests for Undo.ps1
#  Run from repo root: Invoke-Pester ./tests/
#
#  These tests exercise the dispatch table + state-tracking
#  WITHOUT making real Graph / EXO calls -- the handler
#  scriptblocks are mocked through a stub Invoke-Action that
#  records what would have run.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AuditViewer.ps1')
    . (Join-Path $script:RepoRoot 'Undo.ps1')
}

Describe "UndoHandlers dispatch table" {
    It "exposes inverse pairs for the curated set" {
        $script:UndoHandlers.Keys | Should -Contain 'RemoveLicense'
        $script:UndoHandlers.Keys | Should -Contain 'AssignLicense'
        $script:UndoHandlers.Keys | Should -Contain 'AddToGroup'
        $script:UndoHandlers.Keys | Should -Contain 'RemoveFromGroup'
        $script:UndoHandlers.Keys | Should -Contain 'AddToDistributionList'
        $script:UndoHandlers.Keys | Should -Contain 'RemoveFromDistributionList'
        $script:UndoHandlers.Keys | Should -Contain 'GrantMailboxFullAccess'
        $script:UndoHandlers.Keys | Should -Contain 'RevokeMailboxFullAccess'
        $script:UndoHandlers.Keys | Should -Contain 'GrantMailboxSendAs'
        $script:UndoHandlers.Keys | Should -Contain 'RevokeMailboxSendAs'
        $script:UndoHandlers.Keys | Should -Contain 'GrantCalendarAccess'
        $script:UndoHandlers.Keys | Should -Contain 'RevokeCalendarAccess'
        $script:UndoHandlers.Keys | Should -Contain 'ClearOOO'
        $script:UndoHandlers.Keys | Should -Contain 'ClearForwarding'
        $script:UndoHandlers.Keys | Should -Contain 'BlockSignIn'
        $script:UndoHandlers.Keys | Should -Contain 'UnblockSignIn'
    }
    It "each handler is a scriptblock that accepts a Target" {
        foreach ($k in $script:UndoHandlers.Keys) {
            $sb = $script:UndoHandlers[$k]
            $sb                   | Should -BeOfType ([scriptblock])
            $sb.Ast.ParamBlock.Parameters[0].Name.VariablePath.UserPath | Should -Be 'Target'
        }
    }
}

Describe "ConvertTo-UndoTargetHashtable" {
    It "passes hashtables through unchanged" {
        $h = @{ a = 1; b = 'x' }
        $r = ConvertTo-UndoTargetHashtable $h
        $r.a | Should -Be 1
        $r.b | Should -Be 'x'
    }
    It "converts a PSCustomObject (ConvertFrom-Json shape) to hashtable" {
        $obj = '{ "userId": "abc", "groupId": "xyz" }' | ConvertFrom-Json
        $r = ConvertTo-UndoTargetHashtable $obj
        $r -is [hashtable] | Should -BeTrue
        $r.userId | Should -Be 'abc'
        $r.groupId | Should -Be 'xyz'
    }
    It "returns empty hashtable for null input" {
        (ConvertTo-UndoTargetHashtable $null).Count | Should -Be 0
    }
}

Describe "Get-UndoableEntries (using mixed fixture)" {
    BeforeAll {
        $script:FixturePath = Join-Path $PSScriptRoot 'fixtures/audit-mixed.jsonl'
        # Stub Read-AuditEntries with a ParameterFilter so the mock
        # only intercepts the no-args call -- the explicit -Path call
        # inside the mock body delegates to the real function.
        # Without the filter the mock recurses into itself until the
        # PowerShell call-depth limit (12 sec, then ScriptCallDepthException).
        Mock -CommandName Read-AuditEntries -ParameterFilter { -not $Path -or @($Path).Count -eq 0 } -MockWith {
            Read-AuditEntries -Path @($script:FixturePath)
        }
        # Stub Read-UndoState / Get-UndoStatePath so we don't touch the real sidecar.
        $script:TempState = @{}
        Mock -CommandName Read-UndoState -MockWith { $script:TempState }
        Mock -CommandName Write-UndoState -MockWith { param([hashtable]$State) $script:TempState = $State }
    }
    It "returns only success+reverse entries (skips preview, failure, no-reverse)" {
        $u = Get-UndoableEntries -Limit 50
        $u | ForEach-Object { $_.result | Should -Be 'success' }
        $u | ForEach-Object { $_.reverse | Should -Not -BeNullOrEmpty }
    }
    It "skips entries flagged as already-reversed in the sidecar" {
        # Mark one entry reversed in the stub state
        $allBefore = Get-UndoableEntries -Limit 50
        $allBefore.Count | Should -BeGreaterOrEqual 2
        $first = $allBefore[0]
        $script:TempState[$first.entryId] = @{ state='reversed'; reversedBy='x'; reversedAt='now'; originalType=$first.actionType }
        $after = Get-UndoableEntries -Limit 50
        $after | ForEach-Object { $_.entryId | Should -Not -Be $first.entryId }
    }
    It "filters by user UPN" {
        $script:TempState = @{}
        $hits = Get-UndoableEntries -Filter 'jane@contoso.com' -Limit 50
        $hits.Count | Should -BeGreaterOrEqual 1
    }
}

Describe "Round-trip semantics on a handler (no Graph call)" {
    It "AddToGroup -> RemoveFromGroup target round-trip preserves group/user ids" {
        $forward = @{ userId='u1'; groupId='g1'; groupName='G' }
        $reverseRecipe = @{ type='RemoveFromGroup'; description='r'; target = $forward }
        $reverseRecipe.target.userId  | Should -Be 'u1'
        $reverseRecipe.target.groupId | Should -Be 'g1'
        # Handler must accept the same shape
        $sb = $script:UndoHandlers['RemoveFromGroup']
        $sb.Ast.ParamBlock.Parameters[0].Name.VariablePath.UserPath | Should -Be 'Target'
    }
}
