# ============================================================
#  Pester tests for AIPlanner.ps1
#
#  Plan parsing, structural validation, dependency-graph
#  topological order, and the rejection path through
#  Invoke-AIPlanApprovalFlow. Approval is forced via the
#  M365MGR_PLAN_APPROVAL env var so no operator interaction
#  is required.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')
    . (Join-Path $script:RepoRoot 'AIToolDispatch.ps1')
    . (Join-Path $script:RepoRoot 'AIPlanner.ps1')
    Get-AIToolCatalog -Reload | Out-Null
}

Describe "ConvertTo-PlanHashtable" {
    It "normalizes a PSCustomObject plan into a hashtable graph" {
        $raw = ([PSCustomObject]@{
            goal  = 'demo'
            steps = @(
                ([PSCustomObject]@{ id = 1; description='a'; tool='Get-Guests'; params = ([PSCustomObject]@{}); dependsOn = @() })
            )
        })
        $h = ConvertTo-PlanHashtable -PlanInput $raw
        $h           | Should -BeOfType [hashtable]
        $h.goal      | Should -Be 'demo'
        $h.steps[0]  | Should -BeOfType [hashtable]
        $h.steps[0].tool | Should -Be 'Get-Guests'
    }

    It "defaults missing dependsOn to an empty array" {
        $raw = ([PSCustomObject]@{ goal='x'; steps = @( ([PSCustomObject]@{ id=1; description='a'; tool='Get-Guests'; params = ([PSCustomObject]@{}) }) ) })
        $h = ConvertTo-PlanHashtable -PlanInput $raw
        @($h.steps[0].dependsOn).Count | Should -Be 0
    }
}

Describe "Test-PlanShape" {
    It "rejects an empty plan" {
        $r = Test-PlanShape -Plan @{ steps = @() }
        $r.Valid | Should -BeFalse
    }

    It "rejects an unknown tool name" {
        $r = Test-PlanShape -Plan @{
            steps = @(@{ id = 1; description='a'; tool='Get-DoesNotExist'; params=@{}; dependsOn=@() })
        }
        $r.Valid | Should -BeFalse
        ($r.Errors -join ' ') | Should -Match 'unknown tool'
    }

    It "rejects a forward dependency (step depends on later id)" {
        $r = Test-PlanShape -Plan @{
            steps = @(
                @{ id = 1; description='a'; tool='Get-Guests'; params=@{}; dependsOn = @(2) },
                @{ id = 2; description='b'; tool='Get-Guests'; params=@{}; dependsOn = @() }
            )
        }
        $r.Valid | Should -BeFalse
    }

    It "rejects a duplicate step id" {
        $r = Test-PlanShape -Plan @{
            steps = @(
                @{ id = 1; description='a'; tool='Get-Guests'; params=@{}; dependsOn=@() },
                @{ id = 1; description='b'; tool='Get-Guests'; params=@{}; dependsOn=@() }
            )
        }
        $r.Valid | Should -BeFalse
    }

    It "rejects meta tools used as plan steps" {
        $r = Test-PlanShape -Plan @{
            steps = @(@{ id = 1; description='a'; tool='submit_plan'; params=@{}; dependsOn=@() })
        }
        $r.Valid | Should -BeFalse
        ($r.Errors -join ' ') | Should -Match 'meta tool'
    }

    It "accepts a valid two-step plan with one dependency" {
        $r = Test-PlanShape -Plan @{
            steps = @(
                @{ id = 1; description='a'; tool='Get-Guests'; params=@{}; dependsOn=@() },
                @{ id = 2; description='b'; tool='Get-StaleGuests'; params=@{}; dependsOn=@(1) }
            )
        }
        $r.Valid | Should -BeTrue
    }
}

Describe "Get-TopologicalStepOrder" {
    It "returns steps in dependency-respecting order" {
        $steps = @(
            @{ id = 3; tool='t'; dependsOn = @(1,2) },
            @{ id = 1; tool='t'; dependsOn = @() },
            @{ id = 2; tool='t'; dependsOn = @(1) }
        )
        $ordered = Get-TopologicalStepOrder -Steps $steps
        @($ordered).Count | Should -Be 3
        $ordered[0].id | Should -Be 1
        $ordered[2].id | Should -Be 3
    }
}

Describe "Set-AIPlanMode" {
    It "round-trips through Get-AIPlanMode" {
        Set-AIPlanMode -Mode 'force'
        (Get-AIPlanMode) | Should -Be 'force'
        Set-AIPlanMode -Mode 'auto'
    }
}

Describe "Invoke-AIPlanApprovalFlow (rejection path)" {
    It "returns Status='rejected' under non-interactive default" {
        $env:M365MGR_PLAN_APPROVAL = 'reject'
        $script:NonInteractive = $true
        $plan = [PSCustomObject]@{
            goal  = 'demo'
            steps = @(([PSCustomObject]@{ id = 1; description='a'; tool='Get-Guests'; params=([PSCustomObject]@{}); dependsOn=@() }))
        }
        $r = Invoke-AIPlanApprovalFlow -PlanInput $plan -Config @{}
        $r.Status | Should -Be 'rejected'
        $script:NonInteractive = $false
        Remove-Item Env:\M365MGR_PLAN_APPROVAL -ErrorAction SilentlyContinue
    }
}
