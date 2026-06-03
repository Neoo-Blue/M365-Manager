# ============================================================
#  Pester tests for AICostTracker.ps1
#
#  Price lookup (exact / family / unknown), cost arithmetic,
#  monthly rollup, and budget-alert crossings. Uses a temp
#  state dir under $env:LOCALAPPDATA so the real cost log
#  isn't touched.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'AICostTracker.ps1')

    # Redirect cost storage to a per-test temp dir.
    $script:TempStateRoot = Join-Path ([IO.Path]::GetTempPath()) ("cost-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempStateRoot -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempStateRoot
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempStateRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Get-AIModelPrice" {
    It "returns exact match for a known Anthropic model" {
        $p = Get-AIModelPrice -Provider 'Anthropic' -Model 'claude-opus-4-7'
        $p.Source | Should -Be 'exact'
        $p.Input  | Should -Be 15.00
        $p.Output | Should -Be 75.00
    }

    It "falls back to a family wildcard for an unrecognized opus variant" {
        $p = Get-AIModelPrice -Provider 'Anthropic' -Model 'claude-opus-5-experimental'
        $p.Source | Should -Be 'family'
        $p.Input  | Should -Be 15.00
    }

    It "returns unknown / zero for a totally novel provider" {
        $p = Get-AIModelPrice -Provider 'TotallyNewProvider' -Model 'whatever'
        $p.Source | Should -Be 'unknown'
        $p.Input  | Should -Be 0.0
        $p.Output | Should -Be 0.0
    }

    It "prices Ollama at zero via the catch-all" {
        $p = Get-AIModelPrice -Provider 'Ollama' -Model 'llama3.1'
        $p.Input  | Should -Be 0.0
        $p.Output | Should -Be 0.0
    }
}

Describe "Add-AICostEvent" {
    BeforeEach {
        Reset-AICostSession
    }

    It "computes USD = (in/1M)*inputPrice + (out/1M)*outputPrice" {
        $r = Add-AICostEvent `
            -Config @{ Provider='Anthropic'; Model='claude-opus-4-7' } `
            -Usage  @{ InputTokens = 1000000; OutputTokens = 2000000 }
        # 1M in @ $15 + 2M out @ $75 = $15 + $150 = $165
        [math]::Round($r.Cost, 2) | Should -Be 165.00
        $r.PriceSource | Should -Be 'exact'
    }

    It "accumulates the running session total" {
        Add-AICostEvent -Config @{ Provider='Anthropic'; Model='claude-haiku-4-5' } -Usage @{ InputTokens = 1000; OutputTokens = 500 } | Out-Null
        $r = Add-AICostEvent -Config @{ Provider='Anthropic'; Model='claude-haiku-4-5' } -Usage @{ InputTokens = 2000; OutputTokens = 1000 }
        $state = Get-AICostState
        $state.SessionCalls     | Should -Be 2
        $state.SessionInTokens  | Should -Be 3000
        $state.SessionOutTokens | Should -Be 1500
        $r.CumulativeSession    | Should -BeGreaterThan 0
    }

    It "records zero cost for Ollama without alerting" {
        $r = Add-AICostEvent -Config @{ Provider='Ollama'; Model='llama3.1' } -Usage @{ InputTokens = 5000; OutputTokens = 3000 }
        $r.Cost       | Should -Be 0.0
        $r.AlertFired | Should -BeNullOrEmpty
    }

    It "fires a budget alert when crossing 50% / 80% / 100%" {
        # Set a tiny budget so a single call definitely crosses 100%
        $cfg = @{ Provider='Anthropic'; Model='claude-opus-4-7'; MonthlyBudgetUsd = 1.00; AlertAtPct = 80 }
        $r = Add-AICostEvent -Config $cfg -Usage @{ InputTokens = 100000; OutputTokens = 100000 }
        # Cost = 1.5 + 7.5 = $9 -> >= 100% of $1 budget
        $r.AlertFired       | Should -Not -BeNullOrEmpty
        $r.AlertFired.Pct   | Should -BeIn @(50, 80, 100)
    }
}

Describe "Show-AICostSummary" {
    It "runs without throwing on a fresh state" {
        Reset-AICostSession
        { Show-AICostSummary } | Should -Not -Throw
    }
}
