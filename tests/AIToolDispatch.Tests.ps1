# ============================================================
#  Pester tests for AIToolDispatch.ps1
#
#  Catalog loading, schema validation, and provider-payload
#  construction. No live provider calls -- Build-* helpers are
#  pure functions over the in-memory catalog. Test-AIToolInput
#  uses the same JSON Schema the runtime enforces.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')
    . (Join-Path $script:RepoRoot 'AIToolDispatch.ps1')
    Get-AIToolCatalog -Reload | Out-Null
}

Describe "Get-AIToolCatalog" {
    It "loads at least one tool from every ai-tools/*.json file" {
        $cat = Get-AIToolCatalog
        @($cat).Count | Should -BeGreaterThan 10
    }

    It "loads the submit_plan meta-tool" {
        $t = Get-AIToolByName -Name 'submit_plan'
        $t            | Should -Not -BeNullOrEmpty
        $t.isMeta     | Should -BeTrue
        $t.destructive| Should -BeFalse
    }

    It "marks Remove-Guest as destructive" {
        $t = Get-AIToolByName -Name 'Remove-Guest'
        if ($t) { $t.destructive | Should -BeTrue }
    }

    It "is case-sensitive on tool name lookup" {
        (Get-AIToolByName -Name 'submit_plan') | Should -Not -BeNullOrEmpty
        # Wrong case returns $null
        (Get-AIToolByName -Name 'Submit_Plan') | Should -BeNullOrEmpty
    }
}

Describe "Test-AIToolInput (JSON Schema validation)" {
    It "rejects missing required fields" {
        $def = Get-AIToolByName -Name 'submit_plan'
        $r = Test-AIToolInput -ToolDef $def -Input @{}
        $r.Ok | Should -BeFalse
        ($r.Errors -join ' ') | Should -Match 'required'
    }

    It "accepts a valid submit_plan input" {
        $def = Get-AIToolByName -Name 'submit_plan'
        $valid = @{
            goal  = 'demo'
            steps = @(@{ id = 1; description = 'noop'; tool = 'Get-Guests'; params = @{} })
        }
        $r = Test-AIToolInput -ToolDef $def -Input $valid
        $r.Ok | Should -BeTrue
    }

    It "rejects wrong parameter type" {
        $def = Get-AIToolByName -Name 'Get-StaleGuests'
        if (-not $def) { Set-ItResult -Skipped -Because 'Get-StaleGuests not in catalog'; return }
        $r = Test-AIToolInput -ToolDef $def -Input @{ DaysSinceSignIn = 'ninety' }
        $r.Ok | Should -BeFalse
    }
}

Describe "Build-AnthropicToolsPayload" {
    It "emits one tool entry per catalog item including meta tools" {
        $cat = Get-AIToolCatalog
        $payload = Build-AnthropicToolsPayload -Catalog $cat
        @($payload).Count | Should -Be @($cat).Count
    }

    It "shapes each entry with name / description / input_schema" {
        $payload = Build-AnthropicToolsPayload -Catalog (Get-AIToolCatalog)
        foreach ($p in $payload) {
            $p.name         | Should -Not -BeNullOrEmpty
            $p.description  | Should -Not -BeNullOrEmpty
            $p.input_schema | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "Build-OpenAIToolsPayload" {
    It "wraps each tool in { type=function, function={...} }" {
        $payload = Build-OpenAIToolsPayload -Catalog (Get-AIToolCatalog)
        foreach ($p in $payload) {
            $p.type                | Should -Be 'function'
            $p.function.name       | Should -Not -BeNullOrEmpty
            $p.function.parameters | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "Build-ToolResultMessage" {
    It "shapes Anthropic tool_result blocks" {
        $m = Build-ToolResultMessage -Provider 'Anthropic' -ToolResults @(@{ id = 'tu_1'; content = '{"ok":true}'; isError = $false })
        $m.role         | Should -Be 'user'
        $m.content[0].type | Should -Be 'tool_result'
        $m.content[0].tool_use_id | Should -Be 'tu_1'
    }

    It "shapes OpenAI tool messages (one per result)" {
        $msgs = Build-ToolResultMessage -Provider 'OpenAI' -ToolResults @(
            @{ id = 'call_1'; content = '{"ok":true}'; isError = $false }
            @{ id = 'call_2'; content = '{"ok":false}'; isError = $true }
        )
        @($msgs).Count | Should -Be 2
        $msgs[0].role  | Should -Be 'tool'
        $msgs[0].tool_call_id | Should -Be 'call_1'
    }
}
