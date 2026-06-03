# ============================================================
#  Pester tests for IncidentRegistry.ps1
#
#  Registry CRUD round-trip, tenant scoping (incidents in
#  tenant A don't show in tenant B), JSONL append correctness,
#  Get-Incident folding logic, Close-Incident state transitions.
#  No live Graph calls.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AuditViewer.ps1')
    . (Join-Path $script:RepoRoot 'Undo.ps1')
    . (Join-Path $script:RepoRoot 'IncidentResponse.ps1')
    . (Join-Path $script:RepoRoot 'IncidentRegistry.ps1')

    $script:TempState = Join-Path ([IO.Path]::GetTempPath()) ("ir-reg-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempState -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempState
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempState -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Registry append + Get-Incident folding" {
    It "Get-Incident returns null on a miss" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'RegTestEmpty' }
        try {
            (Get-Incident -Id 'INC-does-not-exist') | Should -BeNullOrEmpty
        } finally { $script:SessionState = $prevState }
    }

    It "Get-Incident folds running + completed records into one view, newest fields winning" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'RegTestFold' }
        try {
            $id = 'INC-2026-05-14-fold'
            Write-IncidentRegistryRecord -Record ([ordered]@{ id = $id; upn = 'u@x.com'; severity = 'High'; status = 'running'; startedUtc = '2026-05-14T10:00:00Z' })
            Write-IncidentRegistryRecord -Record ([ordered]@{ id = $id; upn = 'u@x.com'; severity = 'High'; status = 'completed'; startedUtc = '2026-05-14T10:00:00Z'; completedUtc = '2026-05-14T10:05:00Z'; reportPath = '/tmp/r.html' })
            $i = Get-Incident -Id $id
            $i              | Should -Not -BeNullOrEmpty
            $i.status       | Should -Be 'completed'
            $i.reportPath   | Should -Be '/tmp/r.html'
            # ConvertFrom-Json auto-parses ISO-8601 'Z' strings into
            # DateTime, so compare DateTime components instead of
            # exact string match (the [string] cast uses the local
            # culture's default format, which varies by host).
            ([DateTime]$i.startedUtc).ToUniversalTime().ToString('o')   | Should -Match '^2026-05-14T10:00:00'
            ([DateTime]$i.completedUtc).ToUniversalTime().ToString('o') | Should -Match '^2026-05-14T10:05:00'
        } finally { $script:SessionState = $prevState }
    }
}

Describe "Get-Incidents filters" {
    BeforeAll {
        $script:SessionStateBackup = $script:SessionState
        $script:SessionState = @{ TenantName = 'RegTestFilters' }
        # Three incidents across status / severity / age
        Write-IncidentRegistryRecord -Record ([ordered]@{ id = 'INC-a'; upn = 'a@x.com'; severity = 'High';     status = 'completed'; startedUtc = (Get-Date).AddDays(-1).ToUniversalTime().ToString('o') })
        Write-IncidentRegistryRecord -Record ([ordered]@{ id = 'INC-b'; upn = 'b@x.com'; severity = 'Critical'; status = 'closed';    startedUtc = (Get-Date).AddDays(-2).ToUniversalTime().ToString('o') })
        Write-IncidentRegistryRecord -Record ([ordered]@{ id = 'INC-c'; upn = 'c@x.com'; severity = 'Low';      status = 'completed'; startedUtc = (Get-Date).AddDays(-40).ToUniversalTime().ToString('o') })
    }
    AfterAll { $script:SessionState = $script:SessionStateBackup }

    It "Status=Open excludes closed" {
        $ids = @((Get-Incidents -Status Open) | ForEach-Object id)
        $ids | Should -Contain 'INC-a'
        $ids | Should -Not -Contain 'INC-b'
        $ids | Should -Contain 'INC-c'
    }
    It "Status=Closed includes only closed / false-positive" {
        $ids = @((Get-Incidents -Status Closed) | ForEach-Object id)
        $ids | Should -Contain 'INC-b'
        $ids | Should -Not -Contain 'INC-a'
    }
    It "Days filter is inclusive of the window" {
        $ids = @((Get-Incidents -Status All -Days 7) | ForEach-Object id)
        $ids | Should -Contain 'INC-a'
        $ids | Should -Contain 'INC-b'
        $ids | Should -Not -Contain 'INC-c'   # 40d ago
    }
    It "Severity filter narrows correctly" {
        $ids = @((Get-Incidents -Status All -Severity Critical) | ForEach-Object id)
        $ids | Should -Be @('INC-b')
    }
}

Describe "Tenant scoping" {
    It "incidents written under tenant A are not visible from tenant B" {
        $prevState = $script:SessionState
        try {
            $script:SessionState = @{ TenantName = 'TenantAlpha' }
            Write-IncidentRegistryRecord -Record ([ordered]@{ id = 'INC-tenant-a-1'; upn = 'a@alpha.com'; severity = 'High'; status = 'completed'; startedUtc = (Get-Date).ToUniversalTime().ToString('o') })
            $alphaIds = @((Get-Incidents -Status All) | ForEach-Object id)
            $alphaIds | Should -Contain 'INC-tenant-a-1'

            $script:SessionState = @{ TenantName = 'TenantBeta' }
            $betaIds = @((Get-Incidents -Status All) | ForEach-Object id)
            $betaIds | Should -Not -Contain 'INC-tenant-a-1'
        } finally { $script:SessionState = $prevState }
    }
}

Describe "Close-Incident" {
    It "appends a closed record + sets resolution" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'CloseTest1' }
        try {
            $id = 'INC-close-test-1'
            Write-IncidentRegistryRecord -Record ([ordered]@{ id = $id; upn = 'u@x.com'; severity = 'High'; status = 'completed'; startedUtc = (Get-Date).ToUniversalTime().ToString('o') })
            (Close-Incident -Id $id -Resolution 'test resolution') | Should -BeTrue
            $after = Get-Incident -Id $id
            $after.status     | Should -Be 'closed'
            $after.resolution | Should -Be 'test resolution'
        } finally { $script:SessionState = $prevState }
    }
    It "false-positive sets status='false-positive'" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'CloseTest2' }
        try {
            $id = 'INC-close-test-2'
            Write-IncidentRegistryRecord -Record ([ordered]@{ id = $id; upn = 'u@x.com'; severity = 'Low'; status = 'completed'; startedUtc = (Get-Date).ToUniversalTime().ToString('o') })
            (Close-Incident -Id $id -Resolution 'fp' -FalsePositive -SkipUndo) | Should -BeTrue
            $after = Get-Incident -Id $id
            $after.status        | Should -Be 'false-positive'
            $after.falsePositive | Should -BeTrue
        } finally { $script:SessionState = $prevState }
    }
    It "returns false for unknown id" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'CloseTestMissing' }
        try {
            (Close-Incident -Id 'INC-does-not-exist' -Resolution 'x') | Should -BeFalse
        } finally { $script:SessionState = $prevState }
    }
}

Describe "Get-IncidentList AI tool wrapper" {
    It "returns the {ok, result} shape" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'AIWrapTest' }
        try {
            $r = Get-IncidentList -Status All
            $r.ok | Should -BeTrue
            # .result is a hashtable; use .Keys not .PSObject.Properties.Name.
            $r.result.Keys | Should -Contain 'count'
            $r.result.Keys | Should -Contain 'incidents'
        } finally { $script:SessionState = $prevState }
    }
}

Describe "Summarize-AuditEvents gating" {
    It "refuses to run with default (UseAIForNarrative=Disabled)" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'NarrativeTest' }
        try {
            Remove-Item Function:Get-EffectiveConfig -ErrorAction SilentlyContinue
            function global:Get-EffectiveConfig { param([string]$Key) return $null }
            $r = Summarize-AuditEvents -IncidentId 'INC-anything'
            $r.ok    | Should -BeFalse
            $r.error | Should -Be 'ai_narrative_disabled'
        } finally {
            Remove-Item Function:Get-EffectiveConfig -ErrorAction SilentlyContinue
            $script:SessionState = $prevState
        }
    }
}
