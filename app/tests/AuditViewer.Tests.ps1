# ============================================================
#  Pester tests for AuditViewer.ps1
#  Run from repo root: Invoke-Pester ./tests/
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AuditViewer.ps1')
    $script:FixturePath = Join-Path $PSScriptRoot 'fixtures/audit-mixed.jsonl'
}

Describe "ConvertFrom-AuditLine" {
    It "parses a JSONL line into a normalized record" {
        $line = '{"ts":"2026-05-10T09:00:05.123Z","entryId":"abc","mode":"LIVE","event":"EXEC","description":"X","actionType":"AssignLicense","target":{"userId":"u1"},"result":"success","error":null,"reverse":null,"noUndoReason":null}'
        $r = ConvertFrom-AuditLine -Line $line
        $r                  | Should -Not -BeNullOrEmpty
        $r.entryId          | Should -Be 'abc'
        $r.event            | Should -Be 'EXEC'
        $r.actionType       | Should -Be 'AssignLicense'
        $r.result           | Should -Be 'success'
        $r.source           | Should -Be 'jsonl'
        $r.target.userId    | Should -Be 'u1'
    }
    It "parses a legacy session line with MODE= tag" {
        $line = '[2026-05-09 18:42:15.001] [EXEC] [MODE=LIVE] Pre-Phase-2 line: did something'
        $r = ConvertFrom-AuditLine -Line $line
        $r.source           | Should -Be 'legacy-session'
        $r.event            | Should -Be 'EXEC'
        $r.mode             | Should -Be 'LIVE'
        $r.description      | Should -Match 'did something'
        $r.actionType       | Should -BeNullOrEmpty
    }
    It "parses an AI mark log line (no MODE= tag)" {
        $line = '[2026-05-09 18:42:15.001] [PROPOSE] Get-MgUser -Search "displayName:jackie"'
        $r = ConvertFrom-AuditLine -Line $line
        $r.source           | Should -Be 'ai-mark'
        $r.event            | Should -Be 'PROPOSE'
        $r.actionType       | Should -Be 'AICmd'
    }
    It "returns null on garbage input" {
        (ConvertFrom-AuditLine -Line '')       | Should -BeNullOrEmpty
        (ConvertFrom-AuditLine -Line 'random') | Should -BeNullOrEmpty
        (ConvertFrom-AuditLine -Line '{ not valid json') | Should -BeNullOrEmpty
    }
}

Describe "Read-AuditEntries (mixed fixture)" {
    It "reads JSONL + legacy lines + sorts by timestamp ascending" {
        $entries = Read-AuditEntries -Path @($script:FixturePath)
        $entries.Count          | Should -BeGreaterOrEqual 8
        $entries[0].ts          | Should -BeLessOrEqual $entries[1].ts
        ($entries | Where-Object source -eq 'jsonl').Count          | Should -Be 6
        ($entries | Where-Object source -eq 'legacy-session').Count | Should -Be 2
    }
}

Describe "Filter-AuditEntries" {
    BeforeAll {
        $script:All = Read-AuditEntries -Path @($script:FixturePath)
    }
    It "filters by UPN (target match)" {
        $hits = Filter-AuditEntries -Entries $script:All -Filter @{ User = 'jane@contoso.com' }
        $hits.Count             | Should -BeGreaterOrEqual 2
        $hits | ForEach-Object { Test-AuditEntryMatchesUser -Entry $_ -Upn 'jane@contoso.com' | Should -BeTrue }
    }
    It "filters by ActionType" {
        (Filter-AuditEntries -Entries $script:All -Filter @{ ActionType = 'AssignLicense' }).Count | Should -Be 1
        (Filter-AuditEntries -Entries $script:All -Filter @{ ActionType = 'AddToGroup' }).Count    | Should -Be 2
    }
    It "filters by Result" {
        (Filter-AuditEntries -Entries $script:All -Filter @{ Result = 'success' }).Count | Should -BeGreaterOrEqual 3
        (Filter-AuditEntries -Entries $script:All -Filter @{ Result = 'failure' }).Count | Should -Be 1
        (Filter-AuditEntries -Entries $script:All -Filter @{ Result = 'preview' }).Count | Should -Be 1
    }
    It "filters by date range" {
        $hits = Filter-AuditEntries -Entries $script:All -Filter @{ From = ([DateTime]'2026-05-10T00:00:00Z'); To = ([DateTime]'2026-05-11T00:00:00Z') }
        $hits | ForEach-Object {
            ($_.ts -ge ([DateTime]'2026-05-10T00:00:00Z')) | Should -BeTrue
            ($_.ts -le ([DateTime]'2026-05-11T00:00:00Z')) | Should -BeTrue
        }
    }
    It "Target substring matches both description and structured target values" {
        (Filter-AuditEntries -Entries $script:All -Filter @{ Target = 'SG-Sales' }).Count | Should -BeGreaterOrEqual 1
        (Filter-AuditEntries -Entries $script:All -Filter @{ Target = 'POWER_BI_PRO' }).Count | Should -BeGreaterOrEqual 1
    }
}

Describe "Export-AuditEntriesCsv / Export-AuditEntriesHtml" {
    BeforeAll {
        $script:All = Read-AuditEntries -Path @($script:FixturePath)
        $script:OutDir = Join-Path ([IO.Path]::GetTempPath()) ("m365-audit-test-" + [Guid]::NewGuid())
        New-Item -ItemType Directory -Path $script:OutDir -Force | Out-Null
    }
    AfterAll {
        if (Test-Path $script:OutDir) { Remove-Item $script:OutDir -Recurse -Force }
    }
    It "CSV export round-trips through Import-Csv" {
        $csv = Join-Path $script:OutDir 'out.csv'
        Export-AuditEntriesCsv -Entries $script:All -Path $csv
        Test-Path $csv | Should -BeTrue
        $rows = @(Import-Csv -LiteralPath $csv)
        $rows.Count | Should -Be $script:All.Count
        $rows[0].PSObject.Properties.Name -contains 'Timestamp' | Should -BeTrue
        $rows[0].PSObject.Properties.Name -contains 'ActionType' | Should -BeTrue
    }
    It "HTML export contains the expected table and a row per entry" {
        $html = Join-Path $script:OutDir 'out.html'
        Export-AuditEntriesHtml -Entries $script:All -Path $html -Filter @{}
        Test-Path $html | Should -BeTrue
        $text = Get-Content -LiteralPath $html -Raw
        $text | Should -Match '<table>'
        $text | Should -Match 'M365 Manager audit export'
        ([regex]::Matches($text, '<tr')).Count | Should -BeGreaterOrEqual ($script:All.Count + 1)  # +1 header
    }
}
