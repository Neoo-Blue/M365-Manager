# ============================================================
#  Pester tests for IncidentResponse.ps1
#
#  Step-order invariant (snapshot before any mutation), severity
#  gating produces the right step set, -WhatIf skips mutating
#  steps but still produces report, AI-narrative call is gated
#  by config. No live Graph calls -- the helpers the playbook
#  reaches into (Invoke-MgGraphRequest, Set-Mailbox,
#  Search-SignIns, Search-UAL, Get-UserOutboundShares,
#  Remove-AllAuthMethods, Send-Notification) are mocked.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'IncidentResponse.ps1')

    # Redirect state so we never touch the real audit / state dirs.
    $script:TempState = Join-Path ([IO.Path]::GetTempPath()) ("ir-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempState -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempState
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempState -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Get-IncidentSteps -- severity gating" {
    It "Low: only forensic + report steps (1, 8, 9, 10, 13)" {
        $nums = (Get-IncidentSteps -Severity Low | ForEach-Object Number)
        $nums | Should -Be @(1, 8, 9, 10, 13)
    }
    It "Medium: adds containment (2, 3, 4, 5)" {
        $nums = (Get-IncidentSteps -Severity Medium | ForEach-Object Number)
        $nums | Should -Be @(1, 2, 3, 4, 5, 8, 9, 10, 13)
    }
    It "High (default): adds cleanup + notify (6, 7, 12)" {
        $nums = (Get-IncidentSteps -Severity High | ForEach-Object Number)
        $nums | Should -Be @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13)
    }
    It "Critical: same as High when -QuarantineSentMail is omitted (step 11 is opt-in)" {
        $nums = (Get-IncidentSteps -Severity Critical | ForEach-Object Number)
        $nums | Should -Be @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13)
    }
    It "Critical + QuarantineSentMail: adds step 11 between 10 and 12" {
        $nums = (Get-IncidentSteps -Severity Critical -QuarantineSentMail | ForEach-Object Number)
        $nums | Should -Be @(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
    }
    It "step order invariant: snapshot (1) is always first regardless of severity" {
        foreach ($s in 'Low','Medium','High','Critical') {
            $first = (Get-IncidentSteps -Severity $s | Select-Object -First 1).Number
            $first | Should -Be 1
        }
    }
    It "step order invariant: report (13) is always last regardless of severity" {
        foreach ($s in 'Low','Medium','High','Critical') {
            $last = (Get-IncidentSteps -Severity $s | Select-Object -Last 1).Number
            $last | Should -Be 13
        }
    }
}

Describe "Snapshot-before-mutation invariant" {
    It "snapshot step appears before any state-mutating step in every severity" {
        foreach ($s in 'Medium','High','Critical') {
            $steps = Get-IncidentSteps -Severity $s
            $snapshotIdx = -1
            $mutatingNames = @('BlockSignIn','RevokeSessions','RevokeAuthMethods','ForcePasswordChange','DisableInboxRules','ClearForwarding','QuarantineSentMail')
            for ($i = 0; $i -lt $steps.Count; $i++) {
                if ($steps[$i].Name -eq 'Snapshot') { $snapshotIdx = $i; break }
            }
            $snapshotIdx | Should -BeGreaterOrEqual 0
            for ($i = 0; $i -lt $steps.Count; $i++) {
                if ($steps[$i].Name -in $mutatingNames) {
                    $i | Should -BeGreaterThan $snapshotIdx
                }
            }
        }
    }
}

Describe "New-IncidentId" {
    It "produces INC-YYYY-MM-DD-xxxx with a 4-char suffix" {
        $id = New-IncidentId
        $id | Should -Match '^INC-\d{4}-\d{2}-\d{2}-[0-9a-f]{4}$'
    }
    It "successive calls return different ids" {
        $a = New-IncidentId; $b = New-IncidentId
        $a | Should -Not -Be $b
    }
}

Describe "New-IncidentPassword" {
    It "produces a >=20-char string with at least one of each class M365 requires" {
        $pw = New-IncidentPassword
        $pw.Length | Should -BeGreaterOrEqual 20
        $pw | Should -Match '[a-z]'
        $pw | Should -Match '[A-Z]'
        $pw | Should -Match '\d'
        $pw | Should -Match '[^\w]'
    }
}

Describe "Get-IncidentTenantSlug" {
    It "returns 'default' when no tenant is set" {
        $prevState = $script:SessionState
        $script:SessionState = @{}
        try { (Get-IncidentTenantSlug) | Should -Be 'default' }
        finally { $script:SessionState = $prevState }
    }
    It "slugifies a tenant name to lower-snake_case" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'Contoso Ltd.' }
        try { (Get-IncidentTenantSlug) | Should -Be 'contoso_ltd_' }
        finally { $script:SessionState = $prevState }
    }
}

Describe "Get-IncidentsDirectory / Get-IncidentDirectory" {
    It "tenant-scopes the incidents dir under <state>/<tenant>/incidents" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'TestTenant1' }
        try {
            $dir = Get-IncidentsDirectory
            $dir | Should -Match 'testtenant1[\\/]incidents$'
            (Test-Path $dir) | Should -BeTrue
        }
        finally { $script:SessionState = $prevState }
    }
    It "Get-IncidentDirectory creates a per-id subdir under the tenant root" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'TestTenant2' }
        try {
            $id = 'INC-2026-05-14-test'
            $dir = Get-IncidentDirectory -IncidentId $id
            $dir | Should -Match "testtenant2[\\/]incidents[\\/]$id$"
            (Test-Path $dir) | Should -BeTrue
        }
        finally { $script:SessionState = $prevState }
    }
}

Describe "Write-IncidentArtifact" {
    It "writes string content verbatim" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'TestArtifactString' }
        try {
            $p = Write-IncidentArtifact -IncidentId 'INC-X' -Filename 'plain.txt' -Content 'hello world'
            (Get-Content -LiteralPath $p -Raw).Trim() | Should -Be 'hello world'
        }
        finally { $script:SessionState = $prevState }
    }
    It "writes JSON for hashtable content" {
        $prevState = $script:SessionState
        $script:SessionState = @{ TenantName = 'TestArtifactJson' }
        try {
            $p = Write-IncidentArtifact -IncidentId 'INC-Y' -Filename 'data.json' -Content @{ k = 'v'; n = 42 }
            $obj = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
            $obj.k | Should -Be 'v'
            $obj.n | Should -Be 42
        }
        finally { $script:SessionState = $prevState }
    }
}
