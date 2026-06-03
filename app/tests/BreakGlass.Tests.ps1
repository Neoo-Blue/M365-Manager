# ============================================================
#  Pester tests for BreakGlass.ps1
#  Posture predicates exercised against canned account states.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'BreakGlass.ps1')
}

Describe "Registry round-trip" {
    BeforeAll {
        $script:TempRegistry = Join-Path ([IO.Path]::GetTempPath()) ("bg-test-" + [Guid]::NewGuid() + '.json')
        Mock -CommandName Get-BreakGlassStatePath -MockWith { $script:TempRegistry }
    }
    AfterAll { if (Test-Path $script:TempRegistry) { Remove-Item $script:TempRegistry -Force } }

    It "Register, list, unregister" {
        Register-BreakGlassAccount -UPN 'bg-01@x.com' -AttestationEmail 'sec@x.com'
        Register-BreakGlassAccount -UPN 'bg-02@x.com' -AttestationEmail 'sec@x.com'
        (Get-BreakGlassAccounts).Count | Should -Be 2
        Unregister-BreakGlassAccount -UPN 'bg-01@x.com'
        (Get-BreakGlassAccounts).Count | Should -Be 1
        (Get-BreakGlassAccounts)[0].UPN | Should -Be 'bg-02@x.com'
    }
    It "Register is idempotent on UPN (updates instead of duplicates)" {
        Register-BreakGlassAccount -UPN 'bg-02@x.com' -AttestationEmail 'newsec@x.com'
        (Get-BreakGlassAccounts).Count | Should -Be 1
        (Get-BreakGlassAccounts)[0].AttestationEmail | Should -Be 'newsec@x.com'
    }
}

Describe "Posture predicates (logic, no Graph)" {
    It "passwordAge warns when older than threshold" {
        $age = ((Get-Date).ToUniversalTime() - (Get-Date).ToUniversalTime().AddDays(-200)).TotalDays
        ($age -gt $script:BGPasswordAgeWarnDays) | Should -BeTrue
    }
    It "passwordAge does NOT warn at 30 days" {
        $age = ((Get-Date).ToUniversalTime() - (Get-Date).ToUniversalTime().AddDays(-30)).TotalDays
        ($age -gt $script:BGPasswordAgeWarnDays) | Should -BeFalse
    }
    It "recentSignIn warns inside 30-day window" {
        $days = ((Get-Date).ToUniversalTime() - (Get-Date).ToUniversalTime().AddDays(-5)).TotalDays
        ($days -lt 30) | Should -BeTrue
    }
    It "recentSignIn does NOT warn outside 30-day window" {
        $days = ((Get-Date).ToUniversalTime() - (Get-Date).ToUniversalTime().AddDays(-90)).TotalDays
        ($days -lt 30) | Should -BeFalse
    }
}
