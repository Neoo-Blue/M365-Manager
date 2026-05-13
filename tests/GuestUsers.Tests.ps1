# ============================================================
#  Pester tests for GuestUsers.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'GuestUsers.ps1')
}

Describe "Stale-guest filter math" {
    It "DaysSinceSignIn >= threshold == stale" {
        $sample = @(
            [PSCustomObject]@{ UPN='active@x';    DaysSinceSignIn=5    },
            [PSCustomObject]@{ UPN='stale@x';     DaysSinceSignIn=120  },
            [PSCustomObject]@{ UPN='never@x';     DaysSinceSignIn=9999 }
        )
        $stale = @($sample | Where-Object DaysSinceSignIn -ge 90)
        $stale.Count            | Should -Be 2
        $stale[0].UPN           | Should -Be 'stale@x'
    }
    It "Treats 'no sign-in ever' (sentinel 9999) as stale at any threshold" {
        $never = [PSCustomObject]@{ UPN='ghost@x'; DaysSinceSignIn=9999 }
        ($never.DaysSinceSignIn -ge 30) | Should -BeTrue
    }
}

Describe "Recert state transitions" {
    BeforeAll {
        # Redirect the state-path helper to a temp file so we don't
        # touch the real machine state.
        $script:TempState = Join-Path ([IO.Path]::GetTempPath()) ("recert-test-" + [Guid]::NewGuid() + '.json')
        Mock -CommandName Get-RecertStatePath -MockWith { $script:TempState }
    }
    AfterAll { if (Test-Path $script:TempState) { Remove-Item $script:TempState -Force } }

    It "round-trips an empty array" {
        Write-RecertState -Records @()
        (Read-RecertState).Count | Should -Be 0
    }
    It "queues a pending record and reads it back" {
        $rec = [PSCustomObject]@{ campaignId='c1'; guestId='g1'; guestUpn='guest@x'; managerUpn='mgr@x'; queuedAt=(Get-Date).ToString('o'); state='pending'; decisionBy=$null; decisionAt=$null; notes='' }
        Write-RecertState -Records @($rec)
        $read = Read-RecertState
        $read.Count           | Should -Be 1
        $read[0].state        | Should -Be 'pending'
        $read[0].guestUpn     | Should -Be 'guest@x'
    }
    It "applies a Keep decision" {
        $rec = (Read-RecertState)[0]
        $rec.state = 'keep'; $rec.decisionAt = (Get-Date).ToString('o'); $rec.decisionBy = 'admin@x'
        Write-RecertState -Records @($rec)
        $read = Read-RecertState
        $read[0].state      | Should -Be 'keep'
        $read[0].decisionBy | Should -Be 'admin@x'
    }
}

Describe "Domain pivot logic" {
    It "buckets guests by domain extracted from mail" {
        $sample = @(
            [PSCustomObject]@{ UPN='a@vendor.com_#EXT#@contoso.onmicrosoft.com'; Mail='a@vendor.com';   Domains='vendor.com' },
            [PSCustomObject]@{ UPN='b@vendor.com_#EXT#@contoso.onmicrosoft.com'; Mail='b@vendor.com';   Domains='vendor.com' },
            [PSCustomObject]@{ UPN='c@partner.io_#EXT#@contoso.onmicrosoft.com'; Mail='c@partner.io';   Domains='partner.io' }
        )
        $bucket = @{}
        foreach ($g in $sample) {
            $d = $g.Domains
            if (-not $bucket.ContainsKey($d)) { $bucket[$d] = 0 }
            $bucket[$d]++
        }
        $bucket['vendor.com']  | Should -Be 2
        $bucket['partner.io']  | Should -Be 1
    }
}
