# ============================================================
#  Pester tests for TenantOverrides.ps1 (Phase 6 commit E)
#
#  Resolution order: global -> tenant file -> env var -> CLI
#  flag. Each test isolates the source it's testing so they
#  don't bleed into one another.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Notifications.ps1')
    . (Join-Path $script:RepoRoot 'TenantRegistry.ps1')
    . (Join-Path $script:RepoRoot 'TenantOverrides.ps1')

    $script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ("overrides-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempRoot -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempRoot
    Register-Tenant -Name 'Acme' -TenantId 'aaaa' | Out-Null
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Get-EffectiveConfig -- resolution order" {
    BeforeEach {
        # Clear any env override
        if (Test-Path Env:M365MGR_STALEGUESTDAYS) { Remove-Item Env:M365MGR_STALEGUESTDAYS }
        # Clear any tenant override file
        $f = Get-TenantOverrideFilePath -Name 'Acme'
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
        $script:CurrentTenantProfile = $null
    }

    It "returns the global value when no other source supplies the key" {
        $v = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig @{ StaleGuestDays = 90 }
        $v | Should -Be 90
    }

    It "tenant file overrides global" {
        Save-TenantOverrides -Name 'Acme' -Overrides @{ StaleGuestDays = 30 }
        Set-CurrentTenant -Name 'Acme' | Out-Null
        $v = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig @{ StaleGuestDays = 90 }
        [int]$v | Should -Be 30
    }

    It "env var overrides tenant file" {
        Save-TenantOverrides -Name 'Acme' -Overrides @{ StaleGuestDays = 30 }
        Set-CurrentTenant -Name 'Acme' | Out-Null
        $env:M365MGR_STALEGUESTDAYS = '14'
        $v = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig @{ StaleGuestDays = 90 }
        [int]$v | Should -Be 14
    }

    It "CLI value wins over everything" {
        Save-TenantOverrides -Name 'Acme' -Overrides @{ StaleGuestDays = 30 }
        Set-CurrentTenant -Name 'Acme' | Out-Null
        $env:M365MGR_STALEGUESTDAYS = '14'
        $v = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig @{ StaleGuestDays = 90 } -CliValue 7
        [int]$v | Should -Be 7
    }

    It "returns `$null when no source supplies the key" {
        $v = Get-EffectiveConfig -Key 'totally-unknown-key' -GlobalConfig @{}
        $v | Should -BeNullOrEmpty
    }

    It "treats empty-string CLI value as absent (doesn't clobber)" {
        $v = Get-EffectiveConfig -Key 'StaleGuestDays' -GlobalConfig @{ StaleGuestDays = 90 } -CliValue ''
        [int]$v | Should -Be 90
    }
}

Describe "Save-TenantOverrides / Get-TenantOverrides round-trip" {
    It "preserves nested types via JSON round-trip" {
        Save-TenantOverrides -Name 'Acme' -Overrides @{ Recipients = @('a@b','c@d'); Budget = 12.5 }
        $r = Get-TenantOverrides -Name 'Acme'
        @($r.Recipients).Count | Should -Be 2
        [double]$r.Budget      | Should -Be 12.5
    }
}

Describe "Get-TenantOverridableKeys" {
    It "exposes at least the documented core set" {
        $keys = Get-TenantOverridableKeys
        $keys | Should -Contain 'StaleGuestDays'
        $keys | Should -Contain 'AI.MonthlyBudgetUsd'
        $keys | Should -Contain 'Notifications.Recipients'
    }
}
