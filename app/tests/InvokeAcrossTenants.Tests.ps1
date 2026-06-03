# ============================================================
#  Pester tests for Invoke-AcrossTenants (MSPReports.ps1, Phase 6)
#
#  Sequential result-shape contract, try/finally tenant restore
#  when the scriptblock throws, per-step audit attribution.
#  Switch-Tenant is mocked so no Graph / EXO calls run.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Notifications.ps1')
    . (Join-Path $script:RepoRoot 'TenantRegistry.ps1')
    . (Join-Path $script:RepoRoot 'TenantSwitch.ps1')
    . (Join-Path $script:RepoRoot 'MSPReports.ps1')

    $script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ("xtenant-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempRoot -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempRoot

    # Seed registry with three tenants.
    Register-Tenant -Name 'Acme'    -TenantId 'aaaa' -PrimaryDomain 'acme.onmicrosoft.com'    | Out-Null
    Register-Tenant -Name 'Contoso' -TenantId 'bbbb' -PrimaryDomain 'contoso.onmicrosoft.com' | Out-Null
    Register-Tenant -Name 'Fabrikam'-TenantId 'cccc' -PrimaryDomain 'fabrikam.onmicrosoft.com'| Out-Null

    # Stub Switch-Tenant so no real connect runs. Replaces the real
    # function in this Describe's scope only; restored by AfterAll.
    function global:Switch-Tenant {
        param([string]$Name, [switch]$NoReconnect)
        $t = Get-Tenant -Name $Name
        if (-not $t) { return $false }
        $script:CurrentTenantProfile = $t
        $script:SessionState.TenantId     = $t.tenantId
        $script:SessionState.TenantName   = $t.name
        $script:SessionState.TenantDomain = $t.primaryDomain
        return $true
    }
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue
    if (Get-Command Switch-Tenant -ErrorAction SilentlyContinue) { Remove-Item function:\Switch-Tenant -ErrorAction SilentlyContinue }
}

Describe "Invoke-AcrossTenants -- happy path" {
    It "runs the scriptblock once per tenant" {
        $r = Invoke-AcrossTenants -Tenants @('Acme','Contoso') -Script { return $args[0].name }
        @($r).Count | Should -Be 2
        $r[0].Tenant | Should -Be 'Acme'
        $r[1].Tenant | Should -Be 'Contoso'
        $r[0].Result | Should -Be 'Acme'
    }

    It "honors @all" {
        $r = Invoke-AcrossTenants -Tenants @('@all') -Script { return 1 }
        @($r).Count | Should -Be 3
    }

    It "skips unknown tenant names with a warning" {
        $r = Invoke-AcrossTenants -Tenants @('Acme','ghost') -Script { return 1 }
        @($r).Count | Should -Be 1
        $r[0].Tenant | Should -Be 'Acme'
    }
}

Describe "Invoke-AcrossTenants -- result shape contract" {
    It "returns the documented PSCustomObject shape with all required fields" {
        $r = Invoke-AcrossTenants -Tenants @('Acme') -Script { return 'x' }
        $r[0].PSObject.Properties.Name | Should -Contain 'Tenant'
        $r[0].PSObject.Properties.Name | Should -Contain 'TenantId'
        $r[0].PSObject.Properties.Name | Should -Contain 'Success'
        $r[0].PSObject.Properties.Name | Should -Contain 'Result'
        $r[0].PSObject.Properties.Name | Should -Contain 'Error'
        $r[0].PSObject.Properties.Name | Should -Contain 'DurationMs'
        $r[0].Success      | Should -BeTrue
        $r[0].DurationMs   | Should -BeGreaterOrEqual 0
    }
}

Describe "Invoke-AcrossTenants -- error handling" {
    It "captures scriptblock exceptions per tenant without aborting the run" {
        $r = Invoke-AcrossTenants -Tenants @('Acme','Contoso') -Script {
            if ($args[0].name -eq 'Acme') { throw "boom" }
            return 'ok'
        }
        @($r).Count | Should -Be 2
        $r[0].Success | Should -BeFalse
        $r[0].Error   | Should -Match 'boom'
        $r[1].Success | Should -BeTrue
    }
}

Describe "Invoke-AcrossTenants -- tenant restore (try/finally)" {
    It "restores the prior tenant even when the scriptblock throws" {
        # Seed the 'prior' tenant first
        Set-CurrentTenant -Name 'Fabrikam' | Out-Null
        try {
            Invoke-AcrossTenants -Tenants @('Acme') -Script { throw "boom" } | Out-Null
        } catch {}
        (Get-CurrentTenant).name | Should -Be 'Fabrikam'
        $script:SessionState.TenantName | Should -Be 'Fabrikam'
    }
}
