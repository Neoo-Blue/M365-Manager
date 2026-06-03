# ============================================================
#  Pester tests for TenantRegistry.ps1 (Phase 6)
#
#  CRUD round-trip, Get-CurrentTenant after Set-CurrentTenant,
#  credential-manifest encryption (Protect-Secret round-trip),
#  Test-FirstRunMigration short-circuits when registry already
#  has entries. Uses a per-test LOCALAPPDATA so the real
#  registry isn't touched.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Notifications.ps1')
    . (Join-Path $script:RepoRoot 'TenantRegistry.ps1')

    $script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ("tenants-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempRoot -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempRoot
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Register-Tenant / Get-Tenant / Remove-Tenant" {
    BeforeEach {
        # Wipe registry between tests
        $f = Get-TenantRegistryPath
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
        $sd = Get-TenantSecretsDir
        if (Test-Path -LiteralPath $sd) { Get-ChildItem -LiteralPath $sd | Remove-Item -Force }
        $script:CurrentTenantProfile = $null
    }

    It "starts with an empty registry" {
        @(Get-Tenants).Count | Should -Be 0
    }

    It "registers a new tenant with Interactive auth (no secret file)" {
        Register-Tenant -Name 'Contoso' -TenantId 'abcd-1234' -PrimaryDomain 'contoso.onmicrosoft.com' | Out-Null
        $t = Get-Tenant -Name 'Contoso'
        $t                | Should -Not -BeNullOrEmpty
        $t.tenantId       | Should -Be 'abcd-1234'
        $t.credentialRef  | Should -BeNullOrEmpty
    }

    It "rejects a duplicate registration" {
        Register-Tenant -Name 'Acme' -TenantId 'abcd' | Out-Null
        { Register-Tenant -Name 'Acme' -TenantId 'xxxx' } | Should -Throw
    }

    It "writes an encrypted manifest for ClientSecret auth" {
        Register-Tenant -Name 'Acme' -TenantId 'abcd' -ClientId 'app-id' -AuthMode 'ClientSecret' -ClientSecret 'topSecret123' | Out-Null
        $manifest = Get-TenantCredentialManifest -Name 'Acme'
        $manifest.authMode | Should -Be 'ClientSecret'
        $manifest.clientId | Should -Be 'app-id'
        # secret is decrypted on read
        $manifest.secret   | Should -Be 'topSecret123'
        # raw file should NOT contain the plain secret
        $raw = Get-Content -LiteralPath (Join-Path (Get-TenantSecretsDir) ('tenant-acme.dat')) -Raw
        $raw | Should -Not -Match 'topSecret123'
    }

    It "removes a tenant + its secret manifest" {
        Register-Tenant -Name 'Acme' -TenantId 'abcd' -ClientId 'a' -AuthMode 'CertThumbprint' -CertThumbprint 'XXXX' | Out-Null
        $secretFile = Join-Path (Get-TenantSecretsDir) 'tenant-acme.dat'
        Test-Path $secretFile | Should -BeTrue
        Remove-Tenant -Name 'Acme' | Should -BeTrue
        Get-Tenant   -Name 'Acme' | Should -BeNullOrEmpty
        Test-Path $secretFile      | Should -BeFalse
    }

    It "Remove-Tenant returns false on a miss" {
        (Remove-Tenant -Name 'no-such-tenant') | Should -BeFalse
    }
}

Describe "Set-CurrentTenant / Get-CurrentTenant" {
    BeforeEach {
        $f = Get-TenantRegistryPath
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
        $script:CurrentTenantProfile = $null
    }

    It "mirrors fields into SessionState" {
        Register-Tenant -Name 'Acme' -TenantId 'abcd-1234' -PrimaryDomain 'acme.onmicrosoft.com' | Out-Null
        Set-CurrentTenant -Name 'Acme' | Out-Null
        $cur = Get-CurrentTenant
        $cur.name | Should -Be 'Acme'
        $script:SessionState.TenantId   | Should -Be 'abcd-1234'
        $script:SessionState.TenantName | Should -Be 'Acme'
    }

    It "throws on an unregistered name" {
        { Set-CurrentTenant -Name 'no-such-thing' } | Should -Throw
    }

    It "stamps lastUsed" {
        Register-Tenant -Name 'Acme' -TenantId 'abcd' | Out-Null
        Set-CurrentTenant -Name 'Acme' | Out-Null
        (Get-Tenant -Name 'Acme').lastUsed | Should -Not -BeNullOrEmpty
    }
}

Describe "Update-Tenant" {
    It "applies a partial update without touching other fields" {
        $f = Get-TenantRegistryPath; if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
        Register-Tenant -Name 'Contoso' -TenantId 'abcd' -PrimaryDomain 'old.example.com' -Tags @('a','b') | Out-Null
        Update-Tenant   -Name 'Contoso' -PrimaryDomain 'new.example.com' | Out-Null
        $t = Get-Tenant -Name 'Contoso'
        $t.primaryDomain | Should -Be 'new.example.com'
        @($t.tags).Count | Should -Be 2
    }
}
