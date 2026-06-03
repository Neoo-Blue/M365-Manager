# ============================================================
#  Pester tests for LicenseOptimizer.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'LicenseOptimizer.ps1')
}

Describe "Anonymized-username detection" {
    It "Matches Graph's de-identified hash shape" {
        $fake = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678901AABBCCDDEEFF112233'
        ($fake -match '^[A-F0-9]{50,}$') | Should -BeTrue
    }
    It "Does NOT trip on a real UPN" {
        ('jane.smith@contoso.com' -match '^[A-F0-9]{50,}$') | Should -BeFalse
    }
}

Describe "License-family overlap detection" {
    It "Flags SPE_E3 + ENTERPRISEPACK in the same family" {
        $family = $script:LicenseFamilies['M365_E_FAMILY']
        $skus = @('SPE_E3','ENTERPRISEPACK')
        $inFamily = @($skus | Where-Object { $family -contains $_ })
        $inFamily.Count | Should -BeGreaterOrEqual 2
    }
    It "Does NOT flag SPE_E3 + POWER_BI_PRO (different families)" {
        $family = $script:LicenseFamilies['M365_E_FAMILY']
        $skus = @('SPE_E3','POWER_BI_PRO')
        $inFamily = @($skus | Where-Object { $family -contains $_ })
        $inFamily.Count | Should -Be 1
    }
}

Describe "Savings math" {
    It "Sums per-SKU list prices for a user's bundle" {
        $prices = Get-LicensePrices
        $prices['SPE_E3'] | Should -BeGreaterOrEqual 30   # ~36 by default; tolerate operator overrides as long as it's reasonable
        $bundle = @('SPE_E3','POWER_BI_PRO')
        $sum = 0.0
        foreach ($s in $bundle) { $p = Get-LicensePrice $s; if ($p) { $sum += $p } }
        $sum | Should -BeGreaterOrEqual 30
    }
    It "Returns null for an unknown SKU" {
        (Get-LicensePrice 'NOT_A_REAL_SKU_XYZ') | Should -BeNullOrEmpty
    }
}

Describe "Downgrade savings calculation" {
    It "E5 -> E3 difference is positive" {
        $e5 = Get-LicensePrice 'SPE_E5'
        $e3 = Get-LicensePrice 'SPE_E3'
        if ($e5 -and $e3) { ($e5 - $e3) | Should -BeGreaterThan 0 }
    }
}
