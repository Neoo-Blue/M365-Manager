# ============================================================
#  Pester tests for IncidentTriggers.ps1
#
#  Exercises each detector's predicate against canned data.
#  No live Graph / UAL / SPO calls -- detectors that need
#  Search-SignIns / Search-UAL / Get-UserOutboundShares /
#  Invoke-MgGraphRequest are tested by passing pre-fetched
#  data directly via the -SignIns parameter where supported,
#  or by skipping when the helper isn't available.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'IncidentResponse.ps1')
    . (Join-Path $script:RepoRoot 'IncidentTriggers.ps1')
}

Describe "Get-IncidentTriggerConfig defaults" {
    It "supplies sensible defaults when no config is set" {
        $cfg = Get-IncidentTriggerConfig
        $cfg.AutoExecuteOnSeverity         | Should -Be 'None'
        $cfg.UseAIForNarrative             | Should -Be 'Disabled'
        $cfg.MassDownloadFileCount         | Should -Be 50
        $cfg.MassDownloadWindowMinutes     | Should -Be 5
        $cfg.MassShareCount                | Should -Be 20
        $cfg.MFAFatigueRejectCount         | Should -Be 10
        $cfg.ImpossibleTravelMaxKmPerHour  | Should -Be 900
        $cfg.AnomalousLocationLookbackDays | Should -Be 90
    }
}

Describe "Get-CountryDistanceKm" {
    It "returns 0 for identical countries" {
        (Get-CountryDistanceKm -Country1 'US' -Country2 'US') | Should -Be 0
    }
    It "returns the haversine distance for known anchors" {
        # Lagos vs Sao Paulo is ~7300km centroid-to-centroid; allow generous tolerance
        $d = Get-CountryDistanceKm -Country1 'Nigeria' -Country2 'Brazil'
        $d | Should -BeGreaterThan 5000
        $d | Should -BeLessThan 9000
    }
    It "returns null when an anchor is missing" {
        (Get-CountryDistanceKm -Country1 'Atlantis' -Country2 'United States') | Should -BeNullOrEmpty
    }
}

Describe "New-IncidentFinding" {
    It "shapes findings with the documented fields" {
        $f = New-IncidentFinding -TriggerType 'TestTrigger' -UPN 'u@x.com' -Severity 'High' -Evidence @{ key='val' } -RecommendedAction 'do thing'
        $f.TriggerType       | Should -Be 'TestTrigger'
        $f.UPN               | Should -Be 'u@x.com'
        $f.Severity          | Should -Be 'High'
        $f.Evidence.key      | Should -Be 'val'
        $f.RecommendedAction | Should -Be 'do thing'
        $f.DetectedUtc       | Should -Not -BeNullOrEmpty
    }
}

Describe "Detect-AnomalousLocationSignIn" {
    It "returns null when there is only one sign-in" {
        $signIns = @(@{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'US' })
        (Detect-AnomalousLocationSignIn -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
    It "fires when the latest country is not in the baseline set" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-10T10:00:00Z'; Location = 'United States' },
            @{ CreatedDateTime = '2026-05-11T10:00:00Z'; Location = 'United States' },
            @{ CreatedDateTime = '2026-05-14T17:00:00Z'; Location = 'Nigeria' }
        )
        $f = Detect-AnomalousLocationSignIn -UPN 'u@x.com' -SignIns $signIns
        $f                                    | Should -Not -BeNullOrEmpty
        $f.TriggerType                        | Should -Be 'AnomalousLocationSignIn'
        $f.Severity                           | Should -Be 'Low'
        $f.Evidence.latestCountry             | Should -Be 'Nigeria'
        ($f.Evidence.baselineCountries -join ',') | Should -Be 'United States'
    }
    It "does NOT fire when the latest country is already in baseline" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-10T10:00:00Z'; Location = 'United States' },
            @{ CreatedDateTime = '2026-05-11T10:00:00Z'; Location = 'Germany' },
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'United States' }
        )
        (Detect-AnomalousLocationSignIn -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
}

Describe "Detect-ImpossibleTravel" {
    It "fires on Nigeria -> Brazil in 1 hour (~7000 km, 7000 km/h implied)" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'Nigeria' },
            @{ CreatedDateTime = '2026-05-14T11:00:00Z'; Location = 'Brazil' }
        )
        $f = Detect-ImpossibleTravel -UPN 'u@x.com' -SignIns $signIns
        $f | Should -Not -BeNullOrEmpty
        $f.TriggerType         | Should -Be 'ImpossibleTravel'
        $f.Severity            | Should -Be 'High'
        $f.Evidence.fromCountry| Should -Be 'Nigeria'
        $f.Evidence.toCountry  | Should -Be 'Brazil'
    }
    It "does NOT fire on US -> Canada in 4h (~2000 km, 500 km/h ok)" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'United States' },
            @{ CreatedDateTime = '2026-05-14T14:00:00Z'; Location = 'Canada' }
        )
        (Detect-ImpossibleTravel -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
    It "does NOT fire when both sign-ins are in the same country" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'United States' },
            @{ CreatedDateTime = '2026-05-14T10:05:00Z'; Location = 'United States' }
        )
        (Detect-ImpossibleTravel -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
}

Describe "Detect-HighRiskSignIn" {
    It "fires on a risk='high' sign-in" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T17:00:00Z'; IpAddress = '203.0.113.42'; Location = 'Nigeria'; RiskLevel = 'high' }
        )
        $f = Detect-HighRiskSignIn -UPN 'u@x.com' -SignIns $signIns
        $f.TriggerType   | Should -Be 'HighRiskSignIn'
        $f.Severity      | Should -Be 'High'
        $f.Evidence.riskLevel | Should -Be 'high'
    }
    It "does NOT fire on risk='none' sign-ins only" {
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T17:00:00Z'; RiskLevel = 'none' },
            @{ CreatedDateTime = '2026-05-14T16:00:00Z'; RiskLevel = 'low' }
        )
        (Detect-HighRiskSignIn -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
}

Describe "Tuning via config overrides" {
    BeforeEach {
        # Stub Get-EffectiveConfig to return a custom threshold
        function global:Get-EffectiveConfig {
            param([string]$Key)
            switch ($Key) {
                'IncidentResponse.ImpossibleTravelMaxKmPerHour' { return 3000 }
                default { return $null }
            }
        }
    }
    AfterEach {
        Remove-Item Function:Get-EffectiveConfig -ErrorAction SilentlyContinue
    }
    It "honors a raised ImpossibleTravelMaxKmPerHour threshold (Nigeria->Brazil in 3h now OK)" {
        # ~7300 km / 3h = 2433 km/h. Below 3000 km/h threshold.
        $signIns = @(
            @{ CreatedDateTime = '2026-05-14T10:00:00Z'; Location = 'Nigeria' },
            @{ CreatedDateTime = '2026-05-14T13:00:00Z'; Location = 'Brazil' }
        )
        (Detect-ImpossibleTravel -UPN 'u@x.com' -SignIns $signIns) | Should -BeNullOrEmpty
    }
}
