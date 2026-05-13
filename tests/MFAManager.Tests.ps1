# ============================================================
#  Pester tests for MFAManager.ps1
#
#  Tests cover method-type classification and compliance-view
#  filter logic. No live Graph calls -- Get-UserAuthMethods is
#  mocked with a fixed dataset.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'MFAManager.ps1')
}

Describe "Get-MfaMethodInfo (classification)" {
    It "classifies Microsoft Authenticator" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'; id='m1'; displayName='Pixel 8' }
        $info = Get-MfaMethodInfo -Method $m
        $info.Label      | Should -Be 'Microsoft Authenticator'
        $info.UrlSegment | Should -Be 'microsoftAuthenticatorMethods'
        $info.Id         | Should -Be 'm1'
    }
    It "classifies Phone" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.phoneAuthenticationMethod'; id='p1'; phoneNumber='+15551234567'; phoneType='mobile' }
        $info = Get-MfaMethodInfo -Method $m
        $info.Label      | Should -Be 'Phone (SMS/voice)'
        $info.UrlSegment | Should -Be 'phoneMethods'
    }
    It "classifies FIDO2" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.fido2AuthenticationMethod'; id='f1'; model='YubiKey 5C' }
        $info = Get-MfaMethodInfo -Method $m
        $info.Label      | Should -Be 'FIDO2 key'
        $info.UrlSegment | Should -Be 'fido2Methods'
    }
    It "classifies TAP" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.temporaryAccessPassAuthenticationMethod'; id='t1'; lifetimeInMinutes=60; isUsableOnce=$true }
        $info = Get-MfaMethodInfo -Method $m
        $info.Label      | Should -Be 'Temporary Access Pass'
        $info.UrlSegment | Should -Be 'temporaryAccessPassMethods'
    }
    It "Password method has no UrlSegment (cannot be deleted)" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.passwordAuthenticationMethod'; id='pw' }
        $info = Get-MfaMethodInfo -Method $m
        $info.Label      | Should -Be 'Password'
        $info.UrlSegment | Should -Be ''
    }
    It "Unknown odata type still produces a row (graceful fallback)" {
        $m = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.somethingNewAuthenticationMethod'; id='x' }
        $info = Get-MfaMethodInfo -Method $m
        $info | Should -Not -BeNullOrEmpty
        $info.UrlSegment | Should -Be ''
    }
}

Describe "Compliance-view predicate logic" {
    BeforeAll {
        # Build canned method sets:
        # - alice has FIDO2 + Authenticator (strong)
        # - bob has only phone
        # - carol has nothing real (just Password)
        # - dave has a TAP
        $script:Methods_alice = @(
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.fido2AuthenticationMethod'; id='f1' })),
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'; id='m1' })),
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.passwordAuthenticationMethod'; id='pw' }))
        )
        $script:Methods_bob = @(
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.phoneAuthenticationMethod'; id='p1'; phoneNumber='+1...'; phoneType='mobile' })),
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.passwordAuthenticationMethod'; id='pw' }))
        )
        $script:Methods_carol = @(
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.passwordAuthenticationMethod'; id='pw' }))
        )
        $script:Methods_dave = @(
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.temporaryAccessPassAuthenticationMethod'; id='t1'; lifetimeInMinutes=60 })),
            (Get-MfaMethodInfo -Method ([PSCustomObject]@{ '@odata.type'='#microsoft.graph.passwordAuthenticationMethod'; id='pw' }))
        )
    }
    It "OnlyPhone predicate matches bob, not alice/carol/dave" {
        $p = { param($methods) $real = @($methods | Where-Object { $_.Label -ne 'Password' }); if ($real.Count -eq 0) { return $false }; ($real | Where-Object { $_.Label -ne 'Phone (SMS/voice)' }).Count -eq 0 }
        (& $p $script:Methods_alice) | Should -BeFalse
        (& $p $script:Methods_bob)   | Should -BeTrue
        (& $p $script:Methods_carol) | Should -BeFalse
        (& $p $script:Methods_dave)  | Should -BeFalse
    }
    It "NoMfa predicate matches carol, not the others" {
        $p = { param($methods) $real = @($methods | Where-Object { $_.Label -ne 'Password' }); $real.Count -eq 0 }
        (& $p $script:Methods_alice) | Should -BeFalse
        (& $p $script:Methods_bob)   | Should -BeFalse
        (& $p $script:Methods_carol) | Should -BeTrue
        (& $p $script:Methods_dave)  | Should -BeFalse
    }
    It "ActiveTAP predicate matches dave" {
        $p = { param($methods) ($methods | Where-Object { $_.Label -eq 'Temporary Access Pass' }).Count -gt 0 }
        (& $p $script:Methods_alice) | Should -BeFalse
        (& $p $script:Methods_bob)   | Should -BeFalse
        (& $p $script:Methods_carol) | Should -BeFalse
        (& $p $script:Methods_dave)  | Should -BeTrue
    }
    It "FIDO2 predicate matches alice" {
        $p = { param($methods) ($methods | Where-Object { $_.Label -eq 'FIDO2 key' }).Count -gt 0 }
        (& $p $script:Methods_alice) | Should -BeTrue
        (& $p $script:Methods_bob)   | Should -BeFalse
    }
}

Describe "Format-MfaMethodDetail" {
    It "formats phone with number + type" {
        $m = [PSCustomObject]@{ '@odata.type'='#microsoft.graph.phoneAuthenticationMethod'; phoneNumber='+15551234567'; phoneType='mobile' }
        Format-MfaMethodDetail -Method $m | Should -Be '+15551234567 (mobile)'
    }
    It "formats FIDO2 with model + aaguid" {
        $m = [PSCustomObject]@{ '@odata.type'='#microsoft.graph.fido2AuthenticationMethod'; model='YubiKey 5C'; aaGuid='guid-1' }
        Format-MfaMethodDetail -Method $m | Should -Match 'YubiKey 5C'
        Format-MfaMethodDetail -Method $m | Should -Match 'guid-1'
    }
}
