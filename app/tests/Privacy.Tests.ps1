# ============================================================
#  Pester tests for the privacy / PII redaction layer.
#  Run from the repo root: Invoke-Pester ./tests/
#
#  The tests dot-source UI.ps1, Auth.ps1, and AIAssistant.ps1 so the
#  helper functions are visible at script scope. None of the tests
#  hit a network endpoint or a real M365 tenant; everything operates
#  on string inputs.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')
}

Describe "Convert-ToSafePayload" {
    BeforeEach { Reset-PrivacyMap | Out-Null }

    It "tokenizes a UPN as <UPN_n>" {
        (Convert-ToSafePayload -Text 'user@example.com') | Should -Be '<UPN_1>'
    }

    It "tokenizes a GUID as <GUID_n>" {
        (Convert-ToSafePayload -Text '01234567-89ab-cdef-0123-456789abcdef') | Should -Be '<GUID_1>'
    }

    It "tokenizes a JWT as <JWT_n>" {
        $jwt = 'eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjMifQ.SflKxwRJSMeKKF2QT4f-abcDEFghij'
        (Convert-ToSafePayload -Text $jwt) | Should -Be '<JWT_1>'
    }

    It "tokenizes an Anthropic sk-ant- key as <SECRET_n>" {
        $k = 'sk-ant-api03-AbCdEfGhIjKlMnOpQrStUvWxYz0123456789'
        (Convert-ToSafePayload -Text $k) | Should -Be '<SECRET_1>'
    }

    It "tokenizes an OpenAI sk- key as <SECRET_n>" {
        $k = 'sk-1234567890abcdefghijklmnopqrstuvwx'
        (Convert-ToSafePayload -Text $k) | Should -Be '<SECRET_1>'
    }

    It "tokenizes a 40-hex cert thumbprint as <THUMB_n>" {
        (Convert-ToSafePayload -Text 'A1B2C3D4E5F6A7B8C9D0E1F2A3B4C5D6E7F8A9B0') | Should -Be '<THUMB_1>'
    }

    It "reuses the same token for repeated values (round-trip stability)" {
        $a = Convert-ToSafePayload -Text 'alice@example.com'
        $b = Convert-ToSafePayload -Text 'alice@example.com'
        $a | Should -Be $b
    }

    It "issues different tokens for different values" {
        $a = Convert-ToSafePayload -Text 'alice@example.com'
        $b = Convert-ToSafePayload -Text 'bob@example.com'
        $a | Should -Not -Be $b
    }

    It "deduplicates: same UPN reused in one payload shares one token" {
        $text   = 'alice@example.com sent to alice@example.com and bob@example.com'
        $result = Convert-ToSafePayload -Text $text
        $unique = [regex]::Matches($result, '<UPN_\d+>') | ForEach-Object { $_.Value } | Select-Object -Unique
        $unique.Count | Should -Be 2
    }

    It "in SecretsOnly mode leaves UPN and GUID raw but still scrubs JWT and sk- keys" {
        $text   = 'user@example.com guid 01234567-89ab-cdef-0123-456789abcdef token eyJabc.def.ghi key sk-ant-api03-AbCdEfGhIjKlMnOpQrStUvWxYz01234567'
        $result = Convert-ToSafePayload -Text $text -SecretsOnly $true
        $result | Should -Match 'user@example.com'
        $result | Should -Match '01234567-89ab-cdef-0123-456789abcdef'
        $result | Should -Match '<JWT_\d+>'
        $result | Should -Match '<SECRET_\d+>'
    }

    It "overlap: UPN whose local-part contains a GUID-shaped substring tokenizes whole, not split" {
        $upn    = 'user.01234567-89ab-cdef-0123-456789abcdef@contoso.com'
        $result = Convert-ToSafePayload -Text $upn
        $result | Should -Match '^<UPN_\d+>$'
        $result | Should -Not -Match '<GUID_'
    }

    It "tenant ID is tagged separately as <TENANT> (and reused for repeats)" {
        $script:SessionState.TenantId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        try {
            $text   = 'TenantId aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee and again aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
            $result = Convert-ToSafePayload -Text $text
            $result | Should -Match '<TENANT>'
            ([regex]::Matches($result, '<TENANT>')).Count | Should -Be 2
        } finally {
            $script:SessionState.TenantId = $null
        }
    }

    It "counts hashtable accumulates tokenization tallies" {
        $counts = @{}
        Convert-ToSafePayload -Text 'a@x.com b@y.com 01234567-89ab-cdef-0123-456789abcdef' -Counts $counts | Out-Null
        $counts['UPN']  | Should -Be 2
        $counts['GUID'] | Should -Be 1
    }

    It "tokenizes display name captured from Search displayName: idiom" {
        $result = Convert-ToSafePayload -Text 'Get-MgUser -Search "displayName:Jackie Smith"'
        $result | Should -Match '<NAME_\d+>'
        $result | Should -Not -Match 'Jackie Smith'
    }

    It "tokenizes display name captured from -DisplayName idiom" {
        $result = Convert-ToSafePayload -Text 'Set-MgUser -DisplayName "Bob Cole" -Department X'
        $result | Should -Match '<NAME_\d+>'
        $result | Should -Not -Match 'Bob Cole'
    }
}

Describe "Restore-FromSafePayload" {
    BeforeEach { Reset-PrivacyMap | Out-Null }

    It "round-trips a single-token payload back to the original string" {
        $orig = 'Get-MgUser -UserId alice@example.com'
        (Restore-FromSafePayload -Text (Convert-ToSafePayload -Text $orig)) | Should -Be $orig
    }

    It "round-trips a multi-token payload (UPN + GUID)" {
        $orig = 'User alice@example.com has ID 01234567-89ab-cdef-0123-456789abcdef'
        (Restore-FromSafePayload -Text (Convert-ToSafePayload -Text $orig)) | Should -Be $orig
    }

    It "leaves placeholders that were not minted in this session untouched" {
        (Restore-FromSafePayload -Text 'unknown placeholder <UPN_99> stays') | Should -Be 'unknown placeholder <UPN_99> stays'
    }

    It "is a no-op when the map is empty (no tokens minted)" {
        (Restore-FromSafePayload -Text 'plain text alice@x.com') | Should -Be 'plain text alice@x.com'
    }
}

Describe "Reset-PrivacyMap" {
    It "returns the count of dropped tokens and zeroes the maps" {
        Reset-PrivacyMap | Out-Null
        Convert-ToSafePayload -Text 'a@b.com c@d.com' | Out-Null
        $dropped = Reset-PrivacyMap
        $dropped              | Should -BeGreaterOrEqual 2
        $script:PrivacyMap.ByValue.Count | Should -Be 0
        $script:PrivacyMap.ByToken.Count | Should -Be 0
    }
}

Describe "Test-IsLocalEndpoint / Test-IsExternalProvider" {
    It "Ollama at 127.0.0.1 is local" {
        Test-IsExternalProvider -Provider 'Ollama' -Endpoint 'http://127.0.0.1:11434' -TrustedProviders @() | Should -Be $false
    }
    It "Ollama at localhost is local" {
        Test-IsExternalProvider -Provider 'Ollama' -Endpoint 'http://localhost:11434' -TrustedProviders @() | Should -Be $false
    }
    It "Ollama at a remote hostname is external" {
        Test-IsExternalProvider -Provider 'Ollama' -Endpoint 'http://gpu-server:11434' -TrustedProviders @() | Should -Be $true
    }
    It "Anthropic is always external (by default)" {
        Test-IsExternalProvider -Provider 'Anthropic' -Endpoint 'https://api.anthropic.com/v1/messages' -TrustedProviders @() | Should -Be $true
    }
    It "OpenAI is always external (by default)" {
        Test-IsExternalProvider -Provider 'OpenAI' -Endpoint 'https://api.openai.com/v1/chat/completions' -TrustedProviders @() | Should -Be $true
    }
    It "AzureOpenAI listed in TrustedProviders is treated as local" {
        Test-IsExternalProvider -Provider 'AzureOpenAI' -Endpoint 'https://x.openai.azure.com/' -TrustedProviders @('azure-openai') | Should -Be $false
    }
    It "Custom endpoint at localhost is local" {
        Test-IsExternalProvider -Provider 'Custom' -Endpoint 'http://localhost:9000/chat' -TrustedProviders @() | Should -Be $false
    }
    It "Custom endpoint at remote URL is external" {
        Test-IsExternalProvider -Provider 'Custom' -Endpoint 'https://example.com/chat' -TrustedProviders @() | Should -Be $true
    }
}

Describe "Format-DetailForAudit (RedactInAuditLog toggle)" {
    BeforeEach { Reset-PrivacyMap | Out-Null }

    It "Disabled: keeps UPN raw but still scrubs -Password value" {
        $cfg = @{ Privacy = @{ RedactInAuditLog = 'Disabled' } }
        $out = Format-DetailForAudit -Detail 'Set-MgUser -UserId alice@x.com -Password "abc123"' -Config $cfg
        $out | Should -Match 'alice@x.com'
        $out | Should -Match '\*\*\*REDACTED\*\*\*'
        $out | Should -Not -Match 'abc123'
    }

    It "Enabled: tokenizes UPN AND scrubs -Password value" {
        $cfg = @{ Privacy = @{ RedactInAuditLog = 'Enabled' } }
        $out = Format-DetailForAudit -Detail 'Set-MgUser -UserId alice@x.com -Password "abc123"' -Config $cfg
        $out | Should -Match '<UPN_\d+>'
        $out | Should -Not -Match 'alice@x.com'
        $out | Should -Match '\*\*\*REDACTED\*\*\*'
    }

    It "tolerates a null Config (no Privacy section)" {
        $out = Format-DetailForAudit -Detail 'Set-MgUser -Password "x"' -Config $null
        $out | Should -Match '\*\*\*REDACTED\*\*\*'
    }
}

Describe "Fixture: real-looking error trace emits zero recognizable PII" {
    BeforeEach { Reset-PrivacyMap | Out-Null }

    It "no email, no GUID, no JWT, no sk- key, no 40-hex thumbprint, no quoted display name" {
        $path = Join-Path $PSScriptRoot 'fixtures/sample-error-with-pii.txt'
        (Test-Path -LiteralPath $path) | Should -Be $true
        $script:SessionState.TenantId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        try {
            $text = Get-Content -LiteralPath $path -Raw
            $safe = Convert-ToSafePayload -Text $text

            $safe | Should -Not -Match '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}'
            $safe | Should -Not -Match '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b'
            $safe | Should -Not -Match 'eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+'
            $safe | Should -Not -Match 'sk-ant-[A-Za-z0-9_\-]{20,}'
            $safe | Should -Not -Match '(?<![\w-])sk-[A-Za-z0-9]{20,}'
            $safe | Should -Not -Match '\b[0-9A-Fa-f]{40}\b'
            $safe | Should -Not -Match 'Jackie Smith'

            # Round-trip recovers the original
            (Restore-FromSafePayload -Text $safe) | Should -Be $text
        } finally {
            $script:SessionState.TenantId = $null
        }
    }
}
