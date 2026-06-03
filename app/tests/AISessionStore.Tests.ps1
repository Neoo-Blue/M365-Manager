# ============================================================
#  Pester tests for AISessionStore.ps1
#
#  Save / load round-trip, privacy-map preservation, redacted
#  export, index integrity. Encryption is exercised via DPAPI
#  on Windows and falls back to plaintext on POSIX -- both
#  paths satisfy the round-trip assertions.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')   # for Convert-ToSafePayload + PrivacyMap
    . (Join-Path $script:RepoRoot 'AISessionStore.ps1')

    $script:TempStateRoot = Join-Path ([IO.Path]::GetTempPath()) ("session-tests-" + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $script:TempStateRoot -ItemType Directory -Force | Out-Null
    $env:LOCALAPPDATA_BACKUP = $env:LOCALAPPDATA
    $env:LOCALAPPDATA        = $script:TempStateRoot
}

AfterAll {
    if ($env:LOCALAPPDATA_BACKUP) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_BACKUP }
    Remove-Item -LiteralPath $script:TempStateRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe "Save-AISession / Load-AISession round-trip" {
    BeforeEach {
        Reset-PrivacyMap | Out-Null
        Set-AISessionEphemeral -On $false
    }

    It "persists and reloads a one-message chat" {
        $cfg = @{ Provider='Anthropic'; Model='claude-haiku-4-5' }
        $hist = @(@{ role='user'; content='hi' })
        $id = Save-AISession -Config $cfg -History $hist -Title 'hello-test' -Force
        $id | Should -Not -BeNullOrEmpty

        $loaded = Load-AISession -IdOrPrefix $id
        $loaded                 | Should -Not -BeNullOrEmpty
        $loaded.Title           | Should -Be 'hello-test'
        $loaded.History.Count   | Should -Be 1
        $loaded.History[0].role | Should -Be 'user'
    }

    It "resolves by title prefix when id is unknown" {
        $cfg = @{ Provider='Anthropic'; Model='claude-haiku-4-5' }
        Save-AISession -Config $cfg -History @(@{role='user'; content='x'}) -Title 'budget-investigation-march' -Force | Out-Null
        $loaded = Load-AISession -IdOrPrefix 'budget-invest'
        $loaded | Should -Not -BeNullOrEmpty
        $loaded.Title | Should -Be 'budget-investigation-march'
    }

    It "preserves the privacy map across save/load" {
        Reset-PrivacyMap | Out-Null
        Convert-ToSafePayload -Text 'alice@contoso.com' | Out-Null   # registers <UPN_1>
        $cfg = @{ Provider='Anthropic'; Model='claude-haiku-4-5' }
        $id = Save-AISession -Config $cfg -History @(@{role='user'; content='hi'}) -Title 'privacy-test' -Force
        Reset-PrivacyMap | Out-Null
        Load-AISession -IdOrPrefix $id | Out-Null
        # After load the privacy map should still resolve <UPN_1>.
        $script:PrivacyMap.ByValue.ContainsKey('alice@contoso.com') | Should -BeTrue
    }
}

Describe "Show-AISessionList" {
    It "doesn't throw on an empty index" {
        # Wipe any prior entries
        $idxPath = Get-AISessionIndexPath
        if (Test-Path $idxPath) { Set-Content -LiteralPath $idxPath -Value '[]' -Force }
        { Show-AISessionList } | Should -Not -Throw
    }
}

Describe "Remove-AISession" {
    It "returns false on missing id" {
        (Remove-AISession -IdOrPrefix 'definitely-not-real') | Should -BeFalse
    }
}

Describe "Set-AISessionEphemeral" {
    It "suppresses auto-save unless -Force is passed" {
        Set-AISessionEphemeral -On $true
        $cfg = @{ Provider='Anthropic'; Model='claude-haiku-4-5' }
        $id = Save-AISession -Config $cfg -History @(@{role='user'; content='hi'})
        $id | Should -BeNullOrEmpty
        Set-AISessionEphemeral -On $false
    }
}

Describe "Export-AISession (redacted)" {
    It "writes a file with redacted UPNs (no real values)" {
        Reset-PrivacyMap | Out-Null
        Set-AISessionEphemeral -On $false
        $cfg = @{ Provider='Anthropic'; Model='claude-haiku-4-5' }
        $id = Save-AISession -Config $cfg -History @(@{role='user'; content='hello alice@contoso.com'}) -Title 'export-test' -Force
        $dest = Join-Path $script:TempStateRoot 'export.json'
        Export-AISession -IdOrPrefix $id -DestinationPath $dest | Out-Null
        Test-Path $dest | Should -BeTrue
        $content = Get-Content -LiteralPath $dest -Raw
        $content | Should -Match '<UPN_'
        $content | Should -Not -Match 'alice@contoso.com'
    }
}
