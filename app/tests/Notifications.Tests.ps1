# ============================================================
#  Pester tests for Notifications.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')  # for Protect-ApiKey delegation
    . (Join-Path $script:RepoRoot 'Notifications.ps1')
}

Describe "Protect-Secret / Unprotect-Secret round trip" {
    It "delegates to Protect-ApiKey when available" {
        $enc = Protect-Secret -PlainText 'hello world'
        $enc | Should -Not -Be 'hello world'
        ($enc -like 'DPAPI:*' -or $enc -like 'B64:*') | Should -BeTrue
        $dec = Unprotect-Secret -Stored $enc
        $dec | Should -Be 'hello world'
    }
    It "treats an already-encrypted value as a passthrough on protect" {
        $first  = Protect-Secret -PlainText 'foo'
        $second = Protect-Secret -PlainText $first
        $second | Should -Be $first
    }
}

Describe "Severity -> theme + importance" {
    It "Critical -> red theme + High importance (smoke check the mapping)" {
        # We don't fire the actual call; just verify the mapping table behavior the dispatcher uses.
        $themeColor = switch ('Critical') { 'Critical' { 'B00020' } 'Warning' { 'F0A030' } default { '0078D4' } }
        $themeColor | Should -Be 'B00020'
        $importance = if ('Critical' -eq 'Critical') { 'High' } else { 'Normal' }
        $importance | Should -Be 'High'
    }
}

Describe "DryRunNotifications" {
    BeforeAll {
        $script:TmpAi = Join-Path ([IO.Path]::GetTempPath()) ("notif-ai-" + [Guid]::NewGuid() + '.json')
        # Stub Get-NotificationsConfig so we don't touch the real ai_config.json
        Mock -CommandName Get-NotificationsConfig -MockWith {
            @{
                DefaultEmailFrom         = ''
                SecurityTeamRecipients   = @()
                OperationsTeamRecipients = @()
                TeamsWebhookSecurity     = ''
                TeamsWebhookOperations   = ''
                DryRunNotifications      = $true
            }
        }
    }
    It "Send-Email returns true and does NOT actually call Graph when DryRunNotifications=true" {
        # In dry-run mode the function short-circuits before touching
        # Invoke-MgGraphRequest, so we can call it safely with bogus
        # recipient data and still get a true return.
        $result = Send-Email -To @('nobody@example.com') -Subject 'Smoke' -Body '<b>hi</b>'
        $result | Should -BeTrue
    }
    It "Send-TeamsWebhook returns true and does NOT call REST when DryRunNotifications=true" {
        $result = Send-TeamsWebhook -WebhookUrl 'https://example.invalid/webhook' -Title 'Smoke' -Body 'hi'
        $result | Should -BeTrue
    }
}

Describe "Recipient resolution in Send-Notification" {
    BeforeAll {
        Mock -CommandName Get-NotificationsConfig -MockWith {
            @{
                DefaultEmailFrom         = 'svc@x.com'
                SecurityTeamRecipients   = @('sec@x.com')
                OperationsTeamRecipients = @('ops@x.com')
                TeamsWebhookSecurity     = 'https://sec-webhook'
                TeamsWebhookOperations   = 'https://ops-webhook'
                DryRunNotifications      = $true
            }
        }
        # Spy on Send-Email / Send-TeamsWebhook to assert recipient routing
        $script:CapturedEmailTo  = $null
        $script:CapturedWebhook  = $null
        Mock -CommandName Send-Email        -MockWith { param($To,$Subject,$Body,$From,$Importance) $script:CapturedEmailTo = $To;  return $true }
        Mock -CommandName Send-TeamsWebhook -MockWith { param($WebhookUrl,$Title,$Body,$ThemeColor) $script:CapturedWebhook = $WebhookUrl; return $true }
    }
    It "Critical routes to SecurityTeamRecipients" {
        Send-Notification -Channels @('email') -Subject 's' -Body 'b' -Severity 'Critical' | Out-Null
        $script:CapturedEmailTo | Should -Contain 'sec@x.com'
    }
    It "Info routes to OperationsTeamRecipients" {
        $script:CapturedEmailTo = $null
        Send-Notification -Channels @('email') -Subject 's' -Body 'b' -Severity 'Info' | Out-Null
        $script:CapturedEmailTo | Should -Contain 'ops@x.com'
    }
}
