# ============================================================
#  Pester tests for Scheduler.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'AIAssistant.ps1')   # for Protect-ApiKey
    . (Join-Path $script:RepoRoot 'Scheduler.ps1')
}

Describe "ConvertTo-ScheduleSpec parsing" {
    It "parses Daily HH:MM" {
        $s = ConvertTo-ScheduleSpec 'Daily 09:00'
        $s.Frequency | Should -Be 'Daily'
        $s.Time      | Should -Be '09:00'
    }
    It "parses Weekly Day HH:MM" {
        $s = ConvertTo-ScheduleSpec 'Weekly Mon 09:00'
        $s.Frequency | Should -Be 'Weekly'
        $s.Day       | Should -Be 'Mon'
        $s.Time      | Should -Be '09:00'
    }
    It "parses Monthly day HH:MM" {
        $s = ConvertTo-ScheduleSpec 'Monthly 1 09:00'
        $s.Frequency | Should -Be 'Monthly'
        $s.Day       | Should -Be 1
    }
    It "parses Hourly" {
        $s = ConvertTo-ScheduleSpec 'Hourly'
        $s.Frequency | Should -Be 'Hourly'
    }
    It "parses raw cron" {
        $s = ConvertTo-ScheduleSpec 'cron 0 9 * * 1-5'
        $s.Frequency  | Should -Be 'Cron'
        $s.Expression | Should -Be '0 9 * * 1-5'
    }
    It "returns null on garbage" {
        (ConvertTo-ScheduleSpec 'Sometime soon') | Should -BeNullOrEmpty
    }
}

Describe "ConvertTo-CronExpression" {
    It "Daily 09:30 -> 30 9 * * *" {
        ConvertTo-CronExpression -Spec (ConvertTo-ScheduleSpec 'Daily 09:30') | Should -Be '30 9 * * *'
    }
    It "Weekly Wed 14:00 -> 0 14 * * 3" {
        ConvertTo-CronExpression -Spec (ConvertTo-ScheduleSpec 'Weekly Wed 14:00') | Should -Be '0 14 * * 3'
    }
    It "Monthly 15 06:00 -> 0 6 15 * *" {
        ConvertTo-CronExpression -Spec (ConvertTo-ScheduleSpec 'Monthly 15 06:00') | Should -Be '0 6 15 * *'
    }
    It "Hourly -> 0 * * * *" {
        ConvertTo-CronExpression -Spec (ConvertTo-ScheduleSpec 'Hourly') | Should -Be '0 * * * *'
    }
}

Describe "Non-interactive flag propagation" {
    BeforeAll { Set-NonInteractiveMode -Enabled $true }
    AfterAll  { Set-NonInteractiveMode -Enabled $false }

    It "Get-NonInteractiveMode reports true" { (Get-NonInteractiveMode) | Should -BeTrue }
    It "Read-UserInput returns empty string"   { (Read-UserInput 'anything') | Should -Be '' }
    It "Confirm-Action returns false (decline)" { (Confirm-Action 'destroy everything?') | Should -BeFalse }
    It "Show-Menu returns -1 (back)"            { (Show-Menu -Title 't' -Options @('a','b')) | Should -Be -1 }
}

Describe "Credential storage shape" {
    BeforeAll {
        # Redirect state path to a temp file
        $script:TempCred = Join-Path ([IO.Path]::GetTempPath()) ("sched-cred-test-" + [Guid]::NewGuid() + '.xml')
        Mock -CommandName Get-SchedulerCredentialPath -MockWith { $script:TempCred }
    }
    AfterAll { if (Test-Path $script:TempCred) { Remove-Item $script:TempCred -Force } }

    It "writes + reads + decrypts a stored secret" {
        # Direct write of a known shape so we can round-trip without UI prompting
        $enc = Protect-ApiKey -PlainKey 'super-secret-value'
        $cred = @{ tenantId='00000000-0000-0000-0000-000000000000'; appId='11111111-1111-1111-1111-111111111111'; encryptedSecret=$enc; registeredAt=(Get-Date).ToString('o') }
        ($cred | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $script:TempCred -Encoding UTF8 -Force
        $back = Get-SchedulerCredential
        $back.TenantId | Should -Be '00000000-0000-0000-0000-000000000000'
        $back.AppId    | Should -Be '11111111-1111-1111-1111-111111111111'
        $back.Secret   | Should -Be 'super-secret-value'
    }
}
