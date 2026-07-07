# ============================================================
#  Pester tests for GuestUsers.ps1
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'GuestUsers.ps1')
}

Describe "Stale-guest filter math" {
    It "DaysSinceSignIn >= threshold == stale" {
        $sample = @(
            [PSCustomObject]@{ UPN='active@x';    DaysSinceSignIn=5    },
            [PSCustomObject]@{ UPN='stale@x';     DaysSinceSignIn=120  },
            [PSCustomObject]@{ UPN='never@x';     DaysSinceSignIn=9999 }
        )
        $stale = @($sample | Where-Object DaysSinceSignIn -ge 90)
        $stale.Count            | Should -Be 2
        $stale[0].UPN           | Should -Be 'stale@x'
    }
    It "Treats 'no sign-in ever' (sentinel 9999) as stale at any threshold" {
        $never = [PSCustomObject]@{ UPN='ghost@x'; DaysSinceSignIn=9999 }
        ($never.DaysSinceSignIn -ge 30) | Should -BeTrue
    }
}

Describe "Recert state transitions" {
    BeforeAll {
        # Redirect the state-path helper to a temp file so we don't
        # touch the real machine state.
        $script:TempState = Join-Path ([IO.Path]::GetTempPath()) ("recert-test-" + [Guid]::NewGuid() + '.json')
        Mock -CommandName Get-RecertStatePath -MockWith { $script:TempState }
    }
    AfterAll { if (Test-Path $script:TempState) { Remove-Item $script:TempState -Force } }

    It "round-trips an empty array" {
        Write-RecertState -Records @()
        (Read-RecertState).Count | Should -Be 0
    }
    It "queues a pending record and reads it back" {
        $rec = [PSCustomObject]@{ campaignId='c1'; guestId='g1'; guestUpn='guest@x'; managerUpn='mgr@x'; queuedAt=(Get-Date).ToString('o'); state='pending'; decisionBy=$null; decisionAt=$null; notes='' }
        Write-RecertState -Records @($rec)
        $read = Read-RecertState
        $read.Count           | Should -Be 1
        $read[0].state        | Should -Be 'pending'
        $read[0].guestUpn     | Should -Be 'guest@x'
    }
    It "applies a Keep decision" {
        $rec = (Read-RecertState)[0]
        $rec.state = 'keep'; $rec.decisionAt = (Get-Date).ToString('o'); $rec.decisionBy = 'admin@x'
        Write-RecertState -Records @($rec)
        $read = Read-RecertState
        $read[0].state      | Should -Be 'keep'
        $read[0].decisionBy | Should -Be 'admin@x'
    }
}

Describe "Domain pivot logic" {
    It "buckets guests by domain extracted from mail" {
        $sample = @(
            [PSCustomObject]@{ UPN='a@vendor.com_#EXT#@contoso.onmicrosoft.com'; Mail='a@vendor.com';   Domains='vendor.com' },
            [PSCustomObject]@{ UPN='b@vendor.com_#EXT#@contoso.onmicrosoft.com'; Mail='b@vendor.com';   Domains='vendor.com' },
            [PSCustomObject]@{ UPN='c@partner.io_#EXT#@contoso.onmicrosoft.com'; Mail='c@partner.io';   Domains='partner.io' }
        )
        $bucket = @{}
        foreach ($g in $sample) {
            $d = $g.Domains
            if (-not $bucket.ContainsKey($d)) { $bucket[$d] = 0 }
            $bucket[$d]++
        }
        $bucket['vendor.com']  | Should -Be 2
        $bucket['partner.io']  | Should -Be 1
    }
}

Describe "Bulk guest removal CSV validation" {
    It "keeps a valid row and applies the default reason when none is given" {
        $v = Test-BulkGuestRemovalCsv -Rows @([PSCustomObject]@{ UPN = 'guest_x#EXT#@contoso.onmicrosoft.com' })
        $v.Rows.Count     | Should -Be 1
        $v.Errors.Count   | Should -Be 0
        $v.Rows[0].Reason | Should -Be 'Bulk removal'
    }
    It "accepts a UserPrincipalName column as an alias for UPN" {
        $v = Test-BulkGuestRemovalCsv -Rows @([PSCustomObject]@{ UserPrincipalName = 'a@b.com'; Reason = 'x' })
        $v.Rows.Count     | Should -Be 1
        $v.Rows[0].UPN    | Should -Be 'a@b.com'
        $v.Rows[0].Reason | Should -Be 'x'
    }
    It "accepts a bare object GUID as an identifier" {
        $v = Test-BulkGuestRemovalCsv -Rows @([PSCustomObject]@{ UPN = '11111111-2222-3333-4444-555555555555' })
        $v.Errors.Count | Should -Be 0
        $v.Rows.Count   | Should -Be 1
    }
    It "flags a missing UPN, non-identifier junk, and a duplicate" {
        $rows = @(
            [PSCustomObject]@{ UPN = 'guest_x#EXT#@contoso.onmicrosoft.com'; Reason = 'ok' },
            [PSCustomObject]@{ UPN = '';          Reason = 'blank' },
            [PSCustomObject]@{ UPN = 'not-an-id'; Reason = 'junk' },
            [PSCustomObject]@{ UPN = 'guest_x#EXT#@contoso.onmicrosoft.com'; Reason = 'dup' }
        )
        $v = Test-BulkGuestRemovalCsv -Rows $rows
        $v.Rows.Count   | Should -Be 1
        $v.Errors.Count | Should -Be 3
    }
}

Describe "Remove-Guest safety guard" {
    It "refuses to delete a Member account (returns NotAGuest)" {
        function Connect-ForTask { param($Task) $true }
        function Write-AuditEntry { param($EventType,$Detail,$ActionType,$Target,$Result,$ErrorMessage,$Reverse,$NoUndoReason,$EntryId) 'stub' }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            @{ id = 'm-1'; userType = 'Member'; userPrincipalName = 'real.employee@contoso.com' } }
        Set-PreviewMode -Enabled $false
        $r = Remove-Guest -UPN 'real.employee@contoso.com' -Reason 'fat-fingered CSV'
        $r.Status | Should -Be 'NotAGuest'
    }
    It "reports NotFound when the user does not resolve" {
        function Connect-ForTask { param($Task) $true }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction) throw 'Request_ResourceNotFound' }
        Set-PreviewMode -Enabled $false
        $r = Remove-Guest -UPN 'ghost@contoso.com' -Reason 'x'
        $r.Status | Should -Be 'NotFound'
    }
    It "removes a genuine Guest and reports Removed" {
        function Connect-ForTask { param($Task) $true }
        function Write-AuditEntry { param($EventType,$Detail,$ActionType,$Target,$Result,$ErrorMessage,$Reverse,$NoUndoReason,$EntryId) 'stub' }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            if ($Uri -match '/memberOf') { return @{ value = @() } }
            if ($Method -eq 'DELETE')    { return $null }
            return @{ id = 'g-1'; userType = 'Guest'; userPrincipalName = 'ext_x#EXT#@contoso.onmicrosoft.com' } }
        Set-PreviewMode -Enabled $false
        $r = Remove-Guest -UPN 'ext_x#EXT#@contoso.onmicrosoft.com' -Reason 'ended'
        $r.Status | Should -Be 'Removed'
    }
    It "honours -AllowNonGuest to delete a non-guest when explicitly forced" {
        function Connect-ForTask { param($Task) $true }
        function Write-AuditEntry { param($EventType,$Detail,$ActionType,$Target,$Result,$ErrorMessage,$Reverse,$NoUndoReason,$EntryId) 'stub' }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            if ($Uri -match '/memberOf') { return @{ value = @() } }
            if ($Method -eq 'DELETE')    { return $null }
            return @{ id = 'm-1'; userType = 'Member'; userPrincipalName = 'contractor@contoso.com' } }
        Set-PreviewMode -Enabled $false
        $r = Remove-Guest -UPN 'contractor@contoso.com' -Reason 'forced' -AllowNonGuest
        $r.Status | Should -Be 'Removed'
    }
}

Describe "Remove-Guest bulk fast path" {
    It "with -SkipConnect it never connects and skips ancillary teardown" {
        $script:connCalls = 0
        $script:spoCalls  = 0
        function Connect-ForTask { param($Task) $script:connCalls++; $true }
        function Get-UserOutboundShares { @() }
        function Invoke-SharePointOffboardCleanup { param($LeaverUPN,$LookbackDays) $script:spoCalls++ }
        function Write-AuditEntry { param($EventType,$Detail,$ActionType,$Target,$Result,$ErrorMessage,$Reverse,$NoUndoReason,$EntryId) 'stub' }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            if ($Uri -match '/memberOf') { return @{ value = @() } }
            if ($Method -eq 'DELETE')    { return $null }
            return @{ id = 'g-1'; userType = 'Guest'; userPrincipalName = 'ext_x#EXT#@contoso.onmicrosoft.com' } }
        Set-PreviewMode -Enabled $false
        $r = Remove-Guest -UPN 'ext_x#EXT#@contoso.onmicrosoft.com' -Reason 'bulk' -SkipConnect
        $r.Status         | Should -Be 'Removed'
        $script:connCalls | Should -Be 0
        $script:spoCalls  | Should -Be 0
    }
}

Describe "Graph throttling retry" {
    It "retries after a 429 and then succeeds" {
        $script:attempts = 0
        function Start-Sleep { param($Seconds,$Milliseconds) }   # no real waiting in tests
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            $script:attempts++
            if ($script:attempts -eq 1) { throw 'Response status code does not indicate success: 429 (Too Many Requests).' }
            return @{ ok = $true } }
        $out = Invoke-GuestGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/users/x'
        $script:attempts | Should -Be 2
        $out.ok          | Should -BeTrue
    }
    It "re-throws a non-throttle error without retrying" {
        $script:attempts2 = 0
        function Start-Sleep { param($Seconds,$Milliseconds) }
        function Invoke-MgGraphRequest { param($Method,$Uri,$Headers,$Body,$ContentType,$ErrorAction)
            $script:attempts2++; throw 'Request_ResourceNotFound: does not exist' }
        { Invoke-GuestGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/users/x' } | Should -Throw
        $script:attempts2 | Should -Be 1
    }
}
