# ============================================================
#  GuestUsers.ps1 — guest lifecycle
#
#  Discovery via Graph /users?$filter=userType eq 'Guest', joined
#  with /auditLogs/signIns activity. Recertification campaigns are
#  managed via a per-tenant JSON state file at
#  <stateDir>/guest-recerts.json so manager replies (collected
#  outside the tool, by email) can be applied later via
#  Show-PendingRecerts.
# ============================================================

# ============================================================
#  Discovery
# ============================================================

function Get-Guests {
    <#
        Returns Guest users with enriched activity timestamps.
          -Domain   : suffix (e.g. "contoso.com") -- match on mail
                      or otherMails to find domain-tied guests
          -InvitedBy: UPN -- requires Get-MgUser /externalUserState
                      details which Graph doesn't always expose
                      cleanly; we use signInActivity proxies
          -MinAgeDays
          -MaxLastSignInDays
    #>
    param(
        [string]$Domain,
        [string]$InvitedBy,
        [int]$MinAgeDays,
        [int]$MaxLastSignInDays
    )
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=id,userPrincipalName,displayName,mail,otherMails,createdDateTime,accountEnabled,signInActivity,externalUserState,externalUserStateChangeDateTime&`$top=500&`$count=true"
    $headers = @{ ConsistencyLevel = 'eventual' }
    $out = New-Object System.Collections.ArrayList
    try {
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers -ErrorAction Stop
            foreach ($u in $resp.value) {
                $created     = if ($u.createdDateTime) { [DateTime]$u.createdDateTime } else { $null }
                $lastSignIn  = if ($u.signInActivity.lastSignInDateTime) { [DateTime]$u.signInActivity.lastSignInDateTime } else { $null }
                $allEmails   = @($u.mail) + @($u.otherMails)
                $domains     = @($allEmails | Where-Object { $_ -and $_ -match '@' } | ForEach-Object { ($_ -split '@')[1].ToLowerInvariant() } | Sort-Object -Unique)
                [void]$out.Add([PSCustomObject]@{
                    Id                  = [string]$u.id
                    UPN                 = [string]$u.userPrincipalName
                    DisplayName         = [string]$u.displayName
                    Mail                = [string]$u.mail
                    Domains             = ($domains -join ', ')
                    CreatedUtc          = $created
                    AgeDays             = if ($created) { [Math]::Round(((Get-Date).ToUniversalTime() - $created.ToUniversalTime()).TotalDays, 0) } else { $null }
                    LastSignInUtc       = $lastSignIn
                    DaysSinceSignIn     = if ($lastSignIn) { [Math]::Round(((Get-Date).ToUniversalTime() - $lastSignIn.ToUniversalTime()).TotalDays, 0) } else { 9999 }
                    AccountEnabled      = [bool]$u.accountEnabled
                    ExternalUserState   = [string]$u.externalUserState
                })
            }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    } catch { Write-ErrorMsg "Could not enumerate guests: $($_.Exception.Message)"; return @() }

    $filtered = $out
    if ($Domain)            { $d = $Domain.ToLowerInvariant(); $filtered = $filtered | Where-Object { $_.Domains -and ($_.Domains -split ',\s*') -contains $d } }
    if ($MinAgeDays)        { $filtered = $filtered | Where-Object { $_.AgeDays -ge $MinAgeDays } }
    if ($MaxLastSignInDays) { $filtered = $filtered | Where-Object { $_.DaysSinceSignIn -le $MaxLastSignInDays } }
    return @($filtered | Sort-Object DaysSinceSignIn -Descending)
}

function Get-StaleGuests {
    param([int]$DaysSinceSignIn = 90)
    return @(Get-Guests | Where-Object { $_.DaysSinceSignIn -ge $DaysSinceSignIn })
}

function Get-GuestsByInviter {
    <#
        Pivot. externalUserState is the canonical field but it
        doesn't carry the inviter -- we fall back to InvitedBy
        from createdDateTime activity in the audit log if available,
        else group by Domains.
    #>
    $guests = Get-Guests
    return @($guests | Group-Object -Property Domains | ForEach-Object {
        [PSCustomObject]@{
            Group     = $_.Name
            GuestCount= $_.Count
            UPNs      = (($_.Group | Select-Object -First 5 | ForEach-Object { $_.UPN }) -join '; ') + $(if ($_.Count -gt 5) { " (+$($_.Count - 5) more)" } else { '' })
        }
    } | Sort-Object GuestCount -Descending)
}

function Get-GuestsByDomain {
    $guests = Get-Guests
    $bucket = @{}
    foreach ($g in $guests) {
        $doms = if ($g.Domains) { ($g.Domains -split ',\s*') } else { @('(no domain)') }
        foreach ($d in $doms) {
            if (-not $bucket.ContainsKey($d)) { $bucket[$d] = New-Object System.Collections.ArrayList }
            [void]$bucket[$d].Add($g.UPN)
        }
    }
    return @($bucket.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{ Domain = $_.Key; GuestCount = $_.Value.Count; ExampleUPNs = (($_.Value | Select-Object -First 5) -join '; ') }
    } | Sort-Object GuestCount -Descending)
}

# ============================================================
#  Recertification campaigns
# ============================================================

function Get-RecertStatePath {
    $dir = Get-StateDirectory
    if (-not $dir) { return $null }
    return Join-Path $dir 'guest-recerts.json'
}

function Read-RecertState {
    $p = Get-RecertStatePath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return @() }
    try { return @((Get-Content -LiteralPath $p -Raw | ConvertFrom-Json)) } catch { return @() }
}

function Write-RecertState {
    param([Parameter(Mandatory)][array]$Records)
    $p = Get-RecertStatePath
    if (-not $p) { return }
    try { ($Records | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $p -Encoding UTF8 -Force }
    catch { Write-Warn "Could not write recert state: $_" }
}

function Send-GuestRecertEmail {
    <#
        HTML mail to the manager with a recert decision form
        (yes/no, simple text reply expected). The manager's reply
        is reviewed manually -- this commit doesn't wire up a
        webhook callback.
    #>
    param(
        [Parameter(Mandatory)][string]$ManagerUPN,
        [Parameter(Mandatory)][PSCustomObject]$Guest,
        [string]$CampaignId = ''
    )
    $body = @"
<html><body style='font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222'>
<p>Please recertify access for the guest user below.</p>
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px'>
  <tr><th align='left'>UPN</th><td>$([System.Net.WebUtility]::HtmlEncode($Guest.UPN))</td></tr>
  <tr><th align='left'>Display name</th><td>$([System.Net.WebUtility]::HtmlEncode($Guest.DisplayName))</td></tr>
  <tr><th align='left'>Created</th><td>$($Guest.CreatedUtc) UTC (age $($Guest.AgeDays) days)</td></tr>
  <tr><th align='left'>Last sign-in</th><td>$($Guest.LastSignInUtc) UTC ($($Guest.DaysSinceSignIn) days ago)</td></tr>
  <tr><th align='left'>Account enabled</th><td>$($Guest.AccountEnabled)</td></tr>
</table>
<p><b>Action:</b> reply <b>YES</b> to keep this guest, or <b>NO</b> to remove. Decisions are reviewed manually by IT.</p>
<p>Campaign id: $CampaignId</p>
<p style='color:#666;font-size:12px'>Sent automatically by M365 Manager.</p>
</body></html>
"@
    if (Get-Command Send-Email -ErrorAction SilentlyContinue) {
        return [bool] (Send-Email -To @($ManagerUPN) -Subject "[Recertify] Guest access for $($Guest.UPN)" -Body $body)
    }
    # Standalone fallback for non-Phase-4 environments.
    $message = @{
        message = @{
            subject      = "[Recertify] Guest access for $($Guest.UPN)"
            body         = @{ contentType = "HTML"; content = $body }
            toRecipients = @(@{ emailAddress = @{ address = $ManagerUPN } })
        }
        saveToSentItems = $true
    } | ConvertTo-Json -Depth 10
    return Invoke-Action `
        -Description ("Send guest-recertification email for {0} to manager {1}" -f $Guest.UPN, $ManagerUPN) `
        -ActionType 'SendGuestRecertEmail' `
        -Target @{ guestUpn = $Guest.UPN; managerUpn = $ManagerUPN; campaignId = $CampaignId } `
        -NoUndoReason 'Email send is irreversible.' `
        -Action {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/sendMail" -Body $message -ContentType 'application/json' -ErrorAction Stop | Out-Null
            $true
        }
}

function Invoke-GuestRecertification {
    <#
        CSV columns: UPN, ManagerUPN
        Emails each manager + appends to the recert state file.
        The operator handles replies out-of-band and processes
        them via Show-PendingRecerts.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { Write-Warn "Empty CSV."; return }

    if (-not (Connect-ForTask 'GuestUsers')) { return }

    $campaignId = "recert-{0}-{1:X}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), (Get-Random -Maximum 65535)
    $state = Read-RecertState
    $sent = 0; $failed = 0
    foreach ($r in $rows) {
        $upn = [string]$r.UPN
        $mgr = [string]$r.ManagerUPN
        if (-not $upn -or -not $mgr) { Write-Warn "Skip row -- needs UPN + ManagerUPN."; $failed++; continue }
        # Resolve the guest (id needed for downstream Remove-Guest)
        $guest = $null
        try {
            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn?`$select=id,userPrincipalName,displayName,createdDateTime,accountEnabled,signInActivity" -ErrorAction Stop
            $created = if ($g.createdDateTime) { [DateTime]$g.createdDateTime } else { $null }
            $lastSi  = if ($g.signInActivity.lastSignInDateTime) { [DateTime]$g.signInActivity.lastSignInDateTime } else { $null }
            $guest = [PSCustomObject]@{
                Id              = $g.id
                UPN             = $g.userPrincipalName
                DisplayName     = $g.displayName
                CreatedUtc      = $created
                AgeDays         = if ($created) { [Math]::Round(((Get-Date).ToUniversalTime() - $created.ToUniversalTime()).TotalDays, 0) } else { '?' }
                LastSignInUtc   = $lastSi
                DaysSinceSignIn = if ($lastSi) { [Math]::Round(((Get-Date).ToUniversalTime() - $lastSi.ToUniversalTime()).TotalDays, 0) } else { '?' }
                AccountEnabled  = [bool]$g.accountEnabled
            }
        } catch { Write-Warn "Could not resolve $upn -- $($_.Exception.Message)"; $failed++; continue }

        $ok = Send-GuestRecertEmail -ManagerUPN $mgr -Guest $guest -CampaignId $campaignId
        if ($ok) {
            $sent++
            $state += [PSCustomObject]@{
                campaignId  = $campaignId
                guestId     = $guest.Id
                guestUpn    = $guest.UPN
                managerUpn  = $mgr
                queuedAt    = (Get-Date).ToUniversalTime().ToString('o')
                state       = 'pending'   # pending | keep | remove
                decisionBy  = $null
                decisionAt  = $null
                notes       = ''
            }
        } else { $failed++ }
    }
    Write-RecertState -Records $state
    Write-Success "Campaign $campaignId queued: $sent email(s) sent, $failed skipped."
}

function Show-PendingRecerts {
    <#
        Viewer for the recert state file. Lets the operator mark
        a row 'keep' or 'remove' (decision recorded with timestamp +
        decisionBy). When the operator chooses 'remove', the guest
        is fed through Remove-Guest immediately.
    #>
    $state = Read-RecertState
    if ($state.Count -eq 0) { Write-InfoMsg "No recertification records on disk yet."; return }
    while ($true) {
        $pending = @($state | Where-Object { $_.state -eq 'pending' } | Sort-Object queuedAt)
        Write-SectionHeader "Pending guest recertifications"
        if ($pending.Count -eq 0) { Write-InfoMsg "(no pending rows -- all decided)"; break }
        for ($i = 0; $i -lt $pending.Count; $i++) {
            $p = $pending[$i]
            Write-Host ("  [{0,2}] {1}  guest={2}  mgr={3}  queued={4}" -f ($i+1), $p.campaignId, $p.guestUpn, $p.managerUpn, $p.queuedAt) -ForegroundColor White
        }
        Write-Host ""
        $idx = Read-UserInput "Row to decide (1-$($pending.Count); blank = back)"
        if ([string]::IsNullOrWhiteSpace($idx)) { break }
        $n = 0; if (-not [int]::TryParse($idx, [ref]$n) -or $n -lt 1 -or $n -gt $pending.Count) { Write-Warn "Out of range."; continue }
        $rec = $pending[$n-1]
        $decision = Show-Menu -Title "Decision for $($rec.guestUpn)" -Options @("Keep (mark certified)","Remove (run Remove-Guest)","Defer (leave pending)") -BackLabel "Cancel"
        if ($decision -eq -1) { continue }
        switch ($decision) {
            0 {
                $rec.state = 'keep'; $rec.decisionBy = (Get-MgContext).Account; $rec.decisionAt = (Get-Date).ToUniversalTime().ToString('o')
                Write-RecertState -Records $state; Write-Success "Marked keep."
            }
            1 {
                $reason = Read-UserInput "Reason (recorded in audit)"
                Remove-Guest -UPN $rec.guestUpn -Reason ("Recert decision: " + $reason) | Out-Null
                $rec.state = 'remove'; $rec.decisionBy = (Get-MgContext).Account; $rec.decisionAt = (Get-Date).ToUniversalTime().ToString('o'); $rec.notes = $reason
                Write-RecertState -Records $state
            }
            2 { } # defer no-op
        }
    }
}

# ============================================================
#  Removal
# ============================================================

function Remove-Guest {
    <#
        Proper teardown:
          1. Revoke outbound shares the guest CREATED (uses
             SharePoint.ps1's Get-UserOutboundShares + Revoke-Share)
          2. Remove from groups (Get-MgUserMemberOf + DELETE per group)
          3. Remove from teams (TeamsManager.ps1's Remove-UserFromTeam
             per joined team) -- same as part of #2 but uses the
             Teams-aware helper so promote/demote-then-remove logic
             runs for any team where the guest is an owner
          4. Delete the user via Graph
        Graph's delete is reversible within 30 days via
        /directory/deletedItems/{id}/restore -- we surface this in
        the audit entry's noUndoReason.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$Reason
    )
    Write-SectionHeader "Remove guest: $UPN"
    Write-Warn "Reason: $Reason"

    if (-not (Connect-ForTask 'GuestUsers')) { return }

    # 1. Outbound shares (only if SharePoint module is loaded + UAL works)
    if (Get-Command Get-UserOutboundShares -ErrorAction SilentlyContinue) {
        try { $null = Invoke-SharePointOffboardCleanup -LeaverUPN $UPN -LookbackDays 365 }
        catch { Write-Warn "Outbound-share cleanup failed: $_" }
    }

    $userId = $null
    try { $userId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id" -ErrorAction Stop).id }
    catch { Write-ErrorMsg "Could not resolve $UPN -- $_"; return }

    # 2. Groups
    try {
        $memberOf = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userId/memberOf?`$select=id,@odata.type,displayName" -ErrorAction Stop).value)
    } catch { $memberOf = @() }
    foreach ($m in $memberOf) {
        if (-not $m.id) { continue }
        Invoke-Action `
            -Description ("Remove guest {0} from group '{1}'" -f $UPN, $m.displayName) `
            -ActionType 'RemoveFromGroup' `
            -Target @{ userId = [string]$userId; userUpn = $UPN; groupId = [string]$m.id; groupName = [string]$m.displayName } `
            -ReverseType 'AddToGroup' `
            -ReverseDescription ("Re-add guest {0} to group '{1}'" -f $UPN, $m.displayName) `
            -Action {
                try { Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($m.id)/members/$userId/`$ref" -ErrorAction Stop | Out-Null; $true }
                catch { if ($_.Exception.Message -match 'does not exist|not found') { 'missing' } else { throw } }
            } | Out-Null
    }

    # 3. Teams (covers any team where the guest was an owner)
    if (Get-Command Invoke-TeamsOffboardTransfer -ErrorAction SilentlyContinue) {
        try { Invoke-TeamsOffboardTransfer -LeaverUPN $UPN | Out-Null } catch { Write-Warn "Teams cleanup failed: $_" }
    }

    # 4. Delete the user
    Invoke-Action `
        -Description ("DELETE guest user {0} ({1})" -f $UPN, $Reason) `
        -ActionType 'DeleteGuestUser' `
        -Target @{ userId = [string]$userId; userUpn = $UPN; reason = $Reason } `
        -NoUndoReason 'User deletion goes to /directory/deletedItems for 30 days. To restore: POST /directory/deletedItems/{userId}/restore. After 30 days, permanent.' `
        -Action {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$userId" -ErrorAction Stop | Out-Null
            $true
        } | Out-Null
    Write-Success "Guest $UPN removed."
}

function Invoke-BulkGuestRemoval {
    <#
        CSV: UPN, Reason
        Standard validate-then-execute. Result CSV written next
        to input.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$WhatIf)
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { Write-Warn "Empty CSV."; return }

    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    try {
        $results = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $r = $rows[$i]
            $upn = [string]$r.UPN
            $reason = if ($r.Reason) { [string]$r.Reason } else { 'Bulk removal' }
            Write-Progress -Activity "Bulk guest removal" -Status $upn -PercentComplete (($i / $rows.Count) * 100)
            $status = 'Pending'
            try { Remove-Guest -UPN $upn -Reason $reason; $status = if (Get-PreviewMode) { 'Preview' } else { 'Removed' } }
            catch { $status = "Failed: $($_.Exception.Message)" }
            [void]$results.Add([PSCustomObject]@{ UPN = $upn; Status = $status; Reason = $reason })
        }
        Write-Progress -Activity "Bulk guest removal" -Completed
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-guest-removal-$stamp.csv")
        $results | Export-Csv -LiteralPath $out -NoTypeInformation -Force
        Write-Success "Result CSV: $out"
    } finally { Set-PreviewMode -Enabled $previousMode }
}

# ============================================================
#  Menu
# ============================================================

function Start-GuestUsersMenu {
    while ($true) {
        $sel = Show-Menu -Title "Guest Users" -Options @(
            "List guests",
            "Stale guests (90+ days no sign-in)",
            "Group guests by domain",
            "Pivot guests by inviter / domain",
            "Send recertification campaign from CSV...",
            "View / decide pending recertifications",
            "Remove guest (single user)...",
            "Bulk guest removal from CSV..."
        ) -BackLabel "Back"
        switch ($sel) {
            0 { Get-Guests | Format-Table -AutoSize; Pause-ForUser }
            1 {
                $dt = Read-UserInput "Days threshold (default 90)"; $d = 90; [int]::TryParse($dt,[ref]$d) | Out-Null
                Get-StaleGuests -DaysSinceSignIn $d | Format-Table -AutoSize
                Pause-ForUser
            }
            2 { Get-GuestsByDomain | Format-Table -AutoSize; Pause-ForUser }
            3 { Get-GuestsByInviter | Format-Table -AutoSize; Pause-ForUser }
            4 {
                $p = Read-UserInput "Path to CSV (UPN, ManagerUPN)"
                if ($p) { Invoke-GuestRecertification -Path $p.Trim('"').Trim("'") }
                Pause-ForUser
            }
            5 { Show-PendingRecerts; Pause-ForUser }
            6 {
                $upn = Read-UserInput "Guest UPN"; if (-not $upn) { continue }
                $reason = Read-UserInput "Reason"
                if (-not $reason) { Write-Warn "Reason is required for audit."; continue }
                if (Confirm-Action "DELETE guest $upn ($reason)?") { Remove-Guest -UPN $upn -Reason $reason }
                Pause-ForUser
            }
            7 {
                $p = Read-UserInput "Path to CSV (UPN, Reason)"
                if (-not $p) { continue }
                $dry = Confirm-Action "Run as DRY-RUN first?"
                Invoke-BulkGuestRemoval -Path $p.Trim('"').Trim("'") -WhatIf:$dry
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
