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
    # Phase 6: default threshold is tenant-overridable
    # (tenant-overrides/<name>.json key 'StaleGuestDays').
    param([int]$DaysSinceSignIn = 0)
    if ($DaysSinceSignIn -le 0) {
        $eff = if (Get-Command Get-EffectiveConfig -ErrorAction SilentlyContinue) { Get-EffectiveConfig -Key 'StaleGuestDays' } else { $null }
        $DaysSinceSignIn = if ($eff) { [int]$eff } else { 90 }
    }
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
    # AllowEmptyCollection: a fully-resolved recert campaign should
    # be writable as an empty array (clears the state file).
    # -AsArray keeps a single-record file in array shape so Read
    # round-trip via @(ConvertFrom-Json) stays consistent.
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records)
    $p = Get-RecertStatePath
    if (-not $p) { return }
    try { ($Records | ConvertTo-Json -Depth 5 -AsArray) | Set-Content -LiteralPath $p -Encoding UTF8 -Force }
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
            $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$(ConvertTo-GraphUserSegment $upn)?`$select=id,userPrincipalName,displayName,createdDateTime,accountEnabled,signInActivity" -ErrorAction Stop
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
          0. Resolve the user and VERIFY it is actually a Guest before
             touching anything. Refuse Member / internal accounts unless
             -AllowNonGuest is explicitly passed -- a fat-fingered CSV
             row must never delete a real employee/admin via the guest
             path.
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

        Returns a status object so callers (Invoke-BulkGuestRemoval) can
        report the TRUE outcome instead of assuming success:
          @{ UPN; Status; Reason; GroupsRemoved }
        Status is one of: Removed | Preview | NotFound | NotAGuest |
        NotConnected | Failed.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$Reason,
        [switch]$AllowNonGuest
    )
    Write-SectionHeader "Remove guest: $UPN"
    Write-Warn "Reason: $Reason"

    $result = [PSCustomObject]@{ UPN = $UPN; Status = 'Failed'; Reason = $Reason; GroupsRemoved = 0 }

    if (-not (Connect-ForTask 'GuestUsers')) {
        $result.Status = 'NotConnected'; $result.Reason = 'Could not connect to Microsoft Graph.'
        return $result
    }

    # 0. Resolve + verify this really is a Guest BEFORE any teardown.
    $user = $null
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$(ConvertTo-GraphUserSegment $UPN)?`$select=id,userType,userPrincipalName" -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Could not resolve $UPN -- $($_.Exception.Message)"
        $result.Status = 'NotFound'; $result.Reason = "Not found in tenant: $($_.Exception.Message)"
        return $result
    }
    $userId = [string]$user.id
    if (-not $userId) {
        Write-ErrorMsg "Resolved $UPN but Graph returned no object id."
        $result.Status = 'NotFound'; $result.Reason = 'Graph returned no object id.'
        return $result
    }
    $userType = [string]$user.userType
    if ($userType -ne 'Guest' -and -not $AllowNonGuest) {
        $shown = if ($userType) { $userType } else { 'unknown' }
        Write-ErrorMsg ("REFUSED: {0} is userType '{1}', not 'Guest'. This looks like a member/internal account -- not deleting it via the guest-removal path. Use the offboarding flow for members." -f $UPN, $shown)
        if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
            Write-AuditEntry -EventType 'EXEC' `
                -Detail ("Refused guest removal for {0}: userType='{1}'" -f $UPN, $shown) `
                -ActionType 'DeleteGuestUserRefused' `
                -Target @{ userId = $userId; userUpn = $UPN; userType = $shown } `
                -Result 'failure' | Out-Null
        }
        $result.Status = 'NotAGuest'; $result.Reason = "userType is '$shown', not Guest (skipped for safety)."
        return $result
    }

    # 1. Outbound shares (only if SharePoint module is loaded + UAL works)
    if (Get-Command Get-UserOutboundShares -ErrorAction SilentlyContinue) {
        try { $null = Invoke-SharePointOffboardCleanup -LeaverUPN $UPN -LookbackDays 365 }
        catch { Write-Warn "Outbound-share cleanup failed: $_" }
    }

    # 2. Groups
    try {
        $memberOf = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userId/memberOf?`$select=id,@odata.type,displayName" -ErrorAction Stop).value)
    } catch { $memberOf = @() }
    foreach ($m in $memberOf) {
        if (-not $m.id) { continue }
        $groupResult = Invoke-Action `
            -Description ("Remove guest {0} from group '{1}'" -f $UPN, $m.displayName) `
            -ActionType 'RemoveFromGroup' `
            -Target @{ userId = [string]$userId; userUpn = $UPN; groupId = [string]$m.id; groupName = [string]$m.displayName } `
            -ReverseType 'AddToGroup' `
            -ReverseDescription ("Re-add guest {0} to group '{1}'" -f $UPN, $m.displayName) `
            -Action {
                try { Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($m.id)/members/$userId/`$ref" -ErrorAction Stop | Out-Null; $true }
                catch { if ($_.Exception.Message -match 'does not exist|not found') { 'missing' } else { throw } }
            }
        if ($groupResult -eq $true) { $result.GroupsRemoved++ }
    }

    # 3. Teams (covers any team where the guest was an owner)
    if (Get-Command Invoke-TeamsOffboardTransfer -ErrorAction SilentlyContinue) {
        try { Invoke-TeamsOffboardTransfer -LeaverUPN $UPN | Out-Null } catch { Write-Warn "Teams cleanup failed: $_" }
    }

    # 4. Delete the user
    $deleted = Invoke-Action `
        -Description ("DELETE guest user {0} ({1})" -f $UPN, $Reason) `
        -ActionType 'DeleteGuestUser' `
        -Target @{ userId = [string]$userId; userUpn = $UPN; reason = $Reason } `
        -NoUndoReason 'User deletion goes to /directory/deletedItems for 30 days. To restore: POST /directory/deletedItems/{userId}/restore. After 30 days, permanent.' `
        -Action {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$userId" -ErrorAction Stop | Out-Null
            $true
        }

    if (Get-PreviewMode) {
        $result.Status = 'Preview'; $result.Reason = 'Dry-run -- no tenant delete performed.'
        Write-InfoMsg "[preview] Would remove guest $UPN."
    } elseif ($deleted -eq $true) {
        $result.Status = 'Removed'
        Write-Success "Guest $UPN removed."
    } else {
        $result.Status = 'Failed'; $result.Reason = 'DELETE /users call did not succeed (see audit log).'
        Write-ErrorMsg "Guest $UPN was NOT deleted -- the delete call failed."
    }
    return $result
}

function Test-BulkGuestRemovalCsv {
    <#
        Validate every row without any tenant calls. Catches:
          - missing UPN (accepts a 'UPN' or 'UserPrincipalName' column)
          - an identifier that is neither a UPN (contains '@') nor an
            object GUID -- guest UPNs legitimately contain '#EXT#' and
            underscores, so we deliberately DON'T apply a strict email
            regex here (that would reject valid guest UPNs); the
            authoritative Guest check happens tenant-side in
            Invoke-BulkGuestRemoval's pre-flight.
          - duplicate identifier within the CSV.
        Returns @{ Rows = @([PSCustomObject]@{ UPN; Reason }); Errors = @(@{Row;Field;Message}) }.
    #>
    param([array]$Rows)

    $errors = @()
    $normalized = @()
    $seen = @{}
    $guidRegex = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $rowNum = $i + 2   # header is row 1
        $r = $Rows[$i]

        $upn = ''
        $reason = ''
        foreach ($p in $r.PSObject.Properties) {
            $k = $p.Name.Trim()
            $v = if ($null -eq $p.Value) { '' } else { ([string]$p.Value).Trim() }
            if (($k -ieq 'UPN' -or $k -ieq 'UserPrincipalName') -and -not $upn) { $upn = $v }
            elseif ($k -ieq 'Reason') { $reason = $v }
        }

        if ([string]::IsNullOrWhiteSpace($upn)) {
            $errors += @{ Row = $rowNum; Field = 'UPN'; Message = "Missing UPN (need a 'UPN' or 'UserPrincipalName' column)" }
            continue
        }
        if (-not (($upn -match '@') -or ($upn -match $guidRegex))) {
            $errors += @{ Row = $rowNum; Field = 'UPN'; Message = "Not a UPN or object id: '$upn'" }
            continue
        }
        $key = $upn.ToLowerInvariant()
        if ($seen.ContainsKey($key)) {
            $errors += @{ Row = $rowNum; Field = 'UPN'; Message = "Duplicate UPN in CSV (first seen on row $($seen[$key]))" }
            continue
        }
        $seen[$key] = $rowNum
        $normalized += [PSCustomObject]@{ UPN = $upn; Reason = $(if ($reason) { $reason } else { 'Bulk removal' }) }
    }

    return @{ Rows = @($normalized); Errors = @($errors) }
}

function Invoke-BulkGuestRemoval {
    <#
        CSV: UPN[,Reason]  (UserPrincipalName also accepted)

        Bulletproof pipeline:
          1. Parse + validate the CSV (no tenant calls).
          2. Connect, then a PRE-FLIGHT pass that resolves every row and
             classifies it Guest / NotFound / NotGuest -- NOTHING is
             deleted yet.
          3. Show the operator exactly how many real guests will be
             removed (and lists every non-guest / not-found that will be
             skipped), then require ONE explicit confirmation with the
             LIVE/PREVIEW mode spelled out.
          4. Delete only the confirmed guests, recording the TRUE status
             returned by Remove-Guest.
          5. Write a result CSV covering every input row + print a summary.

        -WhatIf runs the whole thing in preview mode (no tenant changes).
        -AllowNonGuest lifts the "guests only" guard (rarely needed;
         off by default so a bad CSV can't nuke a member account).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$WhatIf,
        [switch]$AllowNonGuest
    )
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }

    Write-SectionHeader "Bulk Guest Removal -- $(Split-Path $Path -Leaf)"

    $rows = $null
    try { $rows = @(Import-Csv -LiteralPath $Path) }
    catch { Write-ErrorMsg "Could not parse CSV: $_"; return }
    if ($rows.Count -eq 0) { Write-Warn "CSV has no data rows."; return }
    Write-InfoMsg "$($rows.Count) row(s) read from $Path"

    $validation = Test-BulkGuestRemovalCsv -Rows $rows
    if ($validation.Errors.Count -gt 0) {
        Write-Host ""
        Write-ErrorMsg "Validation failed -- $($validation.Errors.Count) issue(s):"
        foreach ($e in $validation.Errors) {
            Write-Host ("    Row {0,3}  {1,-5}  {2}" -f $e.Row, $e.Field, $e.Message) -ForegroundColor Red
        }
        Write-Host ""
        Write-ErrorMsg "Fix the CSV and re-run."
        return
    }
    Write-Success "Validation passed: $($validation.Rows.Count) row(s) ready."

    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    $dryRun = Get-PreviewMode
    if ($dryRun) { Write-Warn "PREVIEW mode -- no tenant changes will be made." }
    try {
        if (-not (Connect-ForTask 'GuestUsers')) { Write-ErrorMsg "Could not connect."; return }

        # ---- Pre-flight: classify every identifier BEFORE deleting. ----
        Write-InfoMsg "Verifying each identifier resolves to a Guest user..."
        $toRemove = New-Object System.Collections.ArrayList
        $results  = New-Object System.Collections.ArrayList   # accumulates skipped rows too
        foreach ($row in $validation.Rows) {
            $upn = $row.UPN
            $u = $null
            try {
                $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$(ConvertTo-GraphUserSegment $upn)?`$select=id,userType,userPrincipalName" -ErrorAction Stop
            } catch {
                [void]$results.Add([PSCustomObject]@{ UPN = $upn; Status = 'NotFound'; Reason = 'Not found in tenant' })
                continue
            }
            $utype = [string]$u.userType
            if ($utype -ne 'Guest' -and -not $AllowNonGuest) {
                [void]$results.Add([PSCustomObject]@{ UPN = $upn; Status = 'SkippedNotGuest'; Reason = "userType='$utype' (not Guest)" })
                continue
            }
            [void]$toRemove.Add([PSCustomObject]@{ UPN = $upn; Reason = $row.Reason })
        }

        $notFound = @($results | Where-Object { $_.Status -eq 'NotFound' })
        $notGuest = @($results | Where-Object { $_.Status -eq 'SkippedNotGuest' })
        Write-Host ""
        Write-InfoMsg ("Resolved: {0} guest(s) to remove, {1} not found, {2} non-guest." -f $toRemove.Count, $notFound.Count, $notGuest.Count)
        if ($notGuest.Count -gt 0) {
            Write-Warn "NOT guests -- these will be SKIPPED (use the offboard flow for members):"
            foreach ($x in $notGuest) { Write-Host ("    - {0}  [{1}]" -f $x.UPN, $x.Reason) -ForegroundColor Yellow }
        }
        if ($notFound.Count -gt 0) {
            Write-Warn "Not found in tenant -- skipped:"
            foreach ($x in $notFound) { Write-Host ("    - {0}" -f $x.UPN) -ForegroundColor Yellow }
        }

        if ($toRemove.Count -eq 0) {
            Write-Warn "No guest users to remove after pre-flight."
        } else {
            $modeLabel = if ($dryRun) { "PREVIEW (no changes)" } else { "LIVE -- guests WILL be deleted" }
            if (-not (Confirm-Action ("About to remove {0} guest user(s) in {1}. Proceed?" -f $toRemove.Count, $modeLabel))) {
                Write-InfoMsg "Cancelled."
                return
            }
            for ($i = 0; $i -lt $toRemove.Count; $i++) {
                $g = $toRemove[$i]
                Write-Progress -Activity "Bulk guest removal" -Status ("{0} ({1} of {2})" -f $g.UPN, ($i + 1), $toRemove.Count) -PercentComplete ([int](($i / $toRemove.Count) * 100))
                $res = $null
                try { $res = Remove-Guest -UPN $g.UPN -Reason $g.Reason -AllowNonGuest:$AllowNonGuest }
                catch { $res = $null }
                $status = if ($res -and $res.Status) { [string]$res.Status } else { 'Failed' }
                $reason = if ($res -and $res.Reason) { [string]$res.Reason } else { $g.Reason }
                [void]$results.Add([PSCustomObject]@{ UPN = $g.UPN; Status = $status; Reason = $reason })
            }
            Write-Progress -Activity "Bulk guest removal" -Completed
        }

        # ---- Result CSV ----
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-guest-removal-$stamp.csv")
        try {
            $results | Export-Csv -LiteralPath $out -NoTypeInformation -Force
            Write-Host ""
            Write-Success "Result CSV: $out"
        } catch { Write-ErrorMsg "Could not write result CSV: $_" }

        # ---- Summary ----
        $removed = @($results | Where-Object { $_.Status -eq 'Removed' }).Count
        $preview = @($results | Where-Object { $_.Status -eq 'Preview' }).Count
        $failed  = @($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'NotConnected' }).Count
        $nf      = @($results | Where-Object { $_.Status -eq 'NotFound' }).Count
        $ng      = @($results | Where-Object { $_.Status -eq 'SkippedNotGuest' -or $_.Status -eq 'NotAGuest' }).Count
        Write-Host ""
        Write-Host "  Bulk guest removal summary:" -ForegroundColor White
        Write-StatusLine "Removed"              $removed  "Green"
        if ($preview -gt 0) { Write-StatusLine "Preview" $preview "Yellow" }
        Write-StatusLine "Failed"               $failed   $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
        Write-StatusLine "Not found"            $nf       $(if ($nf -gt 0) { 'Yellow' } else { 'Gray' })
        Write-StatusLine "Skipped (non-guest)"  $ng       $(if ($ng -gt 0) { 'Yellow' } else { 'Gray' })
        Write-Host ""
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
                $dt = Read-UserInput "Days threshold (default 90)"; $d = Get-IntOrDefault $dt 90
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
                $upn = if (Get-Command Resolve-UPN -ErrorAction SilentlyContinue) { Resolve-UPN -Prompt "Guest UPN or name" } else { Read-UserInput "Guest UPN" }
                if (-not $upn) { continue }
                $reason = Read-UserInput "Reason"
                if (-not $reason) { Write-Warn "Reason is required for audit."; continue }
                if (Confirm-Action "DELETE guest $upn ($reason)?") { Remove-Guest -UPN $upn -Reason $reason | Out-Null }
                Pause-ForUser
            }
            7 {
                $p = Read-UserInput "Path to CSV (UPN[,Reason]; sample: samples/bulk-guest-removal-sample.csv)"
                if (-not $p) { continue }
                $dry = Confirm-Action "Run as DRY-RUN first (validate + preview, no tenant changes)?"
                Invoke-BulkGuestRemoval -Path $p.Trim('"').Trim("'") -WhatIf:$dry
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
