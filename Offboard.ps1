# ============================================================
#  Offboard.ps1 — User offboarding (canonical 12-step flow)
#
#  Order (matches Phase 3 spec):
#     0  Revoke MFA methods
#     1  Block sign-in            (revoke sessions + AccountEnabled=$false)
#     2  Out-of-office + forwarding
#     3  Remove from security groups
#     4  Remove from distribution lists
#     5  Remove direct license assignments
#     6  Convert mailbox to shared (conditional)
#     6b Grant mailbox access to delegates (conditional)
#     7  Teams ownership transfer
#     8  Revoke outbound SharePoint shares
#     9  OneDrive handoff
#    10  Manager summary email
#    11  Final audit summary line
#
#  Every step is independently skippable via per-operator confirm.
#  BulkOffboard.ps1 mirrors this flow with column-level toggles.
# ============================================================

function Invoke-OffboardRemoveSecurityGroups {
    <#
        Remove the leaver from every pure security group (security-
        enabled, not mail-enabled, not Unified). Audited per group.
    #>
    param([Parameter(Mandatory)][string]$UPN, [Parameter(Mandatory)][string]$UserId)
    $count = 0; $failed = 0
    try {
        $members = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/memberOf?`$select=id,displayName,securityEnabled,mailEnabled,groupTypes" -ErrorAction Stop).value)
    } catch { Write-ErrorMsg "memberOf enumeration failed: $($_.Exception.Message)"; return @{ Removed=0; Failed=0 } }
    foreach ($m in $members) {
        if ([string]$m.'@odata.type' -ne '#microsoft.graph.group') { continue }
        $isUnified = ($m.groupTypes -and ($m.groupTypes -contains 'Unified'))
        if (-not $m.securityEnabled -or $m.mailEnabled -or $isUnified) { continue }
        $ok = Invoke-Action `
            -Description ("Remove {0} from security group '{1}'" -f $UPN, $m.displayName) `
            -ActionType 'RemoveFromGroup' `
            -Target @{ userId = $UserId; userUpn = $UPN; groupId = [string]$m.id; groupName = [string]$m.displayName } `
            -ReverseType 'AddToGroup' `
            -ReverseDescription ("Re-add {0} to security group '{1}'" -f $UPN, $m.displayName) `
            -Action {
                try { Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($m.id)/members/$UserId/`$ref" -ErrorAction Stop | Out-Null; $true }
                catch { if ($_.Exception.Message -match 'does not exist|not found') { 'missing' } else { throw } }
            }
        if ($ok) { $count++ } else { $failed++ }
    }
    return @{ Removed = $count; Failed = $failed }
}

function Invoke-OffboardRemoveDistributionLists {
    <#
        Remove from mail-enabled groups that are NOT pure security
        (so DLs + M365 groups that are NOT team-backed; team-backed
        ones are handled by Step 7). Audited per group.
    #>
    param([Parameter(Mandatory)][string]$UPN, [Parameter(Mandatory)][string]$UserId)
    $count = 0; $failed = 0
    try {
        $members = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/memberOf?`$select=id,displayName,securityEnabled,mailEnabled,groupTypes,resourceProvisioningOptions" -ErrorAction Stop).value)
    } catch { Write-ErrorMsg "memberOf enumeration failed: $($_.Exception.Message)"; return @{ Removed=0; Failed=0 } }
    foreach ($m in $members) {
        if ([string]$m.'@odata.type' -ne '#microsoft.graph.group') { continue }
        if (-not $m.mailEnabled) { continue }
        $isTeam = $false
        if ($m.resourceProvisioningOptions -and ($m.resourceProvisioningOptions -contains 'Team')) { $isTeam = $true }
        if ($isTeam) { continue }   # Step 7 handles teams
        $ok = Invoke-Action `
            -Description ("Remove {0} from distribution list / M365 group '{1}'" -f $UPN, $m.displayName) `
            -ActionType 'RemoveFromGroup' `
            -Target @{ userId = $UserId; userUpn = $UPN; groupId = [string]$m.id; groupName = [string]$m.displayName } `
            -ReverseType 'AddToGroup' `
            -ReverseDescription ("Re-add {0} to '{1}'" -f $UPN, $m.displayName) `
            -Action {
                try { Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($m.id)/members/$UserId/`$ref" -ErrorAction Stop | Out-Null; $true }
                catch { if ($_.Exception.Message -match 'does not exist|not found') { 'missing' } else { throw } }
            }
        if ($ok) { $count++ } else { $failed++ }
    }
    return @{ Removed = $count; Failed = $failed }
}

function Start-Offboard {
    Write-SectionHeader "User Offboarding"

    # ---- Data collection (unchanged from Phase 2) ----
    $fields = @("UserPrincipalName","ForwardingEmail","OOOMessage")
    $choice = Show-Menu -Title "How to provide offboarding data?" -Options @("Parse from a text file","Enter manually") -BackLabel "Cancel"
    if ($choice -eq -1) { return }

    $userData = @{}
    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) { Write-ErrorMsg "File not found."; Pause-ForUser; return }
        foreach ($line in (Get-Content $filePath)) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key = $Matches[1].Trim(); $value = $Matches[2].Trim()
                switch -Regex ($key) {
                    'upn|email|user'               { $userData["UserPrincipalName"] = $value }
                    'forward'                      { $userData["ForwardingEmail"]   = $value }
                    'ooo|out.of.office|auto.reply' { $userData["OOOMessage"]        = $value }
                }
            }
        }
        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    } else {
        $userData["UserPrincipalName"] = Read-UserInput "Enter User Principal Name (email)"
        $userData["ForwardingEmail"] = ""; $userData["OOOMessage"] = ""
    }

    $upn = $userData["UserPrincipalName"]
    if ([string]::IsNullOrWhiteSpace($upn)) { Write-ErrorMsg "UPN is required."; Pause-ForUser; return }

    if (-not (Connect-ForTask "Offboard")) { Pause-ForUser; return }

    try {
        $user = Get-MgUser -UserId $upn -Property "Id,DisplayName,UserPrincipalName" -ErrorAction Stop
        Write-Success "User found: $($user.DisplayName) ($upn)"
    } catch { Write-ErrorMsg "User not found: $_"; Pause-ForUser; return }

    if (-not (Confirm-Action "Begin offboarding for $($user.DisplayName)?")) { Pause-ForUser; return }

    $summary = [ordered]@{ Leaver = $upn; DisplayName = $user.DisplayName }

    # ---- Step 0: Revoke MFA methods ----
    Write-SectionHeader "Step 0 - Revoke MFA Methods"
    if ((Get-Command Remove-AllAuthMethods -ErrorAction SilentlyContinue) -and (Confirm-Action "Revoke all MFA methods for $upn?")) {
        $revoked = Remove-AllAuthMethods -User $user.Id
        if (-not (Get-PreviewMode)) { Write-Success "$revoked MFA method(s) revoked." }
        $summary['MFA methods revoked'] = $revoked
    }

    # ---- Step 1: Block sign-in (revoke sessions then AccountEnabled=false) ----
    Write-SectionHeader "Step 1 - Block Sign-In"
    if (Confirm-Action "Revoke sessions and block sign-in for $upn?") {
        $ok1 = Invoke-Action `
            -Description ("Revoke sign-in sessions for {0}" -f $upn) `
            -ActionType 'RevokeSignInSessions' `
            -Target @{ userId = [string]$user.Id; userUpn = $upn } `
            -NoUndoReason 'Sign-in sessions, once revoked, cannot be restored.' `
            -Action { Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop; $true }
        $ok2 = Invoke-Action `
            -Description ("Block sign-in for {0}" -f $upn) `
            -ActionType 'BlockSignIn' `
            -Target @{ userId = [string]$user.Id; userUpn = $upn } `
            -ReverseType 'UnblockSignIn' `
            -ReverseDescription ("Unblock sign-in for {0}" -f $upn) `
            -Action { Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; $true }
        if ($ok1 -and $ok2 -and -not (Get-PreviewMode)) { Write-Success "Sessions revoked + sign-in blocked." }
        $summary['Sign-in blocked'] = if ($ok2) { 'yes' } else { 'failed' }
    }

    # ---- Step 2: OOO + forwarding ----
    Write-SectionHeader "Step 2 - Out-of-Office + Forwarding"
    if ([string]::IsNullOrWhiteSpace($userData["OOOMessage"])) { $userData["OOOMessage"] = Read-UserInput "OOO message (or 'skip')" }
    if ($userData["OOOMessage"] -and $userData["OOOMessage"] -ne 'skip') {
        if (Confirm-Action "Set auto-reply?" "Message: $($userData['OOOMessage'])") {
            $ok = Invoke-Action `
                -Description ("Set OOO auto-reply for {0}" -f $upn) -ActionType 'SetOOO' `
                -Target @{ identity = $upn } -ReverseType 'ClearOOO' `
                -ReverseDescription ("Clear OOO auto-reply on {0}" -f $upn) `
                -Action { Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $userData["OOOMessage"] -ExternalMessage $userData["OOOMessage"] -ErrorAction Stop; $true }
            if ($ok -and -not (Get-PreviewMode)) { Write-Success "OOO set." }
            $summary['OOO'] = if ($ok) { 'set' } else { 'failed' }
        }
    }
    if ([string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) { $userData["ForwardingEmail"] = Read-UserInput "Forwarding email (or 'skip')" }
    if ($userData["ForwardingEmail"] -and $userData["ForwardingEmail"] -ne 'skip') {
        $fwd = $userData["ForwardingEmail"]
        $dc = Show-Menu -Title "Delivery option" -Options @("Forward only","Forward and keep copy") -BackLabel "Skip"
        if ($dc -ne -1) {
            $keepCopy = ($dc -eq 1)
            if (Confirm-Action "Set forwarding to $fwd (keep copy: $keepCopy)?") {
                $ok = Invoke-Action `
                    -Description ("Set forwarding {0} -> {1} (keep copy: {2})" -f $upn, $fwd, $keepCopy) `
                    -ActionType 'SetForwarding' `
                    -Target @{ identity = $upn; forwardTo = $fwd; keepCopy = $keepCopy } `
                    -ReverseType 'ClearForwarding' `
                    -ReverseDescription ("Clear forwarding on {0}" -f $upn) `
                    -Action { Set-Mailbox -Identity $upn -ForwardingSmtpAddress "smtp:$fwd" -DeliverToMailboxAndForward $keepCopy -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Forwarding set." }
                $summary['Forwarding'] = if ($ok) { "$upn -> $fwd" } else { 'failed' }
            }
        }
    }

    # ---- Step 3: Remove from security groups ----
    Write-SectionHeader "Step 3 - Remove from Security Groups"
    if (Confirm-Action "Remove $upn from all pure security groups?") {
        $r = Invoke-OffboardRemoveSecurityGroups -UPN $upn -UserId $user.Id
        Write-Success ("Security groups removed: {0} (failed: {1})" -f $r.Removed, $r.Failed)
        $summary['Security groups removed'] = $r.Removed
    }

    # ---- Step 4: Remove from distribution lists ----
    Write-SectionHeader "Step 4 - Remove from Distribution Lists"
    if (Confirm-Action "Remove $upn from all DLs / non-team M365 groups?") {
        $r = Invoke-OffboardRemoveDistributionLists -UPN $upn -UserId $user.Id
        Write-Success ("DL / group memberships removed: {0} (failed: {1})" -f $r.Removed, $r.Failed)
        $summary['DLs / groups removed'] = $r.Removed
    }

    # ---- Step 5: Remove direct license assignments ----
    Write-SectionHeader "Step 5 - Remove Licenses"
    try {
        $lics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop)
        if ($lics.Count -eq 0) { Write-InfoMsg "No licenses." }
        else {
            $fullUser = $null
            try { $fullUser = Get-MgUser -UserId $user.Id -Property "LicenseAssignmentStates" -ErrorAction Stop } catch {}
            $assignInfo = @{}
            if ($fullUser -and $fullUser.LicenseAssignmentStates) {
                foreach ($s in $fullUser.LicenseAssignmentStates) {
                    $sid = "$($s.SkuId)"; $ai = @{ Direct = $false; Groups = @() }
                    if ($assignInfo.ContainsKey($sid)) { $ai = $assignInfo[$sid] }
                    if ($null -eq $s.AssignedByGroup -or $s.AssignedByGroup -eq "") { $ai.Direct = $true } else { $ai.Groups += $s.AssignedByGroup }
                    $assignInfo[$sid] = $ai
                }
            }
            Write-InfoMsg "Current licenses:"
            foreach ($lic in $lics) {
                $sid = "$($lic.SkuId)"
                $tag = if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) { " [GROUP]" } else { "" }
                Write-Host "    - $(Format-LicenseLabel $lic.SkuPartNumber)$tag" -ForegroundColor White
            }
            if (Confirm-Action "Remove all directly-assigned licenses?") {
                $removed = 0
                foreach ($lic in $lics) {
                    $sid = "$($lic.SkuId)"
                    $friendly = Get-SkuFriendlyName $lic.SkuPartNumber
                    if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) {
                        Write-Warn "$friendly is group-assigned. Remove user from the licensing group instead."
                        continue
                    }
                    $ok = Invoke-Action `
                        -Description ("Remove license '{0}' from {1}" -f $lic.SkuPartNumber, $upn) `
                        -ActionType 'RemoveLicense' `
                        -Target @{ userId = [string]$user.Id; userUpn = $upn; skuId = [string]$lic.SkuId; skuPart = [string]$lic.SkuPartNumber } `
                        -ReverseType 'AssignLicense' `
                        -ReverseDescription ("Re-assign license '{0}' to {1}" -f $lic.SkuPartNumber, $upn) `
                        -Action { Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; $true }
                    if ($ok) { $removed++; if (-not (Get-PreviewMode)) { Write-Success "Removed: $friendly" } }
                }
                $summary['Licenses removed'] = $removed
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }

    # ---- Step 6: Convert mailbox to shared (conditional) ----
    Write-SectionHeader "Step 6 - Convert to Shared Mailbox"
    if (Confirm-Action "Convert $upn to Shared Mailbox?") {
        $ok = Invoke-Action `
            -Description ("Convert {0} to Shared Mailbox" -f $upn) `
            -ActionType 'ConvertToShared' -Target @{ identity = $upn } `
            -NoUndoReason 'Reverting Shared back to UserMailbox requires re-licensing.' `
            -Action { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop; $true }
        if ($ok -and -not (Get-PreviewMode)) { Write-Success "Converted."; $summary['Converted to shared'] = 'yes' }
    }

    # ---- Step 6b: Grant mailbox access (delegates) ----
    Write-SectionHeader "Step 6b - Grant Mailbox Access"
    $ga = Show-Menu -Title "Anyone need access to the converted mailbox?" -Options @("Yes","No") -BackLabel "Skip"
    if ($ga -eq 0) {
        $delegated = 0
        $adding = $true
        while ($adding) {
            $ai = Read-UserInput "User name or email to grant access (blank to stop)"
            if ([string]::IsNullOrWhiteSpace($ai)) { break }
            try {
                $au = if ($ai -match '@') { Get-MgUser -UserId $ai -ErrorAction Stop } else {
                    $found = @(Get-MgUser -Search "displayName:$ai" -ConsistencyLevel eventual -ErrorAction Stop)
                    if ($found.Count -eq 0) { Write-ErrorMsg "Not found."; continue }
                    if ($found.Count -eq 1) { $found[0] } else {
                        $sel = Show-Menu -Title "Select" -Options ($found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                        if ($sel -eq -1) { continue }; $found[$sel]
                    }
                }
                if (Confirm-Action "Grant Full Access to $($au.DisplayName)?") {
                    $ok = Invoke-Action `
                        -Description ("Grant {0} FullAccess on {1}" -f $au.UserPrincipalName, $upn) `
                        -ActionType 'GrantMailboxFullAccess' `
                        -Target @{ mailbox = $upn; user = $au.UserPrincipalName } `
                        -ReverseType 'RevokeMailboxFullAccess' `
                        -ReverseDescription ("Revoke {0} FullAccess on {1}" -f $au.UserPrincipalName, $upn) `
                        -Action { Add-MailboxPermission -Identity $upn -User $au.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; $true }
                    if ($ok) { $delegated++; if (-not (Get-PreviewMode)) { Write-Success "Full Access granted." } }
                }
            } catch { Write-ErrorMsg "Error: $_" }
            $more = Read-UserInput "Grant access to another user? (y/n)"
            if ($more -notmatch '^[Yy]') { $adding = $false }
        }
        if ($delegated -gt 0) { $summary['Mailbox delegates'] = $delegated }
    }

    # ---- Step 7: Teams ownership transfer ----
    if (Get-Command Invoke-TeamsOffboardTransfer -ErrorAction SilentlyContinue) {
        Write-SectionHeader "Step 7 - Teams Ownership Transfer"
        if (Confirm-Action "Transfer / clean up Teams memberships for $upn?") {
            $teamsSuccessor = Read-UserInput "Default successor for sole-owner teams (UPN, blank = prompt per team)"
            $t = Invoke-TeamsOffboardTransfer -LeaverUPN $upn -TeamsSuccessorUPN $teamsSuccessor.Trim()
            if ($t) {
                Write-StatusLine "Sole-owner transferred" "$($t.SoleOwnerActions)" 'White'
                Write-StatusLine "Co-owner demoted"       "$($t.CoOwnerActions)"   'White'
                Write-StatusLine "Member-only removed"    "$($t.MemberRemovals)"   'White'
                if ($t.Failures -gt 0) { Write-Warn "$($t.Failures) team operation(s) failed." }
                $summary['Teams handed off'] = "$($t.SoleOwnerActions) sole + $($t.CoOwnerActions) co + $($t.MemberRemovals) member"
            }
        }
    }

    # ---- Step 8: Revoke outbound SharePoint shares ----
    if (Get-Command Invoke-SharePointOffboardCleanup -ErrorAction SilentlyContinue) {
        Write-SectionHeader "Step 8 - Revoke Outbound SharePoint Shares"
        if (Confirm-Action "Scan and revoke outbound shares created by $upn?") {
            $s = Invoke-SharePointOffboardCleanup -LeaverUPN $upn -LookbackDays 365
            if ($s) {
                Write-StatusLine "Shares found"  "$($s.ShareCount)" 'White'
                Write-StatusLine "Revoked"       "$($s.Revoked)"    'Green'
                Write-StatusLine "Skipped"       "$($s.Skipped)"    'Yellow'
                if ($s.Failed -gt 0) { Write-Warn "$($s.Failed) revocation(s) failed." }
                $summary['Shares revoked'] = "$($s.Revoked)/$($s.ShareCount)"
            }
        }
    }

    # ---- Step 9: OneDrive handoff ----
    $managerEmail = $null
    if (Get-Command Invoke-OneDriveHandoff -ErrorAction SilentlyContinue) {
        Write-SectionHeader "Step 9 - OneDrive Handoff"
        if (Confirm-Action "Hand off $upn's OneDrive to a successor?") {
            $successor = Read-UserInput "Successor UPN (typically the manager)"
            if (-not [string]::IsNullOrWhiteSpace($successor)) {
                $retention = 60
                $rt = Read-UserInput "Retention extension days (default 60)"
                if ($rt -and [int]::TryParse($rt, [ref]$null)) { $retention = [int]$rt }
                $notify = Confirm-Action "Email the successor a OneDrive handoff summary?"
                $hand = Invoke-OneDriveHandoff -LeaverUPN $upn -SuccessorUPN $successor.Trim() -RetentionDays $retention -NotifyManager:$notify
                if ($hand.SiteUrl) {
                    Write-Success "OneDrive: $($hand.SiteUrl)"
                    $summary['OneDrive successor'] = $successor.Trim()
                    $summary['OneDrive retention end (UTC)'] = (Get-Date).ToUniversalTime().AddDays($retention).ToString('yyyy-MM-dd')
                    $managerEmail = $successor.Trim()
                } elseif ($hand.Note) { Write-Warn $hand.Note }
            }
        }
    }

    # ---- Step 10: Manager summary email ----
    if (Get-Command Send-OffboardManagerSummary -ErrorAction SilentlyContinue) {
        Write-SectionHeader "Step 10 - Manager Summary Email"
        $sendTo = Read-UserInput ("Send offboard summary to (UPN; Enter for '{0}', blank to skip)" -f $managerEmail)
        if (-not $sendTo) { $sendTo = $managerEmail }
        if ($sendTo -and (Confirm-Action "Send summary email to $sendTo ?")) {
            Send-OffboardManagerSummary -ManagerUPN $sendTo -LeaverUPN $upn -Summary ([hashtable]$summary) | Out-Null
            Write-Success "Summary emailed to $sendTo."
        }
    }

    # ---- Step 11: Final audit summary line ----
    $detail = ($summary.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ; '
    Write-AuditEntry -EventType 'OFFBOARD_COMPLETE' -Detail "Offboard summary for $upn :: $detail" -ActionType 'OffboardComplete' -Target ([hashtable]$summary) -Result 'info' | Out-Null

    Write-Success "Offboarding complete for $($user.DisplayName)!"
    Pause-ForUser
}
