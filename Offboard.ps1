# ============================================================
#  Offboard.ps1 - User Offboarding (Microsoft Graph)
# ============================================================

function Start-Offboard {
    Write-SectionHeader "User Offboarding"

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
                    'ooo|out.of.office|auto.reply'  { $userData["OOOMessage"]        = $value }
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

    # Step 1: Revoke sessions
    Write-SectionHeader "Step 1 - Revoke All Sessions"
    if (Confirm-Action "Revoke all sessions for $upn?") {
        $ok = Invoke-Action `
            -Description ("Revoke sign-in sessions for {0}" -f $upn) `
            -ActionType 'RevokeSignInSessions' `
            -Target @{ userId = [string]$user.Id; userUpn = $upn } `
            -NoUndoReason 'Sign-in sessions, once revoked, cannot be restored -- the user re-signs in fresh.' `
            -Action { Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop; $true }
        if ($ok -and -not (Get-PreviewMode)) { Write-Success "Sessions revoked." }
    }

    # Step 2: Block sign-in
    Write-SectionHeader "Step 2 - Block Sign-In"
    if (Confirm-Action "Block sign-in for $upn?") {
        $ok = Invoke-Action `
            -Description ("Block sign-in for {0}" -f $upn) `
            -ActionType 'BlockSignIn' `
            -Target @{ userId = [string]$user.Id; userUpn = $upn } `
            -ReverseType 'UnblockSignIn' `
            -ReverseDescription ("Unblock sign-in for {0}" -f $upn) `
            -Action { Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; $true }
        if ($ok -and -not (Get-PreviewMode)) { Write-Success "Sign-in blocked." }
    }

    # Step 3: OOO
    Write-SectionHeader "Step 3 - Out-of-Office"
    if ([string]::IsNullOrWhiteSpace($userData["OOOMessage"])) { $userData["OOOMessage"] = Read-UserInput "Enter OOO message (or 'skip')" }
    if ($userData["OOOMessage"] -ne 'skip' -and $userData["OOOMessage"]) {
        if (Confirm-Action "Set auto-reply?" "Message: $($userData['OOOMessage'])") {
            $ok = Invoke-Action `
                -Description ("Set OOO auto-reply for {0}" -f $upn) `
                -ActionType 'SetOOO' `
                -Target @{ identity = $upn } `
                -ReverseType 'ClearOOO' `
                -ReverseDescription ("Clear OOO auto-reply on {0}" -f $upn) `
                -Action {
                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $userData["OOOMessage"] -ExternalMessage $userData["OOOMessage"] -ErrorAction Stop; $true
                }
            if ($ok -and -not (Get-PreviewMode)) { Write-Success "OOO set." }
        }
    }

    # Step 4: Forwarding
    Write-SectionHeader "Step 4 - Email Forwarding"
    if ([string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) { $userData["ForwardingEmail"] = Read-UserInput "Forwarding email (or 'skip')" }
    if ($userData["ForwardingEmail"] -ne 'skip' -and $userData["ForwardingEmail"]) {
        $fwdEmail = $userData["ForwardingEmail"]
        if (-not (Get-PreviewMode)) {
            try { $fwdUser = Get-MgUser -UserId $fwdEmail -ErrorAction Stop; Write-Success "Target found: $($fwdUser.DisplayName)" }
            catch { Write-Warn "'$fwdEmail' not found in tenant."; if (-not (Confirm-Action "Use anyway?")) { $fwdEmail = $null } }
        }

        if ($fwdEmail) {
            $dc = Show-Menu -Title "Delivery option" -Options @("Forward only","Forward and keep copy") -BackLabel "Skip"
            if ($dc -ne -1) {
                $keepCopy = ($dc -eq 1)
                if (Confirm-Action "Set forwarding to $fwdEmail (keep copy: $keepCopy)?") {
                    $ok = Invoke-Action `
                        -Description ("Set forwarding {0} -> {1} (keep copy: {2})" -f $upn, $fwdEmail, $keepCopy) `
                        -ActionType 'SetForwarding' `
                        -Target @{ identity = $upn; forwardTo = $fwdEmail; keepCopy = $keepCopy } `
                        -ReverseType 'ClearForwarding' `
                        -ReverseDescription ("Clear forwarding on {0}" -f $upn) `
                        -Action {
                            Set-Mailbox -Identity $upn -ForwardingSmtpAddress "smtp:$fwdEmail" -DeliverToMailboxAndForward $keepCopy -ErrorAction Stop; $true
                        }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Forwarding set." }
                }
            }
        }
    }

    # Step 5: Shared mailbox
    Write-SectionHeader "Step 5 - Convert to Shared Mailbox"
    if (Confirm-Action "Convert $upn to Shared Mailbox?") {
        $ok = Invoke-Action `
            -Description ("Convert {0} to Shared Mailbox" -f $upn) `
            -ActionType 'ConvertToShared' `
            -Target @{ identity = $upn } `
            -NoUndoReason 'Reverting Shared back to UserMailbox requires re-licensing and operator judgment; not auto-reversible.' `
            -Action { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop; $true }
        if ($ok -and -not (Get-PreviewMode)) { Write-Success "Converted." }
    }

    # Step 6: Grant access
    Write-SectionHeader "Step 6 - Grant Mailbox Access"
    $ga = Show-Menu -Title "Anyone need access?" -Options @("Yes","No") -BackLabel "Skip"
    if ($ga -eq 0) {
        $adding = $true
        while ($adding) {
            $ai = Read-UserInput "User name or email to grant access"
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
                        -Action {
                            Add-MailboxPermission -Identity $upn -User $au.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; $true
                        }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Full Access granted." }
                }
                $pc = Show-Menu -Title "Send permissions for $($au.DisplayName)?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
                if ($pc -ne -1 -and $pc -ne 3) {
                    if ($pc -eq 0 -or $pc -eq 2) {
                        if (Confirm-Action "Grant Send As?") {
                            $ok = Invoke-Action `
                                -Description ("Grant {0} SendAs on {1}" -f $au.UserPrincipalName, $upn) `
                                -ActionType 'GrantMailboxSendAs' `
                                -Target @{ mailbox = $upn; user = $au.UserPrincipalName } `
                                -ReverseType 'RevokeMailboxSendAs' `
                                -ReverseDescription ("Revoke {0} SendAs on {1}" -f $au.UserPrincipalName, $upn) `
                                -Action {
                                    Add-RecipientPermission -Identity $upn -Trustee $au.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop; $true
                                }
                            if ($ok -and -not (Get-PreviewMode)) { Write-Success "Send As granted." }
                        }
                    }
                    if ($pc -eq 1 -or $pc -eq 2) {
                        if (Confirm-Action "Grant Send on Behalf?") {
                            $ok = Invoke-Action -Description ("Grant {0} SendOnBehalf on {1}" -f $au.UserPrincipalName, $upn) -Action {
                                Set-Mailbox -Identity $upn -GrantSendOnBehalfTo @{Add = $au.UserPrincipalName} -ErrorAction Stop; $true
                            }
                            if ($ok -and -not (Get-PreviewMode)) { Write-Success "Send on Behalf granted." }
                        }
                    }
                }
            } catch { Write-ErrorMsg "Error: $_" }
            $more = Read-UserInput "Grant access to another user? (y/n)"
            if ($more -notmatch '^[Yy]') { $adding = $false }
        }
    }

    # Step 7: Remove licenses
    Write-SectionHeader "Step 7 - Remove Licenses"
    try {
        $lics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop)
        if ($lics.Count -eq 0) { Write-InfoMsg "No licenses." }
        else {
            # Check assignment states
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
                $tag = ""
                $sid = "$($lic.SkuId)"
                if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) { $tag = " [GROUP]" }
                Write-Host "    - $(Format-LicenseLabel $lic.SkuPartNumber)$tag" -ForegroundColor White
            }

            if (Confirm-Action "Remove all directly-assigned licenses?") {
                foreach ($lic in $lics) {
                    $sid = "$($lic.SkuId)"
                    $friendly = Get-SkuFriendlyName $lic.SkuPartNumber
                    if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) {
                        Write-Warn "$friendly is group-assigned. Skipping (remove user from the group instead)."
                        continue
                    }
                    $ok = Invoke-Action `
                        -Description ("Remove license '{0}' from {1}" -f $lic.SkuPartNumber, $upn) `
                        -ActionType 'RemoveLicense' `
                        -Target @{ userId = [string]$user.Id; userUpn = $upn; skuId = [string]$lic.SkuId; skuPart = [string]$lic.SkuPartNumber } `
                        -ReverseType 'AssignLicense' `
                        -ReverseDescription ("Re-assign license '{0}' to {1}" -f $lic.SkuPartNumber, $upn) `
                        -Action {
                            Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; $true
                        }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Removed: $friendly" }
                }
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }

    Write-Success "Offboarding complete for $($user.DisplayName)!"
    Pause-ForUser
}
