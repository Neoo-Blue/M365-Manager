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
        try { Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop; Write-Success "Sessions revoked." }
        catch { Write-ErrorMsg "Failed: $_" }
    }

    # Step 2: Block sign-in
    Write-SectionHeader "Step 2 - Block Sign-In"
    if (Confirm-Action "Block sign-in for $upn?") {
        try { Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; Write-Success "Sign-in blocked." }
        catch { Write-ErrorMsg "Failed: $_" }
    }

    # Step 3: OOO
    Write-SectionHeader "Step 3 - Out-of-Office"
    if ([string]::IsNullOrWhiteSpace($userData["OOOMessage"])) { $userData["OOOMessage"] = Read-UserInput "Enter OOO message (or 'skip')" }
    if ($userData["OOOMessage"] -ne 'skip' -and $userData["OOOMessage"]) {
        if (Confirm-Action "Set auto-reply?" "Message: $($userData['OOOMessage'])") {
            try {
                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $userData["OOOMessage"] -ExternalMessage $userData["OOOMessage"] -ErrorAction Stop
                Write-Success "OOO set."
            } catch { Write-ErrorMsg "Failed: $_" }
        }
    }

    # Step 4: Forwarding
    Write-SectionHeader "Step 4 - Email Forwarding"
    if ([string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) { $userData["ForwardingEmail"] = Read-UserInput "Forwarding email (or 'skip')" }
    if ($userData["ForwardingEmail"] -ne 'skip' -and $userData["ForwardingEmail"]) {
        $fwdEmail = $userData["ForwardingEmail"]
        try { $fwdUser = Get-MgUser -UserId $fwdEmail -ErrorAction Stop; Write-Success "Target found: $($fwdUser.DisplayName)" }
        catch { Write-Warn "'$fwdEmail' not found in tenant."; if (-not (Confirm-Action "Use anyway?")) { $fwdEmail = $null } }

        if ($fwdEmail) {
            $dc = Show-Menu -Title "Delivery option" -Options @("Forward only","Forward and keep copy") -BackLabel "Skip"
            if ($dc -ne -1) {
                $keepCopy = ($dc -eq 1)
                if (Confirm-Action "Set forwarding to $fwdEmail (keep copy: $keepCopy)?") {
                    try { Set-Mailbox -Identity $upn -ForwardingSmtpAddress "smtp:$fwdEmail" -DeliverToMailboxAndForward $keepCopy -ErrorAction Stop; Write-Success "Forwarding set." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        }
    }

    # Step 5: Shared mailbox
    Write-SectionHeader "Step 5 - Convert to Shared Mailbox"
    if (Confirm-Action "Convert $upn to Shared Mailbox?") {
        try { Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop; Write-Success "Converted." }
        catch { Write-ErrorMsg "Failed: $_" }
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
                    try { Add-MailboxPermission -Identity $upn -User $au.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; Write-Success "Full Access granted." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
                $pc = Show-Menu -Title "Send permissions for $($au.DisplayName)?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
                if ($pc -ne -1 -and $pc -ne 3) {
                    if ($pc -eq 0 -or $pc -eq 2) { if (Confirm-Action "Grant Send As?") { try { Add-RecipientPermission -Identity $upn -Trustee $au.UserPrincipalName -AccessRights SendAs -Confirm:$false -ErrorAction Stop; Write-Success "Send As granted." } catch { Write-ErrorMsg "$_" } } }
                    if ($pc -eq 1 -or $pc -eq 2) { if (Confirm-Action "Grant Send on Behalf?") { try { Set-Mailbox -Identity $upn -GrantSendOnBehalfTo @{Add = $au.UserPrincipalName} -ErrorAction Stop; Write-Success "Send on Behalf granted." } catch { Write-ErrorMsg "$_" } } }
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
                    try { Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; Write-Success "Removed: $friendly" }
                    catch { Write-ErrorMsg "Failed to remove $friendly : $_" }
                }
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }

    Write-Success "Offboarding complete for $($user.DisplayName)!"
    Pause-ForUser
}
