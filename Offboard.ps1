# ============================================================
#  Offboard.ps1 - User Offboarding Workflow
# ============================================================

function Start-Offboard {
    Write-SectionHeader "User Offboarding"

    # ---- Step 0: Gather user info ----
    $fields = @("UserPrincipalName","ForwardingEmail","OOOMessage")

    $choice = Show-Menu -Title "How to provide offboarding data?" -Options @(
        "Parse from a text file",
        "Enter manually"
    ) -BackLabel "Cancel"

    if ($choice -eq -1) { return }

    $userData = @{}

    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) {
            Write-ErrorMsg "File not found: $filePath"
            Pause-ForUser; return
        }

        Write-InfoMsg "Parsing file..."
        $lines = Get-Content $filePath
        foreach ($line in $lines) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key   = $Matches[1].Trim()
                $value = $Matches[2].Trim()
                switch -Regex ($key) {
                    'upn|email|user'          { $userData["UserPrincipalName"] = $value }
                    'forward'                 { $userData["ForwardingEmail"]   = $value }
                    'ooo|out.of.office|auto.reply' { $userData["OOOMessage"]  = $value }
                    default                   { $userData[$key]               = $value }
                }
            }
        }
        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }

        Write-Success "Parsed data:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    else {
        $userData["UserPrincipalName"] = Read-UserInput "Enter User Principal Name (email)"
        $userData["ForwardingEmail"]   = ""
        $userData["OOOMessage"]        = ""
    }

    $upn = $userData["UserPrincipalName"]
    if ([string]::IsNullOrWhiteSpace($upn)) {
        Write-ErrorMsg "UPN is required."
        Pause-ForUser; return
    }

    # ---- Connect services ----
    if (-not (Connect-ForTask "Offboard")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # Verify user exists
    try {
        $aadUser  = Get-AzureADUser -ObjectId $upn -ErrorAction Stop
        $msolUser = Get-MsolUser -UserPrincipalName $upn -ErrorAction Stop
        Write-Success "User found: $($aadUser.DisplayName) ($upn)"
    } catch {
        Write-ErrorMsg "User not found: $_"
        Pause-ForUser; return
    }

    if (-not (Confirm-Action "Begin offboarding for $($aadUser.DisplayName) ($upn)?")) {
        Write-Warn "Offboarding cancelled."
        Pause-ForUser; return
    }

    # ---- Step 1: Revoke all sessions ----
    Write-SectionHeader "Step 1 - Revoke All Sessions"
    if (Confirm-Action "Revoke all active sessions for $upn?") {
        try {
            Revoke-AzureADUserAllRefreshToken -ObjectId $aadUser.ObjectId -ErrorAction Stop
            Write-Success "All sessions revoked."
        } catch {
            Write-ErrorMsg "Failed to revoke sessions: $_"
        }
    }

    # ---- Step 2: Block sign-in ----
    Write-SectionHeader "Step 2 - Block Sign-In"
    if (Confirm-Action "Block sign-in for $upn?") {
        try {
            Set-AzureADUser -ObjectId $aadUser.ObjectId -AccountEnabled $false -ErrorAction Stop
            Write-Success "Sign-in blocked."
        } catch {
            Write-ErrorMsg "Failed to block sign-in: $_"
        }
    }

    # ---- Step 3: Set Out-of-Office ----
    Write-SectionHeader "Step 3 - Out-of-Office Auto-Reply"

    if ([string]::IsNullOrWhiteSpace($userData["OOOMessage"])) {
        $userData["OOOMessage"] = Read-UserInput "Enter OOO message (or 'skip')"
    }

    if ($userData["OOOMessage"] -ne 'skip' -and -not [string]::IsNullOrWhiteSpace($userData["OOOMessage"])) {
        $oooDetails = "Message: $($userData['OOOMessage'])"
        if (Confirm-Action "Set auto-reply for $upn?" $oooDetails) {
            try {
                Set-MailboxAutoReplyConfiguration -Identity $upn `
                    -AutoReplyState Enabled `
                    -InternalMessage $userData["OOOMessage"] `
                    -ExternalMessage $userData["OOOMessage"] `
                    -ErrorAction Stop
                Write-Success "Out-of-office message set."
            } catch {
                Write-ErrorMsg "Failed to set OOO: $_"
            }
        }
    } else {
        Write-InfoMsg "Skipping OOO setup."
    }

    # ---- Step 4: Set up auto-forwarding ----
    Write-SectionHeader "Step 4 - Email Forwarding"

    if ([string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) {
        $userData["ForwardingEmail"] = Read-UserInput "Enter forwarding email (or 'skip')"
    }

    if ($userData["ForwardingEmail"] -ne 'skip' -and -not [string]::IsNullOrWhiteSpace($userData["ForwardingEmail"])) {
        $fwdEmail = $userData["ForwardingEmail"]

        # Validate forwarding target exists in tenant
        Write-InfoMsg "Validating forwarding target in tenant..."
        try {
            $fwdUser = Get-AzureADUser -ObjectId $fwdEmail -ErrorAction Stop
            Write-Success "Forwarding target found: $($fwdUser.DisplayName) ($fwdEmail)"
        } catch {
            Write-Warn "Forwarding email '$fwdEmail' was NOT found in the tenant."
            if (-not (Confirm-Action "Proceed with this external/unknown forwarding address anyway?")) {
                Write-InfoMsg "Skipping forwarding."
                $fwdEmail = $null
            }
        }

        if ($fwdEmail) {
            # Ask about delivery options
            $deliveryChoice = Show-Menu -Title "Forwarding delivery option" -Options @(
                "Forward only (no copy in original mailbox)",
                "Forward AND keep a copy in original mailbox"
            ) -BackLabel "Skip forwarding"

            if ($deliveryChoice -ne -1) {
                $keepCopy = ($deliveryChoice -eq 1)
                $fwdDetails = "Forward to: $fwdEmail`nKeep copy: $keepCopy"
                if (Confirm-Action "Set up mail forwarding?" $fwdDetails) {
                    try {
                        Set-Mailbox -Identity $upn `
                            -ForwardingSmtpAddress "smtp:$fwdEmail" `
                            -DeliverToMailboxAndForward $keepCopy `
                            -ErrorAction Stop
                        Write-Success "Forwarding configured to $fwdEmail (Keep copy: $keepCopy)."
                    } catch {
                        Write-ErrorMsg "Failed to set forwarding: $_"
                    }
                }
            }
        }
    } else {
        Write-InfoMsg "Skipping email forwarding."
    }

    # ---- Step 5: Convert to shared mailbox ----
    Write-SectionHeader "Step 5 - Convert to Shared Mailbox"
    if (Confirm-Action "Convert $upn to a Shared Mailbox?") {
        try {
            Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
            Write-Success "Mailbox converted to Shared."
        } catch {
            Write-ErrorMsg "Failed to convert mailbox: $_"
        }
    }

    # ---- Step 6: Grant mailbox access ----
    Write-SectionHeader "Step 6 - Grant Mailbox Access"

    $grantAccess = Show-Menu -Title "Does anyone need access to this mailbox?" -Options @(
        "Yes, grant access to one or more users",
        "No, skip"
    ) -BackLabel "Skip"

    if ($grantAccess -eq 0) {
        $addingUsers = $true
        while ($addingUsers) {
            $accessInput = Read-UserInput "Enter the name or email of the person to grant access"
            if ([string]::IsNullOrWhiteSpace($accessInput)) { break }

            try {
                $accessUser = $null
                if ($accessInput -match '@') {
                    $accessUser = Get-AzureADUser -ObjectId $accessInput -ErrorAction Stop
                } else {
                    $found = Get-AzureADUser -SearchString $accessInput -ErrorAction Stop
                    if ($found.Count -eq 0) {
                        Write-ErrorMsg "No user found matching '$accessInput'."
                        continue
                    }
                    if ($found.Count -eq 1) {
                        $accessUser = $found[0]
                    } else {
                        $names = $found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                        $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                        if ($sel -eq -1) { continue }
                        $accessUser = $found[$sel]
                    }
                }

                Write-StatusLine "Granting to" "$($accessUser.DisplayName) ($($accessUser.UserPrincipalName))" "White"

                # Full Access
                if (Confirm-Action "Grant Full Access to $($accessUser.DisplayName) on $upn mailbox?") {
                    try {
                        Add-MailboxPermission -Identity $upn `
                            -User $accessUser.UserPrincipalName `
                            -AccessRights FullAccess `
                            -InheritanceType All -AutoMapping $true `
                            -ErrorAction Stop
                        Write-Success "Full Access granted to $($accessUser.DisplayName)."
                    } catch {
                        Write-ErrorMsg "Failed to grant Full Access: $_"
                    }
                }

                # Send As / Send on Behalf
                $permChoice = Show-Menu -Title "Grant additional send permissions to $($accessUser.DisplayName)?" -Options @(
                    "Grant Send As",
                    "Grant Send on Behalf",
                    "Both Send As and Send on Behalf",
                    "No additional send permissions"
                ) -BackLabel "Skip"

                if ($permChoice -ne -1 -and $permChoice -ne 3) {
                    if ($permChoice -eq 0 -or $permChoice -eq 2) {
                        if (Confirm-Action "Grant Send As on $upn to $($accessUser.DisplayName)?") {
                            try {
                                Add-RecipientPermission -Identity $upn `
                                    -Trustee $accessUser.UserPrincipalName `
                                    -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                                Write-Success "Send As permission granted to $($accessUser.DisplayName)."
                            } catch {
                                Write-ErrorMsg "Send As failed: $_"
                            }
                        }
                    }
                    if ($permChoice -eq 1 -or $permChoice -eq 2) {
                        if (Confirm-Action "Grant Send on Behalf on $upn to $($accessUser.DisplayName)?") {
                            try {
                                Set-Mailbox -Identity $upn `
                                    -GrantSendOnBehalfTo @{Add = $accessUser.UserPrincipalName} `
                                    -ErrorAction Stop
                                Write-Success "Send on Behalf permission granted to $($accessUser.DisplayName)."
                            } catch {
                                Write-ErrorMsg "Send on Behalf failed: $_"
                            }
                        }
                    }
                }
            } catch {
                Write-ErrorMsg "Error finding user: $_"
            }

            # Ask if they want to add another user
            $more = Read-UserInput "Grant access to another user? (y/n)"
            if ($more -notmatch '^[Yy]') { $addingUsers = $false }
        }
    }

    # ---- Step 7: Remove licenses ----
    Write-SectionHeader "Step 7 - Remove Licenses"

    try {
        $currentLicenses = $msolUser.Licenses
        if ($currentLicenses.Count -eq 0) {
            Write-InfoMsg "User has no licenses assigned."
        } else {
            $licLabels = $currentLicenses | ForEach-Object { $_.AccountSkuId }
            Write-InfoMsg "Current licenses:"
            $licLabels | ForEach-Object { Write-Host "    - $_" -ForegroundColor White }
            Write-Host ""

            if (Confirm-Action "Remove ALL licenses from $upn?") {
                foreach ($lic in $currentLicenses) {
                    try {
                        Set-MsolUserLicense -UserPrincipalName $upn `
                            -RemoveLicenses $lic.AccountSkuId -ErrorAction Stop
                        Write-Success "Removed: $($lic.AccountSkuId)"
                    } catch {
                        Write-ErrorMsg "Failed to remove $($lic.AccountSkuId): $_"
                    }
                }
            }
        }
    } catch {
        Write-ErrorMsg "Error retrieving licenses: $_"
    }

    Write-Host ""
    Write-Success "Offboarding complete for $($aadUser.DisplayName)!"
    Pause-ForUser
}
