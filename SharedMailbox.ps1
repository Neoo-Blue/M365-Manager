# ============================================================
#  SharedMailbox.ps1 - Shared Mailbox Management
# ============================================================

function Start-SharedMailboxManagement {
    Write-SectionHeader "Shared Mailbox Management"

    if (-not (Connect-ForTask "SharedMailbox")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new shared mailbox",
        "Add / remove user access",
        "View / edit mailbox properties",
        "Delete a shared mailbox"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-SharedMailboxFlow }
        1 { Edit-SharedMailboxAccess }
        2 { Edit-SharedMailboxProperties }
        3 { Remove-SharedMailboxFlow }
        -1 { return }
    }
}

# ==================================================================
#  Create
# ==================================================================
function New-SharedMailboxFlow {
    Write-SectionHeader "Create New Shared Mailbox"

    $name  = Read-UserInput "Display name for the shared mailbox"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-ErrorMsg "Name is required."; Pause-ForUser; return }

    $email = Read-UserInput "Email address (e.g. info@contoso.com)"
    if ([string]::IsNullOrWhiteSpace($email)) { Write-ErrorMsg "Email is required."; Pause-ForUser; return }

    $alias = Read-UserInput "Alias (or press Enter to auto-generate)"
    if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($email -split '@')[0] }

    $details = "Name  : $name`nEmail : $email`nAlias : $alias"
    if (-not (Confirm-Action "Create this shared mailbox?" $details)) { Pause-ForUser; return }

    try {
        New-Mailbox -Name $name -DisplayName $name -Alias $alias `
            -PrimarySmtpAddress $email -Shared -ErrorAction Stop | Out-Null
        Write-Success "Shared mailbox '$name' ($email) created."

        $addNow = Read-UserInput "Grant access to users now? (y/n)"
        if ($addNow -match '^[Yy]') {
            Add-SharedMailboxAccessLoop -MailboxIdentity $email -MailboxName $name
        }
    } catch {
        Write-ErrorMsg "Failed to create shared mailbox: $_"
    }
    Pause-ForUser
}

# ==================================================================
#  Access (add / remove)
# ==================================================================
function Edit-SharedMailboxAccess {
    Write-SectionHeader "Manage Shared Mailbox Access"

    $box = Find-SharedMailbox
    if ($null -eq $box) { Pause-ForUser; return }

    # Show current permissions
    Show-MailboxPermissions -MailboxIdentity $box.PrimarySmtpAddress -MailboxName $box.DisplayName

    $action = Show-Menu -Title "Action for '$($box.DisplayName)'" -Options @(
        "Grant access to a user",
        "Remove user access"
    ) -BackLabel "Done"

    if ($action -eq 0) {
        Add-SharedMailboxAccessLoop -MailboxIdentity $box.PrimarySmtpAddress -MailboxName $box.DisplayName
    }
    elseif ($action -eq 1) {
        # ---- REMOVE ----
        try {
            $perms = @(Get-MailboxPermission -Identity $box.PrimarySmtpAddress -ErrorAction Stop |
                Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and $_.IsInherited -eq $false })

            if ($perms.Count -eq 0) {
                Write-InfoMsg "No custom user permissions found."
                Pause-ForUser; return
            }

            $labels = $perms | ForEach-Object { "$($_.User)  ($($_.AccessRights -join ', '))" }
            $selected = Show-MultiSelect -Title "Select permission(s) to remove" -Options $labels `
                -Prompt "Enter number(s) (e.g. 1,3)"

            foreach ($idx in $selected) {
                $perm = $perms[$idx]
                if (Confirm-Action "Remove all permissions for '$($perm.User)' on '$($box.DisplayName)'?") {
                    try {
                        Remove-MailboxPermission -Identity $box.PrimarySmtpAddress `
                            -User $perm.User -AccessRights FullAccess `
                            -InheritanceType All -Confirm:$false -ErrorAction Stop
                        Write-Success "Full Access removed for '$($perm.User)'."
                    } catch { Write-ErrorMsg "Failed to remove Full Access: $_" }

                    try {
                        Remove-RecipientPermission -Identity $box.PrimarySmtpAddress `
                            -Trustee $perm.User -AccessRights SendAs `
                            -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Success "Send As removed (if existed)."
                    } catch {}

                    try {
                        Set-Mailbox -Identity $box.PrimarySmtpAddress `
                            -GrantSendOnBehalfTo @{Remove = $perm.User} -ErrorAction SilentlyContinue
                        Write-Success "Send on Behalf removed (if existed)."
                    } catch {}
                }
            }
        } catch { Write-ErrorMsg "Error: $_" }
    }
    Pause-ForUser
}

# ==================================================================
#  Properties
# ==================================================================
function Edit-SharedMailboxProperties {
    Write-SectionHeader "View / Edit Shared Mailbox Properties"

    $box = Find-SharedMailbox
    if ($null -eq $box) { Pause-ForUser; return }

    # Refresh
    try { $box = Get-Mailbox -Identity $box.PrimarySmtpAddress -ErrorAction Stop } catch {}

    Write-StatusLine "Display Name"     $box.DisplayName "White"
    Write-StatusLine "Primary Email"    $box.PrimarySmtpAddress "White"
    Write-StatusLine "Alias"            $box.Alias "White"
    Write-StatusLine "Mailbox Type"     "$($box.RecipientTypeDetails)" "Cyan"
    Write-StatusLine "Hidden from GAL"  "$($box.HiddenFromAddressListsEnabled)" "White"
    Write-StatusLine "Forwarding"       $(if ($box.ForwardingSmtpAddress) { $box.ForwardingSmtpAddress } else { "(none)" }) "White"
    Write-StatusLine "Deliver + Fwd"    "$($box.DeliverToMailboxAndForward)" "White"
    Write-StatusLine "Archive Status"   "$($box.ArchiveStatus)" "White"

    # Email aliases
    $aliases = $box.EmailAddresses | Where-Object { $_ -like "smtp:*" -and $_ -ne "SMTP:$($box.PrimarySmtpAddress)" }
    if ($aliases.Count -gt 0) {
        Write-StatusLine "Aliases" ($aliases -join "; ") "White"
    }

    # Auto-reply
    try {
        $autoReply = Get-MailboxAutoReplyConfiguration -Identity $box.PrimarySmtpAddress -ErrorAction Stop
        Write-StatusLine "Auto-Reply" "$($autoReply.AutoReplyState)" "White"
    } catch {}

    # Send on Behalf
    if ($box.GrantSendOnBehalfTo.Count -gt 0) {
        Write-StatusLine "Send on Behalf" ($box.GrantSendOnBehalfTo -join "; ") "White"
    }
    Write-Host ""

    $editChoice = Show-Menu -Title "Edit Properties" -Options @(
        "Change display name",
        "Add an email alias",
        "Remove an email alias",
        "Set up email forwarding",
        "Remove email forwarding",
        "Toggle hidden from address book",
        "Set auto-reply message"
    ) -BackLabel "Done"

    switch ($editChoice) {
        0 {
            $newVal = Read-UserInput "New display name"
            if ($newVal -and (Confirm-Action "Rename to '$newVal'?")) {
                try {
                    Set-Mailbox -Identity $box.PrimarySmtpAddress -DisplayName $newVal -ErrorAction Stop
                    Write-Success "Display name updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        1 {
            $newAlias = Read-UserInput "New email alias (full address, e.g. alt@contoso.com)"
            if ($newAlias -and (Confirm-Action "Add alias '$newAlias'?")) {
                try {
                    Set-Mailbox -Identity $box.PrimarySmtpAddress `
                        -EmailAddresses @{Add = "smtp:$newAlias"} -ErrorAction Stop
                    Write-Success "Alias '$newAlias' added."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        2 {
            $currentAliases = @($box.EmailAddresses | Where-Object { $_ -like "smtp:*" })
            if ($currentAliases.Count -eq 0) {
                Write-InfoMsg "No aliases to remove."
            } else {
                $aliasLabels = $currentAliases | ForEach-Object { $_ -replace '^smtp:','' }
                $sel = Show-MultiSelect -Title "Select alias(es) to remove" -Options $aliasLabels
                foreach ($i in $sel) {
                    $removeAddr = $currentAliases[$i]
                    if (Confirm-Action "Remove alias '$($aliasLabels[$i])'?") {
                        try {
                            Set-Mailbox -Identity $box.PrimarySmtpAddress `
                                -EmailAddresses @{Remove = $removeAddr} -ErrorAction Stop
                            Write-Success "Alias removed."
                        } catch { Write-ErrorMsg "Failed: $_" }
                    }
                }
            }
        }
        3 {
            $fwdAddr = Read-UserInput "Forward to email address"
            $keepCopy = Show-Menu -Title "Keep a copy in the shared mailbox?" -Options @("Yes","No") -BackLabel "Cancel"
            if ($keepCopy -ne -1 -and $fwdAddr) {
                $deliver = ($keepCopy -eq 0)
                if (Confirm-Action "Set forwarding to '$fwdAddr' (keep copy: $deliver)?") {
                    try {
                        Set-Mailbox -Identity $box.PrimarySmtpAddress `
                            -ForwardingSmtpAddress "smtp:$fwdAddr" `
                            -DeliverToMailboxAndForward $deliver -ErrorAction Stop
                        Write-Success "Forwarding configured."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        }
        4 {
            if (Confirm-Action "Remove all forwarding from '$($box.DisplayName)'?") {
                try {
                    Set-Mailbox -Identity $box.PrimarySmtpAddress `
                        -ForwardingSmtpAddress $null `
                        -DeliverToMailboxAndForward $false -ErrorAction Stop
                    Write-Success "Forwarding removed."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        5 {
            $current = $box.HiddenFromAddressListsEnabled
            $newVal = -not $current
            $label = if ($newVal) { "Hidden" } else { "Visible" }
            if (Confirm-Action "Set address book visibility to: $label ?") {
                try {
                    Set-Mailbox -Identity $box.PrimarySmtpAddress `
                        -HiddenFromAddressListsEnabled $newVal -ErrorAction Stop
                    Write-Success "Visibility set to: $label"
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        6 {
            $internalMsg = Read-UserInput "Internal auto-reply message"
            $externalMsg = Read-UserInput "External auto-reply message (or press Enter to use same)"
            if ([string]::IsNullOrWhiteSpace($externalMsg)) { $externalMsg = $internalMsg }
            if ($internalMsg -and (Confirm-Action "Set auto-reply on '$($box.DisplayName)'?")) {
                try {
                    Set-MailboxAutoReplyConfiguration -Identity $box.PrimarySmtpAddress `
                        -AutoReplyState Enabled `
                        -InternalMessage $internalMsg `
                        -ExternalMessage $externalMsg -ErrorAction Stop
                    Write-Success "Auto-reply configured."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
    }
    Pause-ForUser
}

# ==================================================================
#  Delete
# ==================================================================
function Remove-SharedMailboxFlow {
    Write-SectionHeader "Delete Shared Mailbox"

    $box = Find-SharedMailbox
    if ($null -eq $box) { Pause-ForUser; return }

    Write-StatusLine "Name"  $box.DisplayName "White"
    Write-StatusLine "Email" $box.PrimarySmtpAddress "White"

    # Show who has access
    Show-MailboxPermissions -MailboxIdentity $box.PrimarySmtpAddress -MailboxName $box.DisplayName

    Write-Host ""
    Write-Warn "This will permanently delete the mailbox and ALL its contents!"

    if (Confirm-Action "DELETE shared mailbox '$($box.DisplayName)'?") {
        $doubleCheck = Read-UserInput "Type the mailbox email address to confirm deletion"
        if ($doubleCheck -eq $box.PrimarySmtpAddress) {
            try {
                Remove-Mailbox -Identity $box.PrimarySmtpAddress -Confirm:$false -ErrorAction Stop
                Write-Success "Shared mailbox '$($box.DisplayName)' deleted."
            } catch { Write-ErrorMsg "Failed to delete: $_" }
        } else {
            Write-Warn "Email did not match. Deletion cancelled."
        }
    }
    Pause-ForUser
}

# ==================================================================
#  Shared helpers
# ==================================================================
function Find-SharedMailbox {
    $searchMethod = Show-Menu -Title "Find shared mailbox by" -Options @(
        "Search by name",
        "Search by email address"
    ) -BackLabel "Cancel"
    if ($searchMethod -eq -1) { return $null }

    $searchInput = Read-UserInput $(if ($searchMethod -eq 0) { "Enter mailbox name (partial match)" } else { "Enter mailbox email address" })
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }

    try {
        if ($searchMethod -eq 0) {
            $boxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox `
                -Filter "DisplayName -like '*$searchInput*'" -ResultSize 50 -ErrorAction Stop)
        } else {
            $boxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox `
                -Filter "PrimarySmtpAddress -like '*$searchInput*'" -ResultSize 50 -ErrorAction Stop)
        }

        if ($boxes.Count -eq 0) {
            Write-ErrorMsg "No shared mailboxes found matching '$searchInput'."
            return $null
        }
        if ($boxes.Count -eq 1) {
            Write-Success "Found: $($boxes[0].DisplayName) ($($boxes[0].PrimarySmtpAddress))"
            return $boxes[0]
        }

        $labels = $boxes | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
        $sel = Show-Menu -Title "Select Shared Mailbox" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return $boxes[$sel]
    } catch {
        Write-ErrorMsg "Search error: $_"
        return $null
    }
}

function Show-MailboxPermissions {
    param([string]$MailboxIdentity, [string]$MailboxName)

    Write-InfoMsg "Current permissions on '$MailboxName':"
    try {
        $perms = Get-MailboxPermission -Identity $MailboxIdentity -ErrorAction Stop |
            Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and $_.IsInherited -eq $false }

        if ($perms.Count -eq 0) {
            Write-InfoMsg "  (no custom permissions)"
        } else {
            foreach ($p in $perms) {
                Write-Host "    - $($p.User)  [$($p.AccessRights -join ', ')]" -ForegroundColor White
            }
        }

        # Send As
        $sendAs = Get-RecipientPermission -Identity $MailboxIdentity -ErrorAction SilentlyContinue |
            Where-Object { $_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -ne "S-1-*" }
        if ($sendAs.Count -gt 0) {
            Write-InfoMsg "Send As permissions:"
            foreach ($sa in $sendAs) {
                Write-Host "    - $($sa.Trustee)  [SendAs]" -ForegroundColor White
            }
        }
    } catch { Write-Warn "Could not read permissions: $_" }
    Write-Host ""
}

function Add-SharedMailboxAccessLoop {
    param([string]$MailboxIdentity, [string]$MailboxName)

    $adding = $true
    while ($adding) {
        $userInput = Read-UserInput "Enter user name or email to grant access (or 'done')"
        if ($userInput -match '^done$') { break }

        try {
            $targetUser = $null
            if ($userInput -match '@') {
                $targetUser = Get-AzureADUser -ObjectId $userInput -ErrorAction Stop
            } else {
                $found = Get-AzureADUser -SearchString $userInput -ErrorAction Stop
                if ($found.Count -eq 0) { Write-ErrorMsg "No user found."; continue }
                if ($found.Count -eq 1) { $targetUser = $found[0] }
                else {
                    $names = $found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                    $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                    if ($sel -eq -1) { continue }
                    $targetUser = $found[$sel]
                }
            }

            $upn = $targetUser.UserPrincipalName

            # Full Access
            if (Confirm-Action "Grant Full Access to '$($targetUser.DisplayName)' on '$MailboxName'?") {
                try {
                    Add-MailboxPermission -Identity $MailboxIdentity `
                        -User $upn -AccessRights FullAccess `
                        -InheritanceType All -AutoMapping $true -ErrorAction Stop
                    Write-Success "Full Access granted."
                } catch { Write-ErrorMsg "Full Access failed: $_" }
            }

            # Send permissions
            $permChoice = Show-Menu -Title "Grant additional send permissions?" -Options @(
                "Grant Send As",
                "Grant Send on Behalf",
                "Both Send As and Send on Behalf",
                "No additional permissions"
            ) -BackLabel "Skip"

            if ($permChoice -ne -1 -and $permChoice -ne 3) {
                if ($permChoice -eq 0 -or $permChoice -eq 2) {
                    if (Confirm-Action "Grant Send As on '$MailboxName' to $upn?") {
                        try {
                            Add-RecipientPermission -Identity $MailboxIdentity `
                                -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                            Write-Success "Send As granted."
                        } catch { Write-ErrorMsg "Send As failed: $_" }
                    }
                }
                if ($permChoice -eq 1 -or $permChoice -eq 2) {
                    if (Confirm-Action "Grant Send on Behalf on '$MailboxName' to $upn?") {
                        try {
                            Set-Mailbox -Identity $MailboxIdentity `
                                -GrantSendOnBehalfTo @{Add = $upn} -ErrorAction Stop
                            Write-Success "Send on Behalf granted."
                        } catch { Write-ErrorMsg "Send on Behalf failed: $_" }
                    }
                }
            }
        } catch { Write-ErrorMsg "Error: $_" }
    }
}
