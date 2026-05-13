# ============================================================
#  SharedMailbox.ps1 - Shared Mailbox Management (MS Graph + EXO)
# ============================================================

function Start-SharedMailboxManagement {
    Write-SectionHeader "Shared Mailbox Management"
    if (-not (Connect-ForTask "SharedMailbox")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new shared mailbox","Add / remove user access",
        "View / edit mailbox properties","Delete a shared mailbox"
    ) -BackLabel "Back to Main Menu"

    switch ($action) { 0 { New-SharedMailboxFlow } 1 { Edit-SharedMailboxAccess } 2 { Edit-SharedMailboxProperties } 3 { Remove-SharedMailboxFlow } }
}

function New-SharedMailboxFlow {
    Write-SectionHeader "Create New Shared Mailbox"
    $name = Read-UserInput "Display name"; if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $email = Read-UserInput "Email address"; if ([string]::IsNullOrWhiteSpace($email)) { Pause-ForUser; return }
    $alias = Read-UserInput "Alias (or Enter to auto)"; if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($email -split '@')[0] }
    if (Confirm-Action "Create shared mailbox '$name' ($email)?") {
        try {
            New-Mailbox -Name $name -DisplayName $name -Alias $alias -PrimarySmtpAddress $email -Shared -ErrorAction Stop | Out-Null
            Write-Success "Created."
            $add = Read-UserInput "Grant access now? (y/n)"; if ($add -match '^[Yy]') { Add-SharedMailboxAccessLoop -Id $email -Name $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-SharedMailboxAccess {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    Show-MailboxPermissions -Id $box.PrimarySmtpAddress -Name $box.DisplayName
    $action = Show-Menu -Title "Action" -Options @("Grant access","Remove access") -BackLabel "Done"
    if ($action -eq 0) { Add-SharedMailboxAccessLoop -Id $box.PrimarySmtpAddress -Name $box.DisplayName }
    elseif ($action -eq 1) {
        try {
            $perms = @(Get-MailboxPermission -Identity $box.PrimarySmtpAddress -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and -not $_.IsInherited })
            if ($perms.Count -eq 0) { Write-InfoMsg "No custom permissions."; Pause-ForUser; return }
            $labels = $perms | ForEach-Object { "$($_.User)  ($($_.AccessRights -join ', '))" }
            $sel = Show-MultiSelect -Title "Remove" -Options $labels
            foreach ($idx in $sel) {
                $p = $perms[$idx]
                if (Confirm-Action "Remove all permissions for '$($p.User)'?") {
                    $ok1 = Invoke-Action -Description ("Revoke FullAccess for {0} on '{1}'" -f $p.User, $box.PrimarySmtpAddress) -Action {
                        Remove-MailboxPermission -Identity $box.PrimarySmtpAddress -User $p.User -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop; $true
                    }
                    if ($ok1 -and -not (Get-PreviewMode)) { Write-Success "Full Access removed." }
                    $ok2 = Invoke-Action -Description ("Revoke SendAs for {0} on '{1}'" -f $p.User, $box.PrimarySmtpAddress) -Action {
                        Remove-RecipientPermission -Identity $box.PrimarySmtpAddress -Trustee $p.User -AccessRights SendAs -Confirm:$false -ErrorAction Stop; $true
                    }
                    if ($ok2 -and -not (Get-PreviewMode)) { Write-Success "Send As removed." }
                    $ok3 = Invoke-Action -Description ("Revoke SendOnBehalf for {0} on '{1}'" -f $p.User, $box.PrimarySmtpAddress) -Action {
                        Set-Mailbox -Identity $box.PrimarySmtpAddress -GrantSendOnBehalfTo @{Remove=$p.User} -ErrorAction Stop; $true
                    }
                    if ($ok3 -and -not (Get-PreviewMode)) { Write-Success "Send on Behalf removed." }
                }
            }
        } catch { Write-ErrorMsg "$_" }
    }
    Pause-ForUser
}

function Edit-SharedMailboxProperties {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    try { $box = Get-Mailbox -Identity $box.PrimarySmtpAddress -ErrorAction Stop } catch {}
    Write-StatusLine "Name" $box.DisplayName "White"; Write-StatusLine "Email" $box.PrimarySmtpAddress "White"
    Write-StatusLine "Alias" $box.Alias "White"; Write-StatusLine "Hidden" "$($box.HiddenFromAddressListsEnabled)" "White"
    Write-StatusLine "Forwarding" $(if ($box.ForwardingSmtpAddress) { $box.ForwardingSmtpAddress } else { "(none)" }) "White"

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Add email alias","Remove email alias","Set forwarding","Remove forwarding","Toggle hidden","Set auto-reply") -BackLabel "Done"
    $boxId = $box.PrimarySmtpAddress
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename?")) {
                $ok = Invoke-Action -Description ("Rename shared mailbox '{0}' -> '{1}'" -f $boxId, $v) -Action { Set-Mailbox -Identity $boxId -DisplayName $v -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        1 { $v = Read-UserInput "New alias email"; if ($v -and (Confirm-Action "Add '$v'?")) {
                $ok = Invoke-Action -Description ("Add alias smtp:{0} to '{1}'" -f $v, $boxId) -Action { Set-Mailbox -Identity $boxId -EmailAddresses @{Add="smtp:$v"} -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Added." } } }
        2 { $a = @($box.EmailAddresses | Where-Object { $_ -like "smtp:*" }); if ($a.Count -eq 0) { Write-InfoMsg "No aliases." } else { $al = $a | ForEach-Object { $_ -replace '^smtp:','' }; $s = Show-MultiSelect -Title "Remove" -Options $al; foreach ($i in $s) { if (Confirm-Action "Remove '$($al[$i])'?") {
                $rmv = $a[$i]
                $ok = Invoke-Action -Description ("Remove alias {0} from '{1}'" -f $rmv, $boxId) -Action { Set-Mailbox -Identity $boxId -EmailAddresses @{Remove=$rmv} -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Removed." } } } } }
        3 { $v = Read-UserInput "Forward to"; $kc = Show-Menu -Title "Keep copy?" -Options @("Yes","No") -BackLabel "Cancel"; if ($kc -ne -1 -and $v) { if (Confirm-Action "Set forwarding?") {
                $keepCopy = ($kc -eq 0)
                $ok = Invoke-Action -Description ("Set forwarding on '{0}' -> {1} (keep copy: {2})" -f $boxId, $v, $keepCopy) -Action { Set-Mailbox -Identity $boxId -ForwardingSmtpAddress "smtp:$v" -DeliverToMailboxAndForward $keepCopy -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } } }
        4 { if (Confirm-Action "Remove forwarding?") {
                $ok = Invoke-Action -Description ("Remove forwarding from '{0}'" -f $boxId) -Action { Set-Mailbox -Identity $boxId -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        5 { $nv = -not $box.HiddenFromAddressListsEnabled; if (Confirm-Action "Set hidden to $nv?") {
                $ok = Invoke-Action -Description ("Set HiddenFromAddressListsEnabled = {0} on '{1}'" -f $nv, $boxId) -Action { Set-Mailbox -Identity $boxId -HiddenFromAddressListsEnabled $nv -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        6 { $im = Read-UserInput "Internal message"; $em = Read-UserInput "External (Enter for same)"; if ([string]::IsNullOrWhiteSpace($em)) { $em = $im }; if ($im -and (Confirm-Action "Set auto-reply?")) {
                $ok = Invoke-Action -Description ("Enable auto-reply on '{0}'" -f $boxId) -Action { Set-MailboxAutoReplyConfiguration -Identity $boxId -AutoReplyState Enabled -InternalMessage $im -ExternalMessage $em -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
    }
    Pause-ForUser
}

function Remove-SharedMailboxFlow {
    $box = Find-SharedMailbox; if ($null -eq $box) { Pause-ForUser; return }
    Write-Warn "This permanently deletes the mailbox and ALL contents!"
    if (Confirm-Action "DELETE '$($box.DisplayName)'?") {
        $check = Read-UserInput "Type email to confirm"
        if ($check -eq $box.PrimarySmtpAddress) {
            $ok = Invoke-Action `
                -Description ("DELETE shared mailbox '{0}' <{1}>" -f $box.DisplayName, $box.PrimarySmtpAddress) `
                -ActionType 'DeleteMailbox' `
                -Target @{ mailbox = [string]$box.PrimarySmtpAddress; displayName = [string]$box.DisplayName } `
                -NoUndoReason 'Mailbox deletion is irreversible; restoration requires backup/recovery.' `
                -Action { Remove-Mailbox -Identity $box.PrimarySmtpAddress -Confirm:$false -ErrorAction Stop; $true }
            if ($ok -and -not (Get-PreviewMode)) { Write-Success "Deleted." }
        } else { Write-Warn "Mismatch." }
    }
    Pause-ForUser
}

function Find-SharedMailbox {
    $sm = Show-Menu -Title "Find by" -Options @("Name","Email") -BackLabel "Cancel"; if ($sm -eq -1) { return $null }
    $si = Read-UserInput $(if ($sm -eq 0) { "Mailbox name" } else { "Mailbox email" }); if ([string]::IsNullOrWhiteSpace($si)) { return $null }
    try {
        $boxes = @(if ($sm -eq 0) { Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter "DisplayName -like '*$si*'" -ResultSize 50 } else { Get-Mailbox -RecipientTypeDetails SharedMailbox -Filter "PrimarySmtpAddress -like '*$si*'" -ResultSize 50 })
        if ($boxes.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($boxes.Count -eq 1) { return $boxes[0] }
        $sel = Show-Menu -Title "Select" -Options ($boxes | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }) -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $boxes[$sel]
    } catch { Write-ErrorMsg "$_"; return $null }
}

function Show-MailboxPermissions { param([string]$Id, [string]$Name)
    Write-InfoMsg "Permissions on '$Name':"
    try { $p = Get-MailboxPermission -Identity $Id -ErrorAction Stop | Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and -not $_.IsInherited }
        if ($p.Count -eq 0) { Write-InfoMsg "  (none)" } else { $p | ForEach-Object { Write-Host "    - $($_.User) [$($_.AccessRights -join ', ')]" -ForegroundColor White } }
    } catch { Write-Warn "$_" }
}

function Add-SharedMailboxAccessLoop { param([string]$Id, [string]$Name)
    while ($true) {
        $ui = Read-UserInput "User to grant access (or 'done')"; if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }; if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"; if ($sel -eq -1) { continue }; $f[$sel] } }
            $upn = $tu.UserPrincipalName
            if (Confirm-Action "Grant Full Access to $($tu.DisplayName)?") {
                $ok = Invoke-Action `
                    -Description ("Grant {0} FullAccess on '{1}'" -f $upn, $Name) `
                    -ActionType 'GrantMailboxFullAccess' `
                    -Target @{ mailbox = [string]$Id; user = $upn } `
                    -ReverseType 'RevokeMailboxFullAccess' `
                    -ReverseDescription ("Revoke {0} FullAccess on '{1}'" -f $upn, $Name) `
                    -Action { Add-MailboxPermission -Identity $Id -User $upn -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Granted." }
            }
            $pc = Show-Menu -Title "Send permissions?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
            if ($pc -ne -1 -and $pc -ne 3) {
                if ($pc -eq 0 -or $pc -eq 2) { if (Confirm-Action "Send As?") {
                    $ok = Invoke-Action `
                        -Description ("Grant {0} SendAs on '{1}'" -f $upn, $Name) `
                        -ActionType 'GrantMailboxSendAs' `
                        -Target @{ mailbox = [string]$Id; user = $upn } `
                        -ReverseType 'RevokeMailboxSendAs' `
                        -ReverseDescription ("Revoke {0} SendAs on '{1}'" -f $upn, $Name) `
                        -Action { Add-RecipientPermission -Identity $Id -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop; $true }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Granted." } } }
                if ($pc -eq 1 -or $pc -eq 2) { if (Confirm-Action "Send on Behalf?") {
                    $ok = Invoke-Action -Description ("Grant {0} SendOnBehalf on '{1}'" -f $upn, $Name) -Action { Set-Mailbox -Identity $Id -GrantSendOnBehalfTo @{Add=$upn} -ErrorAction Stop; $true }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Granted." } } }
            }
        } catch { Write-ErrorMsg "$_" }
    }
}
