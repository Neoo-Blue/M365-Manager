# ============================================================
#  CalendarAccess.ps1 - Calendar Permission Management (MS Graph + EXO)
# ============================================================

function Start-CalendarAccessManagement {
    Write-SectionHeader "Calendar Access Management"
    if (-not (Connect-ForTask "CalendarAccess")) { Pause-ForUser; return }

    Write-InfoMsg "Identify the calendar OWNER."
    $owner = Resolve-UserIdentity -PromptText "Enter calendar owner name or email"
    if ($null -eq $owner) { Pause-ForUser; return }
    $calId = "$($owner.UserPrincipalName):\Calendar"

    Write-SectionHeader "Current Permissions"
    try { $perms = Get-MailboxFolderPermission -Identity $calId -ErrorAction Stop
        foreach ($p in $perms) { $pu = if ($p.User.DisplayName) { $p.User.DisplayName } else { $p.User.ToString() }; Write-Host "    $pu  -  $($p.AccessRights -join ', ')" -ForegroundColor White }
    } catch { Write-ErrorMsg "Could not read permissions: $_"; Pause-ForUser; return }

    $action = Show-Menu -Title "Action" -Options @("Add access","Remove access") -BackLabel "Cancel"
    if ($action -eq -1) { return }

    if ($action -eq 0) {
        $si = Read-UserInput "User to grant access (name or email)"
        try {
            $tu = if ($si -match '@') { Get-MgUser -UserId $si -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$si" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; Pause-ForUser; return }
                if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { Pause-ForUser; return }; $f[$sel]
                }
            }
        } catch { Write-ErrorMsg "$_"; Pause-ForUser; return }

        $pl = Show-Menu -Title "Access level" -Options @("Reviewer (read-only)","Editor (full edit)","Author (create + edit own)","Contributor (create only)") -BackLabel "Cancel"
        if ($pl -eq -1) { Pause-ForUser; return }
        $accessMap = @("Reviewer","Editor","Author","Contributor"); $ar = $accessMap[$pl]

        if (Confirm-Action "Grant $ar to $($tu.DisplayName) on $($owner.DisplayName) calendar?") {
            try {
                try { Add-MailboxFolderPermission -Identity $calId -User $tu.UserPrincipalName -AccessRights $ar -ErrorAction Stop; Write-Success "Granted: $ar" }
                catch { if ($_.Exception.Message -match "already exists") { Set-MailboxFolderPermission -Identity $calId -User $tu.UserPrincipalName -AccessRights $ar -ErrorAction Stop; Write-Success "Updated to: $ar" } else { throw $_ } }
            } catch { Write-ErrorMsg "Failed: $_" }
        }
    } else {
        $removable = @($perms | Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" -and $_.User.ToString() -ne "Default" -and $_.User.ToString() -ne "Anonymous" })
        if ($removable.Count -eq 0) { Write-InfoMsg "No custom permissions."; Pause-ForUser; return }
        $labels = $removable | ForEach-Object { $pu = if ($_.User.DisplayName) { $_.User.DisplayName } else { $_.User.ToString() }; "$pu ($($_.AccessRights -join ', '))" }
        $sel = Show-MultiSelect -Title "Remove" -Options $labels
        foreach ($idx in $sel) {
            $p = $removable[$idx]; $pu = if ($p.User.DisplayName) { $p.User.DisplayName } else { $p.User.ToString() }
            if (Confirm-Action "Remove access for '$pu'?") { try { Remove-MailboxFolderPermission -Identity $calId -User $pu -Confirm:$false -ErrorAction Stop; Write-Success "Removed." } catch { Write-ErrorMsg "$_" } }
        }
    }
    Pause-ForUser
}
