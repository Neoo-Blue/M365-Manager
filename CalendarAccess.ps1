# ============================================================
#  CalendarAccess.ps1 - Calendar Permission Management
# ============================================================

function Start-CalendarAccessManagement {
    Write-SectionHeader "Calendar Access Management"

    if (-not (Connect-ForTask "CalendarAccess")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ---- Identify the calendar owner ----
    Write-InfoMsg "First, identify the calendar OWNER (whose calendar to modify)."
    $owner = Resolve-UserIdentity -PromptText "Enter calendar owner name or email"
    if ($null -eq $owner) { Pause-ForUser; return }

    $ownerUpn = $owner.UserPrincipalName
    $calendarId = "${ownerUpn}:\Calendar"

    # Show current permissions
    Write-SectionHeader "Current Calendar Permissions for $($owner.DisplayName)"
    try {
        $currentPerms = Get-MailboxFolderPermission -Identity $calendarId -ErrorAction Stop
        if ($currentPerms.Count -eq 0) {
            Write-InfoMsg "No custom permissions found."
        } else {
            foreach ($perm in $currentPerms) {
                $permUser = if ($perm.User.DisplayName) { $perm.User.DisplayName } else { $perm.User.ToString() }
                Write-Host "    $permUser  -  $($perm.AccessRights -join ', ')" -ForegroundColor White
            }
        }
    } catch {
        Write-ErrorMsg "Could not read calendar permissions: $_"
        Pause-ForUser; return
    }

    Write-Host ""

    $action = Show-Menu -Title "Action" -Options @(
        "Add calendar access for a user",
        "Remove calendar access for a user"
    ) -BackLabel "Cancel"

    if ($action -eq -1) { return }

    if ($action -eq 0) {
        # ---- ADD ----
        $searchMethod = Show-Menu -Title "Find the user to grant access to by" -Options @(
            "Search by name",
            "Search by email address"
        ) -BackLabel "Cancel"

        if ($searchMethod -eq -1) { return }

        $searchInput = Read-UserInput $(if ($searchMethod -eq 0) { "Enter user name" } else { "Enter user email" })

        try {
            if ($searchInput -match '@') {
                $targetUser = Get-AzureADUser -ObjectId $searchInput -ErrorAction Stop
            } else {
                $found = Get-AzureADUser -SearchString $searchInput -ErrorAction Stop
                if ($found.Count -eq 0) {
                    Write-ErrorMsg "No user found matching '$searchInput'."
                    Pause-ForUser; return
                }
                if ($found.Count -eq 1) {
                    $targetUser = $found[0]
                } else {
                    $names = $found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                    $sel = Show-Menu -Title "Select User" -Options $names -BackLabel "Cancel"
                    if ($sel -eq -1) { Pause-ForUser; return }
                    $targetUser = $found[$sel]
                }
            }

            Write-StatusLine "Granting to" "$($targetUser.DisplayName) ($($targetUser.UserPrincipalName))" "White"
        } catch {
            Write-ErrorMsg "Could not find user: $_"
            Pause-ForUser; return
        }

        # ---- Permission level ----
        $permLevel = Show-Menu -Title "Access level" -Options @(
            "Reviewer  (read-only, can view all details)",
            "Editor    (can create, edit, and delete items)",
            "Author    (can create items and edit own items)",
            "Contributor (can create items only)"
        ) -BackLabel "Cancel"

        if ($permLevel -eq -1) { return }

        $accessMap = @("Reviewer","Editor","Author","Contributor")
        $accessRight = $accessMap[$permLevel]

        $details = "Owner: $ownerUpn`nUser: $($targetUser.UserPrincipalName)`nAccess: $accessRight"
        if (Confirm-Action "Add calendar permission?" $details) {
            try {
                # Try to add; if user already has perms, set instead
                try {
                    Add-MailboxFolderPermission -Identity $calendarId `
                        -User $targetUser.UserPrincipalName `
                        -AccessRights $accessRight -ErrorAction Stop
                    Write-Success "Calendar access granted: $accessRight"
                } catch {
                    if ($_.Exception.Message -match "already exists") {
                        Write-Warn "User already has permissions. Updating..."
                        Set-MailboxFolderPermission -Identity $calendarId `
                            -User $targetUser.UserPrincipalName `
                            -AccessRights $accessRight -ErrorAction Stop
                        Write-Success "Calendar access updated to: $accessRight"
                    } else {
                        throw $_
                    }
                }
            } catch {
                Write-ErrorMsg "Failed to set calendar permission: $_"
            }
        }
    }
    else {
        # ---- REMOVE ----
        try {
            $removable = $currentPerms | Where-Object {
                $_.User.DisplayName -ne "Default" -and
                $_.User.DisplayName -ne "Anonymous" -and
                $_.User.ToString() -ne "Default" -and
                $_.User.ToString() -ne "Anonymous"
            }

            if ($removable.Count -eq 0) {
                Write-InfoMsg "No custom user permissions to remove."
                Pause-ForUser; return
            }

            $permLabels = $removable | ForEach-Object {
                $permUser = if ($_.User.DisplayName) { $_.User.DisplayName } else { $_.User.ToString() }
                "$permUser  ($($_.AccessRights -join ', '))"
            }

            $selected = Show-MultiSelect -Title "Select permission(s) to remove" -Options $permLabels `
                -Prompt "Enter number(s) (e.g. 1,3)"

            foreach ($idx in $selected) {
                $perm = $removable[$idx]
                $permUser = if ($perm.User.DisplayName) { $perm.User.DisplayName } else { $perm.User.ToString() }

                if (Confirm-Action "Remove calendar access for '$permUser' on $($owner.DisplayName)'s calendar?") {
                    try {
                        Remove-MailboxFolderPermission -Identity $calendarId `
                            -User $permUser -Confirm:$false -ErrorAction Stop
                        Write-Success "Removed access for '$permUser'."
                    } catch {
                        Write-ErrorMsg "Failed to remove: $_"
                    }
                }
            }
        } catch {
            Write-ErrorMsg "Error processing removal: $_"
        }
    }

    Write-Success "Calendar access management complete."
    Pause-ForUser
}
