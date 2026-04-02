# ============================================================
#  SecurityGroup.ps1 - Security Group Management
# ============================================================

function Start-SecurityGroupManagement {
    Write-SectionHeader "Security Group Management"

    if (-not (Connect-ForTask "SecurityGroup")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new security group",
        "Add / remove members",
        "View / edit group properties",
        "Delete a security group"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-SecurityGroup }
        1 { Edit-SecurityGroupMembers }
        2 { Edit-SecurityGroupProperties }
        3 { Remove-SecurityGroupFlow }
        -1 { return }
    }
}

# ==================================================================
#  Create
# ==================================================================
function New-SecurityGroup {
    Write-SectionHeader "Create New Security Group"

    $name  = Read-UserInput "Group display name"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-ErrorMsg "Name is required."; Pause-ForUser; return }

    $desc  = Read-UserInput "Description (or press Enter to skip)"
    $mail  = Read-UserInput "Mail nickname (no spaces, used for email-enabled groups)"
    if ([string]::IsNullOrWhiteSpace($mail)) { $mail = ($name -replace '[^a-zA-Z0-9]','').ToLower() }

    $mailEnabled = Show-Menu -Title "Mail-enabled?" -Options @(
        "No  (standard security group)",
        "Yes (mail-enabled security group)"
    ) -BackLabel "Cancel"
    if ($mailEnabled -eq -1) { return }

    $details = "Name       : $name`nNickname   : $mail`nDescription: $desc`nMail-enabled: $(if ($mailEnabled -eq 1) { 'Yes' } else { 'No' })"
    if (-not (Confirm-Action "Create this security group?" $details)) { Pause-ForUser; return }

    try {
        $params = @{
            DisplayName     = $name
            MailEnabled     = ($mailEnabled -eq 1)
            MailNickName    = $mail
            SecurityEnabled = $true
        }
        if ($desc) { $params["Description"] = $desc }

        $newGroup = New-AzureADGroup @params -ErrorAction Stop
        Write-Success "Security group '$name' created. ObjectId: $($newGroup.ObjectId)"

        $addNow = Read-UserInput "Add members now? (y/n)"
        if ($addNow -match '^[Yy]') {
            Add-MembersLoop -GroupObjectId $newGroup.ObjectId -GroupName $name
        }
    } catch {
        Write-ErrorMsg "Failed to create group: $_"
    }
    Pause-ForUser
}

# ==================================================================
#  Members
# ==================================================================
function Edit-SecurityGroupMembers {
    Write-SectionHeader "Manage Security Group Members"

    $group = Find-SecurityGroup
    if ($null -eq $group) { Pause-ForUser; return }

    # Show current members
    Show-GroupMembers -GroupObjectId $group.ObjectId -GroupName $group.DisplayName

    $action = Show-Menu -Title "Action for '$($group.DisplayName)'" -Options @(
        "Add member(s)",
        "Remove member(s)"
    ) -BackLabel "Done"

    if ($action -eq 0) {
        Add-MembersLoop -GroupObjectId $group.ObjectId -GroupName $group.DisplayName
    }
    elseif ($action -eq 1) {
        # Remove
        try {
            $members = @(Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true -ErrorAction Stop)
            if ($members.Count -eq 0) {
                Write-InfoMsg "Group has no members."
                Pause-ForUser; return
            }
            $labels = $members | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
            $selected = Show-MultiSelect -Title "Select member(s) to remove" -Options $labels `
                -Prompt "Enter number(s) (e.g. 1,3)"

            foreach ($idx in $selected) {
                $m = $members[$idx]
                if (Confirm-Action "Remove '$($m.DisplayName)' from '$($group.DisplayName)'?") {
                    try {
                        Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $m.ObjectId -ErrorAction Stop
                        Write-Success "Removed '$($m.DisplayName)'."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        } catch { Write-ErrorMsg "Error: $_" }
    }
    Pause-ForUser
}

# ==================================================================
#  Properties
# ==================================================================
function Edit-SecurityGroupProperties {
    Write-SectionHeader "View / Edit Security Group Properties"

    $group = Find-SecurityGroup
    if ($null -eq $group) { Pause-ForUser; return }

    Write-StatusLine "Display Name"   $group.DisplayName "White"
    Write-StatusLine "Description"    $(if ($group.Description) { $group.Description } else { "(none)" }) "White"
    Write-StatusLine "Mail Nickname"  $group.MailNickName "White"
    Write-StatusLine "Mail Enabled"   "$($group.MailEnabled)" "White"
    Write-StatusLine "Mail"           $(if ($group.Mail) { $group.Mail } else { "(none)" }) "White"
    Write-StatusLine "Object ID"      $group.ObjectId "Gray"
    Write-Host ""

    # Show member count
    try {
        $members = @(Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true -ErrorAction Stop)
        Write-StatusLine "Member Count" "$($members.Count)" "Cyan"
    } catch {}

    Write-Host ""

    $editChoice = Show-Menu -Title "Edit Properties" -Options @(
        "Change display name",
        "Change description",
        "Change mail nickname"
    ) -BackLabel "Done"

    switch ($editChoice) {
        0 {
            $newVal = Read-UserInput "New display name"
            if ($newVal -and (Confirm-Action "Rename group to '$newVal'?")) {
                try {
                    Set-AzureADGroup -ObjectId $group.ObjectId -DisplayName $newVal -ErrorAction Stop
                    Write-Success "Display name updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        1 {
            $newVal = Read-UserInput "New description (or 'clear')"
            $setVal = if ($newVal -eq 'clear') { $null } else { $newVal }
            if (Confirm-Action "Update description?") {
                try {
                    Set-AzureADGroup -ObjectId $group.ObjectId -Description $setVal -ErrorAction Stop
                    Write-Success "Description updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        2 {
            $newVal = Read-UserInput "New mail nickname"
            if ($newVal -and (Confirm-Action "Change mail nickname to '$newVal'?")) {
                try {
                    Set-AzureADGroup -ObjectId $group.ObjectId -MailNickName $newVal -ErrorAction Stop
                    Write-Success "Mail nickname updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
    }
    Pause-ForUser
}

# ==================================================================
#  Delete
# ==================================================================
function Remove-SecurityGroupFlow {
    Write-SectionHeader "Delete Security Group"

    $group = Find-SecurityGroup
    if ($null -eq $group) { Pause-ForUser; return }

    Write-StatusLine "Group"    $group.DisplayName "White"
    Write-StatusLine "ObjectId" $group.ObjectId "Gray"

    try {
        $members = @(Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true -ErrorAction Stop)
        Write-StatusLine "Members" "$($members.Count)" "Cyan"
    } catch {}

    Write-Host ""
    Write-Warn "This action is irreversible!"

    if (Confirm-Action "DELETE security group '$($group.DisplayName)'?") {
        $doubleCheck = Read-UserInput "Type the group name to confirm deletion"
        if ($doubleCheck -eq $group.DisplayName) {
            try {
                Remove-AzureADGroup -ObjectId $group.ObjectId -ErrorAction Stop
                Write-Success "Security group '$($group.DisplayName)' deleted."
            } catch { Write-ErrorMsg "Failed to delete: $_" }
        } else {
            Write-Warn "Name did not match. Deletion cancelled."
        }
    }
    Pause-ForUser
}

# ==================================================================
#  Shared helpers
# ==================================================================
function Find-SecurityGroup {
    $searchInput = Read-UserInput "Search for security group by name or email"
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }

    try {
        $groups = Get-AzureADGroup -SearchString $searchInput -All $true |
            Where-Object { $_.SecurityEnabled -eq $true }

        if ($groups.Count -eq 0) {
            Write-ErrorMsg "No security groups found matching '$searchInput'."
            return $null
        }
        if ($groups.Count -eq 1) {
            Write-Success "Found: $($groups[0].DisplayName)"
            return $groups[0]
        }

        $labels = $groups | ForEach-Object {
            "$($_.DisplayName)  $(if ($_.Mail) { "($($_.Mail))" } else { '' })"
        }
        $sel = Show-Menu -Title "Select Security Group" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return $groups[$sel]
    } catch {
        Write-ErrorMsg "Search error: $_"
        return $null
    }
}

function Show-GroupMembers {
    param([string]$GroupObjectId, [string]$GroupName)
    try {
        $members = @(Get-AzureADGroupMember -ObjectId $GroupObjectId -All $true -ErrorAction Stop)
        Write-InfoMsg "Current members of '$GroupName' ($($members.Count)):"
        if ($members.Count -eq 0) {
            Write-InfoMsg "  (no members)"
        } else {
            foreach ($m in $members) {
                Write-Host "    - $($m.DisplayName) ($($m.UserPrincipalName))" -ForegroundColor White
            }
        }
        Write-Host ""
    } catch { Write-Warn "Could not retrieve members: $_" }
}

function Add-MembersLoop {
    param([string]$GroupObjectId, [string]$GroupName)

    $adding = $true
    while ($adding) {
        $userInput = Read-UserInput "Enter user name or email to add (or 'done')"
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

            if (Confirm-Action "Add '$($targetUser.DisplayName)' to '$GroupName'?") {
                Add-AzureADGroupMember -ObjectId $GroupObjectId -RefObjectId $targetUser.ObjectId -ErrorAction Stop
                Write-Success "Added '$($targetUser.DisplayName)'."
            }
        } catch {
            if ($_.Exception.Message -match "already exist") {
                Write-Warn "User is already a member."
            } else { Write-ErrorMsg "Failed: $_" }
        }
    }
}
