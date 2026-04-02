# ============================================================
#  DistributionList.ps1 - Distribution List Management
# ============================================================

function Start-DistributionListManagement {
    Write-SectionHeader "Distribution List Management"

    if (-not (Connect-ForTask "DistributionList")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new distribution list",
        "Add / remove members",
        "View / edit DL properties",
        "Delete a distribution list"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-DistributionListFlow }
        1 { Edit-DistributionListMembers }
        2 { Edit-DistributionListProperties }
        3 { Remove-DistributionListFlow }
        -1 { return }
    }
}

# ==================================================================
#  Create
# ==================================================================
function New-DistributionListFlow {
    Write-SectionHeader "Create New Distribution List"

    $name  = Read-UserInput "Distribution list display name"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-ErrorMsg "Name is required."; Pause-ForUser; return }

    $alias = Read-UserInput "Email alias (e.g. 'sales-team', will become alias@yourdomain.com)"
    if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($name -replace '[^a-zA-Z0-9-]','').ToLower() }

    $primarySmtp = Read-UserInput "Full primary email address (e.g. sales-team@contoso.com)"
    if ([string]::IsNullOrWhiteSpace($primarySmtp)) {
        Write-ErrorMsg "Primary email address is required."
        Pause-ForUser; return
    }

    $managedBy = Read-UserInput "Managed by (owner email, or press Enter to skip)"

    $reqAuth = Show-Menu -Title "Who can send to this DL?" -Options @(
        "Anyone (internal and external)",
        "Only authenticated / internal senders"
    ) -BackLabel "Cancel"
    if ($reqAuth -eq -1) { return }

    $details = "Name    : $name`nAlias   : $alias`nEmail   : $primarySmtp`nOwner   : $(if ($managedBy) { $managedBy } else { '(you)' })`nRestricted: $(if ($reqAuth -eq 1) { 'Yes' } else { 'No' })"
    if (-not (Confirm-Action "Create this distribution list?" $details)) { Pause-ForUser; return }

    try {
        $params = @{
            Name               = $name
            DisplayName        = $name
            Alias              = $alias
            PrimarySmtpAddress = $primarySmtp
            Type               = "Distribution"
        }
        if ($managedBy) { $params["ManagedBy"] = $managedBy }
        if ($reqAuth -eq 1) { $params["RequireSenderAuthenticationEnabled"] = $true }
        else { $params["RequireSenderAuthenticationEnabled"] = $false }

        New-DistributionGroup @params -ErrorAction Stop | Out-Null
        Write-Success "Distribution list '$name' ($primarySmtp) created."

        $addNow = Read-UserInput "Add members now? (y/n)"
        if ($addNow -match '^[Yy]') {
            Add-DLMembersLoop -DLIdentity $primarySmtp -DLName $name
        }
    } catch {
        Write-ErrorMsg "Failed to create DL: $_"
    }
    Pause-ForUser
}

# ==================================================================
#  Members
# ==================================================================
function Edit-DistributionListMembers {
    Write-SectionHeader "Manage DL Members"

    # Identify user first
    $user = Resolve-UserIdentity -PromptText "Enter user name or email"
    if ($null -eq $user) { Pause-ForUser; return }
    $upn = $user.UserPrincipalName

    $action = Show-Menu -Title "Action for $($user.DisplayName)" -Options @(
        "Add to distribution list(s)",
        "Remove from distribution list(s)"
    ) -BackLabel "Cancel"

    if ($action -eq -1) { return }

    if ($action -eq 0) {
        # ---- ADD ----
        $dl = Find-DistributionList
        if ($null -eq $dl) { Pause-ForUser; return }

        if (Confirm-Action "Add '$($user.DisplayName)' to '$($dl.DisplayName)'?") {
            try {
                Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $upn -ErrorAction Stop
                Write-Success "Added to '$($dl.DisplayName)'."
            } catch {
                if ($_.Exception.Message -match "already a member") {
                    Write-Warn "User is already a member."
                } else { Write-ErrorMsg "Failed: $_" }
            }
        }

        # Send permissions
        $permChoice = Show-Menu -Title "Grant additional permissions on '$($dl.DisplayName)'?" -Options @(
            "Grant Send As",
            "Grant Send on Behalf",
            "Both Send As and Send on Behalf",
            "No additional permissions"
        ) -BackLabel "Skip"

        if ($permChoice -ne -1 -and $permChoice -ne 3) {
            if ($permChoice -eq 0 -or $permChoice -eq 2) {
                if (Confirm-Action "Grant Send As on '$($dl.DisplayName)' to $upn?") {
                    try {
                        Add-RecipientPermission -Identity $dl.PrimarySmtpAddress `
                            -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                        Write-Success "Send As granted."
                    } catch { Write-ErrorMsg "Send As failed: $_" }
                }
            }
            if ($permChoice -eq 1 -or $permChoice -eq 2) {
                if (Confirm-Action "Grant Send on Behalf on '$($dl.DisplayName)' to $upn?") {
                    try {
                        Set-DistributionGroup -Identity $dl.PrimarySmtpAddress `
                            -GrantSendOnBehalfTo @{Add = $upn} -ErrorAction Stop
                        Write-Success "Send on Behalf granted."
                    } catch { Write-ErrorMsg "Send on Behalf failed: $_" }
                }
            }
        }
    }
    else {
        # ---- REMOVE ----
        Write-InfoMsg "Retrieving DL memberships for $($user.DisplayName)..."
        try {
            $allDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            $memberOf = @()

            foreach ($dl in $allDLs) {
                $members = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ErrorAction SilentlyContinue
                if ($members | Where-Object { $_.PrimarySmtpAddress -eq $upn }) {
                    $memberOf += $dl
                }
            }

            if ($memberOf.Count -eq 0) {
                Write-InfoMsg "User is not a member of any distribution lists."
                Pause-ForUser; return
            }

            $labels = $memberOf | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
            $selected = Show-MultiSelect -Title "Select DL(s) to remove user from" -Options $labels `
                -Prompt "Enter number(s) (e.g. 1,3)"

            foreach ($idx in $selected) {
                $dl = $memberOf[$idx]
                if (Confirm-Action "Remove '$($user.DisplayName)' from '$($dl.DisplayName)'?") {
                    try {
                        Remove-DistributionGroupMember -Identity $dl.PrimarySmtpAddress `
                            -Member $upn -Confirm:$false -ErrorAction Stop
                        Write-Success "Removed from '$($dl.DisplayName)'."
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
function Edit-DistributionListProperties {
    Write-SectionHeader "View / Edit DL Properties"

    $dl = Find-DistributionList
    if ($null -eq $dl) { Pause-ForUser; return }

    # Refresh full object
    try { $dl = Get-DistributionGroup -Identity $dl.PrimarySmtpAddress -ErrorAction Stop } catch {}

    Write-StatusLine "Display Name"     $dl.DisplayName "White"
    Write-StatusLine "Primary Email"    $dl.PrimarySmtpAddress "White"
    Write-StatusLine "Alias"            $dl.Alias "White"
    Write-StatusLine "Description"      $(if ($dl.Description) { $dl.Description } else { "(none)" }) "White"
    Write-StatusLine "Managed By"       $(($dl.ManagedBy -join "; ")) "White"
    Write-StatusLine "Require Auth"     "$($dl.RequireSenderAuthenticationEnabled)" "White"
    Write-StatusLine "Hidden from GAL"  "$($dl.HiddenFromAddressListsEnabled)" "White"
    Write-Host ""

    # Member count
    try {
        $members = @(Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ErrorAction Stop)
        Write-StatusLine "Member Count" "$($members.Count)" "Cyan"
    } catch {}

    # Send on Behalf
    if ($dl.GrantSendOnBehalfTo.Count -gt 0) {
        Write-StatusLine "Send on Behalf" ($dl.GrantSendOnBehalfTo -join "; ") "White"
    }
    Write-Host ""

    $editChoice = Show-Menu -Title "Edit Properties" -Options @(
        "Change display name",
        "Change description",
        "Change managed by (owner)",
        "Toggle sender authentication (internal only vs open)",
        "Toggle hidden from address book"
    ) -BackLabel "Done"

    switch ($editChoice) {
        0 {
            $newVal = Read-UserInput "New display name"
            if ($newVal -and (Confirm-Action "Rename DL to '$newVal'?")) {
                try {
                    Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -DisplayName $newVal -ErrorAction Stop
                    Write-Success "Display name updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        1 {
            $newVal = Read-UserInput "New description (or 'clear')"
            $setVal = if ($newVal -eq 'clear') { "" } else { $newVal }
            if (Confirm-Action "Update description?") {
                try {
                    Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -Description $setVal -ErrorAction Stop
                    Write-Success "Description updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        2 {
            $newOwner = Read-UserInput "New owner email address"
            if ($newOwner -and (Confirm-Action "Set owner to '$newOwner'?")) {
                try {
                    Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -ManagedBy $newOwner -ErrorAction Stop
                    Write-Success "Owner updated."
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        3 {
            $current = $dl.RequireSenderAuthenticationEnabled
            $newVal = -not $current
            $label = if ($newVal) { "internal senders only" } else { "anyone (internal + external)" }
            if (Confirm-Action "Change sender restriction to: $label ?") {
                try {
                    Set-DistributionGroup -Identity $dl.PrimarySmtpAddress `
                        -RequireSenderAuthenticationEnabled $newVal -ErrorAction Stop
                    Write-Success "Sender authentication set to: $label"
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        4 {
            $current = $dl.HiddenFromAddressListsEnabled
            $newVal = -not $current
            $label = if ($newVal) { "Hidden" } else { "Visible" }
            if (Confirm-Action "Set address book visibility to: $label ?") {
                try {
                    Set-DistributionGroup -Identity $dl.PrimarySmtpAddress `
                        -HiddenFromAddressListsEnabled $newVal -ErrorAction Stop
                    Write-Success "Address book visibility set to: $label"
                } catch { Write-ErrorMsg "Failed: $_" }
            }
        }
    }
    Pause-ForUser
}

# ==================================================================
#  Delete
# ==================================================================
function Remove-DistributionListFlow {
    Write-SectionHeader "Delete Distribution List"

    $dl = Find-DistributionList
    if ($null -eq $dl) { Pause-ForUser; return }

    Write-StatusLine "Name"  $dl.DisplayName "White"
    Write-StatusLine "Email" $dl.PrimarySmtpAddress "White"

    try {
        $members = @(Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ErrorAction Stop)
        Write-StatusLine "Members" "$($members.Count)" "Cyan"
    } catch {}

    Write-Host ""
    Write-Warn "This action is irreversible!"

    if (Confirm-Action "DELETE distribution list '$($dl.DisplayName)'?") {
        $doubleCheck = Read-UserInput "Type the DL email address to confirm deletion"
        if ($doubleCheck -eq $dl.PrimarySmtpAddress) {
            try {
                Remove-DistributionGroup -Identity $dl.PrimarySmtpAddress -Confirm:$false -ErrorAction Stop
                Write-Success "Distribution list '$($dl.DisplayName)' deleted."
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
function Find-DistributionList {
    $searchMethod = Show-Menu -Title "Find distribution list by" -Options @(
        "Search by name",
        "Search by email address"
    ) -BackLabel "Cancel"
    if ($searchMethod -eq -1) { return $null }

    $searchInput = Read-UserInput $(if ($searchMethod -eq 0) { "Enter DL name (partial match)" } else { "Enter DL email address" })
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }

    try {
        if ($searchMethod -eq 0) {
            $dls = @(Get-DistributionGroup -Filter "DisplayName -like '*$searchInput*'" -ResultSize 50 -ErrorAction Stop)
        } else {
            $dls = @(Get-DistributionGroup -Filter "PrimarySmtpAddress -like '*$searchInput*'" -ResultSize 50 -ErrorAction Stop)
        }

        if ($dls.Count -eq 0) {
            Write-ErrorMsg "No distribution lists found matching '$searchInput'."
            return $null
        }
        if ($dls.Count -eq 1) {
            Write-Success "Found: $($dls[0].DisplayName) ($($dls[0].PrimarySmtpAddress))"
            return $dls[0]
        }

        $labels = $dls | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
        $sel = Show-Menu -Title "Select Distribution List" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return $dls[$sel]
    } catch {
        Write-ErrorMsg "Search error: $_"
        return $null
    }
}

function Add-DLMembersLoop {
    param([string]$DLIdentity, [string]$DLName)

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

            if (Confirm-Action "Add '$($targetUser.DisplayName)' to '$DLName'?") {
                Add-DistributionGroupMember -Identity $DLIdentity -Member $targetUser.UserPrincipalName -ErrorAction Stop
                Write-Success "Added '$($targetUser.DisplayName)'."
            }
        } catch {
            if ($_.Exception.Message -match "already a member") {
                Write-Warn "User is already a member."
            } else { Write-ErrorMsg "Failed: $_" }
        }
    }
}
