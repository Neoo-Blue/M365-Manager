# ============================================================
#  GroupManager.ps1 - Unified Group Membership Manager
#  View, edit, bulk remove, and replicate SG/DL/M365 memberships
# ============================================================

function Start-GroupManagerMenu {
    Write-SectionHeader "Group Membership Manager"

    if (-not (Connect-ForTask "GroupManager")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Group Membership Manager" -Options @(
            "View all memberships for a user",
            "Bulk remove memberships",
            "Add user to groups (search)",
            "Replicate memberships between users"
        ) -BackLabel "Back to Main Menu"

        switch ($action) {
            0 { View-UserAllMemberships }
            1 { BulkRemove-UserMemberships }
            2 { BulkAdd-UserToGroups }
            3 { Replicate-Memberships }
            -1 { $keepGoing = $false }
        }
    }
}

# ============================================================
#  Fetch all memberships (SG, DL, M365 groups)
# ============================================================

function Get-AllUserMemberships {
    <# Returns a combined list of all group memberships with type tags. #>
    param([object]$User)

    $upn = $User.UserPrincipalName
    $userId = $User.Id
    $all = @()

    # ---- Azure AD groups (SG, M365) via Graph ----
    Write-InfoMsg "Fetching Azure AD group memberships..."
    try {
        $graphGroups = @(Get-MgUserMemberOf -UserId $userId -All -ErrorAction Stop)
        foreach ($g in $graphGroups) {
            if ($g.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                $secEnabled = $g.AdditionalProperties["securityEnabled"]
                $mailEnabled = $g.AdditionalProperties["mailEnabled"]
                $gType = "M365 Group"
                if ($secEnabled -and -not $mailEnabled) { $gType = "Security Group" }
                elseif ($secEnabled -and $mailEnabled)  { $gType = "Mail-Enabled SG" }
                elseif (-not $secEnabled -and $mailEnabled) { $gType = "M365 Group" }

                $all += [PSCustomObject]@{
                    Name    = $g.AdditionalProperties["displayName"]
                    Email   = $g.AdditionalProperties["mail"]
                    Type    = $gType
                    Id      = $g.Id
                    Source  = "Graph"
                }
            }
        }
    } catch { Write-Warn "Graph group fetch error: $_" }

    # ---- Distribution Lists via EXO ----
    Write-InfoMsg "Fetching distribution list memberships..."
    try {
        $allDLs = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
        foreach ($dl in $allDLs) {
            $members = @(Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ResultSize Unlimited -ErrorAction SilentlyContinue)
            $isMember = $members | Where-Object { $_.PrimarySmtpAddress -eq $upn -or $_.WindowsLiveID -eq $upn }
            if ($isMember) {
                # Check if already in list (avoid duplicates with mail-enabled SGs)
                $exists = $all | Where-Object { $_.Email -eq $dl.PrimarySmtpAddress }
                if (-not $exists) {
                    $all += [PSCustomObject]@{
                        Name    = $dl.DisplayName
                        Email   = $dl.PrimarySmtpAddress
                        Type    = "Distribution List"
                        Id      = $dl.Guid.ToString()
                        Source  = "EXO"
                    }
                }
            }
        }
    } catch { Write-Warn "DL fetch error: $_" }

    return $all
}

# ============================================================
#  1. View all memberships
# ============================================================

function View-UserAllMemberships {
    Write-SectionHeader "View All Group Memberships"

    $user = Resolve-UserIdentity
    if ($null -eq $user) { return }

    $memberships = Get-AllUserMemberships -User $user
    if ($memberships.Count -eq 0) {
        Write-InfoMsg "$($user.DisplayName) has no group memberships."
        Pause-ForUser; return
    }

    # Display grouped by type
    $grouped = $memberships | Group-Object Type | Sort-Object Name

    Write-SectionHeader "Memberships for $($user.DisplayName) ($($memberships.Count) total)"

    foreach ($group in $grouped) {
        Write-Host ""
        Write-Host "  --- $($group.Name) ($($group.Count)) ---" -ForegroundColor $script:Colors.Highlight
        foreach ($m in ($group.Group | Sort-Object Name)) {
            $emailStr = if ($m.Email) { " ($($m.Email))" } else { "" }
            Write-Host "    - $($m.Name)$emailStr" -ForegroundColor White
        }
    }

    Pause-ForUser
}

# ============================================================
#  2. Bulk remove memberships
# ============================================================

function BulkRemove-UserMemberships {
    Write-SectionHeader "Bulk Remove Group Memberships"

    $user = Resolve-UserIdentity
    if ($null -eq $user) { return }

    $memberships = Get-AllUserMemberships -User $user
    if ($memberships.Count -eq 0) {
        Write-InfoMsg "No memberships to remove."
        Pause-ForUser; return
    }

    # Filter by type?
    $filterChoice = Show-Menu -Title "Show which group types?" -Options @(
        "All types",
        "Security Groups only",
        "Distribution Lists only",
        "M365 / Mail-Enabled Groups only"
    ) -BackLabel "Back"
    if ($filterChoice -eq -1) { return }

    $filtered = switch ($filterChoice) {
        0 { $memberships }
        1 { @($memberships | Where-Object { $_.Type -match "Security" }) }
        2 { @($memberships | Where-Object { $_.Type -eq "Distribution List" }) }
        3 { @($memberships | Where-Object { $_.Type -match "M365|Mail-Enabled" }) }
    }

    if ($filtered.Count -eq 0) {
        Write-InfoMsg "No groups matching that filter."
        Pause-ForUser; return
    }

    $labels = $filtered | ForEach-Object { "$($_.Name) [$($_.Type)]$(if ($_.Email) { " ($($_.Email))" })" }
    $selected = Show-MultiSelect -Title "Select group(s) to remove $($user.DisplayName) from" -Options $labels

    foreach ($idx in $selected) {
        $grp = $filtered[$idx]
        if (Confirm-Action "Remove from '$($grp.Name)' ($($grp.Type))?") {
            try {
                if ($grp.Source -eq "Graph") {
                    Remove-MgGroupMemberByRef -GroupId $grp.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                    Write-Success "Removed from '$($grp.Name)'."
                }
                elseif ($grp.Source -eq "EXO") {
                    Remove-DistributionGroupMember -Identity $grp.Email -Member $user.UserPrincipalName -Confirm:$false -ErrorAction Stop
                    Write-Success "Removed from '$($grp.Name)'."
                }
            } catch { Write-ErrorMsg "Failed to remove from '$($grp.Name)': $_" }
        }
    }

    Pause-ForUser
}

# ============================================================
#  3. Add user to groups (search)
# ============================================================

function BulkAdd-UserToGroups {
    Write-SectionHeader "Add User to Groups"

    $user = Resolve-UserIdentity
    if ($null -eq $user) { return }

    $adding = $true
    while ($adding) {
        $groupType = Show-Menu -Title "What type of group to add?" -Options @(
            "Security Group",
            "Distribution List",
            "Search any group by name"
        ) -BackLabel "Done"
        if ($groupType -eq -1) { break }

        $searchInput = Read-UserInput "Search group by name"
        if ([string]::IsNullOrWhiteSpace($searchInput)) { continue }

        if ($groupType -eq 0 -or $groupType -eq 2) {
            # Search Graph groups
            try {
                $gGroups = @(Get-MgGroup -Search "displayName:$searchInput" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($groupType -eq 0) { $gGroups = @($gGroups | Where-Object { $_.SecurityEnabled -and -not $_.MailEnabled }) }
                if ($gGroups.Count -gt 0) {
                    $gLabels = $gGroups | ForEach-Object { "$($_.DisplayName) $(if ($_.Mail) { "($($_.Mail))" })" }
                    $sel = Show-MultiSelect -Title "Select group(s)" -Options $gLabels
                    foreach ($i in $sel) {
                        $g = $gGroups[$i]
                        if (Confirm-Action "Add to '$($g.DisplayName)'?") {
                            try {
                                New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                                Write-Success "Added to '$($g.DisplayName)'."
                            } catch {
                                if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." }
                                else { Write-ErrorMsg "Failed: $_" }
                            }
                        }
                    }
                } else { Write-Warn "No Graph groups found." }
            } catch { Write-ErrorMsg "Graph search error: $_" }
        }

        if ($groupType -eq 1 -or $groupType -eq 2) {
            # Search DLs
            try {
                $dls = @(Get-DistributionGroup -Filter "DisplayName -like '*$searchInput*'" -ResultSize 50 -ErrorAction Stop)
                if ($dls.Count -gt 0) {
                    $dlLabels = $dls | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
                    $sel = Show-MultiSelect -Title "Select DL(s)" -Options $dlLabels
                    foreach ($i in $sel) {
                        $dl = $dls[$i]
                        if (Confirm-Action "Add to '$($dl.DisplayName)'?") {
                            try {
                                Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $user.UserPrincipalName -ErrorAction Stop
                                Write-Success "Added to '$($dl.DisplayName)'."
                            } catch {
                                if ($_.Exception.Message -match "already") { Write-Warn "Already a member." }
                                else { Write-ErrorMsg "Failed: $_" }
                            }
                        }
                    }
                } else { Write-Warn "No DLs found." }
            } catch { Write-ErrorMsg "DL search error: $_" }
        }

        $cont = Read-UserInput "Add to more groups? (y/n)"
        if ($cont -notmatch '^[Yy]') { $adding = $false }
    }
    Pause-ForUser
}

# ============================================================
#  4. Replicate memberships between users
# ============================================================

function Replicate-Memberships {
    Write-SectionHeader "Replicate Group Memberships"

    Write-InfoMsg "First, select the SOURCE user (copy FROM)."
    $sourceUser = Resolve-UserIdentity -PromptText "Source user name or email"
    if ($null -eq $sourceUser) { return }

    Write-InfoMsg "Now, select the TARGET user (copy TO)."
    $targetUser = Resolve-UserIdentity -PromptText "Target user name or email"
    if ($null -eq $targetUser) { return }

    if ($sourceUser.Id -eq $targetUser.Id) {
        Write-ErrorMsg "Source and target are the same user."
        Pause-ForUser; return
    }

    # Fetch both users' memberships
    Write-InfoMsg "Fetching memberships for source: $($sourceUser.DisplayName)..."
    $sourceMemberships = Get-AllUserMemberships -User $sourceUser

    Write-InfoMsg "Fetching memberships for target: $($targetUser.DisplayName)..."
    $targetMemberships = Get-AllUserMemberships -User $targetUser

    # Display both
    Write-Host ""
    Write-StatusLine "Source" "$($sourceUser.DisplayName) - $($sourceMemberships.Count) group(s)" "Cyan"
    Write-StatusLine "Target" "$($targetUser.DisplayName) - $($targetMemberships.Count) group(s)" "Cyan"
    Write-Host ""

    # Show source groups
    if ($sourceMemberships.Count -gt 0) {
        Write-InfoMsg "Source groups:"
        foreach ($m in ($sourceMemberships | Sort-Object Type, Name)) {
            Write-Host "    - $($m.Name) [$($m.Type)]" -ForegroundColor White
        }
    }
    Write-Host ""
    if ($targetMemberships.Count -gt 0) {
        Write-InfoMsg "Target current groups:"
        foreach ($m in ($targetMemberships | Sort-Object Type, Name)) {
            Write-Host "    - $($m.Name) [$($m.Type)]" -ForegroundColor White
        }
    }
    Write-Host ""

    # Replication mode
    $mode = Show-Menu -Title "Replication Mode" -Options @(
        "Selective - choose which groups to copy",
        "Full Copy - add all source groups (keep target's existing groups)",
        "Full Replace - remove target's current groups and replace with source's"
    ) -BackLabel "Cancel"

    if ($mode -eq -1) { return }

    switch ($mode) {
        0 { Invoke-SelectiveReplicate -Source $sourceMemberships -Target $targetUser -TargetMemberships $targetMemberships }
        1 { Invoke-FullCopyReplicate -Source $sourceMemberships -Target $targetUser -TargetMemberships $targetMemberships }
        2 { Invoke-FullReplaceReplicate -Source $sourceMemberships -Target $targetUser -TargetMemberships $targetMemberships }
    }

    Pause-ForUser
}

function Invoke-SelectiveReplicate {
    param([array]$Source, [object]$Target, [array]$TargetMemberships)

    $labels = $Source | ForEach-Object {
        $alreadyIn = $TargetMemberships | Where-Object { $_.Id -eq $_.Id -or $_.Name -eq $_.Name }
        "$($_.Name) [$($_.Type)]$(if ($TargetMemberships | Where-Object { $_.Name -eq $Source[$([array]::IndexOf($Source, $_))].Name }) { ' (already member)' })"
    }

    # Recalculate labels properly
    $labels = @()
    for ($i = 0; $i -lt $Source.Count; $i++) {
        $s = $Source[$i]
        $existing = $TargetMemberships | Where-Object { $_.Name -eq $s.Name }
        $tag = if ($existing) { " ** already member **" } else { "" }
        $labels += "$($s.Name) [$($s.Type)]$tag"
    }

    $selected = Show-MultiSelect -Title "Select groups to add to $($Target.DisplayName)" -Options $labels

    foreach ($idx in $selected) {
        $grp = $Source[$idx]
        $existing = $TargetMemberships | Where-Object { $_.Name -eq $grp.Name }
        if ($existing) { Write-InfoMsg "Skipping '$($grp.Name)' - already a member."; continue }

        if (Confirm-Action "Add $($Target.DisplayName) to '$($grp.Name)' ($($grp.Type))?") {
            Add-UserToGroup -UserId $Target.Id -UserUPN $Target.UserPrincipalName -Group $grp
        }
    }
}

function Invoke-FullCopyReplicate {
    param([array]$Source, [object]$Target, [array]$TargetMemberships)

    $toAdd = @()
    foreach ($s in $Source) {
        $existing = $TargetMemberships | Where-Object { $_.Name -eq $s.Name }
        if (-not $existing) { $toAdd += $s }
    }

    if ($toAdd.Count -eq 0) {
        Write-InfoMsg "$($Target.DisplayName) already has all source groups."
        return
    }

    Write-InfoMsg "Groups to add ($($toAdd.Count)):"
    foreach ($g in $toAdd) { Write-Host "    + $($g.Name) [$($g.Type)]" -ForegroundColor $script:Colors.Success }
    Write-InfoMsg "Groups already present (skipping): $($Source.Count - $toAdd.Count)"
    Write-Host ""

    if (-not (Confirm-Action "Add $($toAdd.Count) group(s) to $($Target.DisplayName)?")) { return }

    foreach ($grp in $toAdd) {
        Add-UserToGroup -UserId $Target.Id -UserUPN $Target.UserPrincipalName -Group $grp
    }
}

function Invoke-FullReplaceReplicate {
    param([array]$Source, [object]$Target, [array]$TargetMemberships)

    # Calculate changes
    $toRemove = @()
    foreach ($t in $TargetMemberships) {
        $inSource = $Source | Where-Object { $_.Name -eq $t.Name }
        if (-not $inSource) { $toRemove += $t }
    }

    $toAdd = @()
    foreach ($s in $Source) {
        $existing = $TargetMemberships | Where-Object { $_.Name -eq $s.Name }
        if (-not $existing) { $toAdd += $s }
    }

    $keepCount = $Source.Count - $toAdd.Count

    Write-Host ""
    if ($toRemove.Count -gt 0) {
        Write-Warn "Groups to REMOVE from $($Target.DisplayName) ($($toRemove.Count)):"
        foreach ($g in $toRemove) { Write-Host "    - $($g.Name) [$($g.Type)]" -ForegroundColor $script:Colors.Error }
    }
    if ($toAdd.Count -gt 0) {
        Write-InfoMsg "Groups to ADD ($($toAdd.Count)):"
        foreach ($g in $toAdd) { Write-Host "    + $($g.Name) [$($g.Type)]" -ForegroundColor $script:Colors.Success }
    }
    Write-InfoMsg "Groups unchanged: $keepCount"
    Write-Host ""

    Write-Warn "This will modify $($Target.DisplayName)'s memberships to match $($Source[0].Name -replace '.*','the source user')."
    if (-not (Confirm-Action "Proceed with full replace? ($($toRemove.Count) removals, $($toAdd.Count) additions)")) { return }

    # Remove first
    foreach ($grp in $toRemove) {
        try {
            if ($grp.Source -eq "Graph") {
                Remove-MgGroupMemberByRef -GroupId $grp.Id -DirectoryObjectId $Target.Id -ErrorAction Stop
                Write-Success "Removed from '$($grp.Name)'."
            } elseif ($grp.Source -eq "EXO") {
                Remove-DistributionGroupMember -Identity $grp.Email -Member $Target.UserPrincipalName -Confirm:$false -ErrorAction Stop
                Write-Success "Removed from '$($grp.Name)'."
            }
        } catch { Write-ErrorMsg "Failed to remove from '$($grp.Name)': $_" }
    }

    # Then add
    foreach ($grp in $toAdd) {
        Add-UserToGroup -UserId $Target.Id -UserUPN $Target.UserPrincipalName -Group $grp
    }
}

# ============================================================
#  Helper: Add user to a group based on Source type
# ============================================================

function Add-UserToGroup {
    param([string]$UserId, [string]$UserUPN, [PSCustomObject]$Group)

    try {
        if ($Group.Source -eq "Graph") {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $UserId -ErrorAction Stop
            Write-Success "Added to '$($Group.Name)'."
        }
        elseif ($Group.Source -eq "EXO") {
            Add-DistributionGroupMember -Identity $Group.Email -Member $UserUPN -ErrorAction Stop
            Write-Success "Added to '$($Group.Name)'."
        }
    } catch {
        if ($_.Exception.Message -match "already exist|already a member") {
            Write-Warn "Already a member of '$($Group.Name)'."
        } else {
            Write-ErrorMsg "Failed to add to '$($Group.Name)': $_"
        }
    }
}
