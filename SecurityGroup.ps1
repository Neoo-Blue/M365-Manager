# ============================================================
#  SecurityGroup.ps1 - Security Group Management (MS Graph)
# ============================================================

function Start-SecurityGroupManagement {
    Write-SectionHeader "Security Group Management"
    if (-not (Connect-ForTask "SecurityGroup")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new security group","Add / remove members",
        "View / edit group properties","Delete a security group"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-SecurityGroup }
        1 { Edit-SecurityGroupMembers }
        2 { Edit-SecurityGroupProperties }
        3 { Remove-SecurityGroupFlow }
    }
}

function New-SecurityGroup {
    Write-SectionHeader "Create New Security Group"
    $name = Read-UserInput "Group display name"
    if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $desc = Read-UserInput "Description (or Enter to skip)"
    $mail = Read-UserInput "Mail nickname (no spaces)"
    if ([string]::IsNullOrWhiteSpace($mail)) { $mail = ($name -replace '[^a-zA-Z0-9]','').ToLower() }
    $me = Show-Menu -Title "Mail-enabled?" -Options @("No (standard)","Yes (mail-enabled)") -BackLabel "Cancel"
    if ($me -eq -1) { return }

    if (Confirm-Action "Create security group '$name'?") {
        try {
            $body = @{ DisplayName = $name; MailEnabled = ($me -eq 1); MailNickname = $mail; SecurityEnabled = $true }
            if ($desc) { $body["Description"] = $desc }
            $g = New-MgGroup -BodyParameter $body -ErrorAction Stop
            Write-Success "Created. Id: $($g.Id)"
            $add = Read-UserInput "Add members now? (y/n)"
            if ($add -match '^[Yy]') { Add-MembersLoop -GroupId $g.Id -GroupName $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-SecurityGroupMembers {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Show-GroupMembers -GroupId $group.Id -GroupName $group.DisplayName
    $action = Show-Menu -Title "Action" -Options @("Add member(s)","Remove member(s)") -BackLabel "Done"
    if ($action -eq 0) { Add-MembersLoop -GroupId $group.Id -GroupName $group.DisplayName }
    elseif ($action -eq 1) {
        try {
            $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
            if ($members.Count -eq 0) { Write-InfoMsg "No members."; Pause-ForUser; return }
            $labels = $members | ForEach-Object { "$($_.AdditionalProperties['displayName']) ($($_.AdditionalProperties['userPrincipalName']))" }
            $selected = Show-MultiSelect -Title "Select member(s) to remove" -Options $labels
            foreach ($idx in $selected) {
                $m = $members[$idx]
                if (Confirm-Action "Remove '$($m.AdditionalProperties['displayName'])'?") {
                    $ok = Invoke-Action -Description ("Remove {0} from group '{1}'" -f $m.AdditionalProperties['userPrincipalName'], $group.DisplayName) -Action {
                        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $m.Id -ErrorAction Stop; $true
                    }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Removed." }
                }
            }
        } catch { Write-ErrorMsg "Error: $_" }
    }
    Pause-ForUser
}

function Edit-SecurityGroupProperties {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Write-StatusLine "Display Name" $group.DisplayName "White"
    Write-StatusLine "Description" $(if ($group.Description) { $group.Description } else { "(none)" }) "White"
    Write-StatusLine "Mail Nickname" $group.MailNickname "White"
    Write-StatusLine "Mail Enabled" "$($group.MailEnabled)" "White"
    Write-StatusLine "Object ID" $group.Id "Gray"
    try { $mc = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop).Count; Write-StatusLine "Members" "$mc" "Cyan" } catch {}

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Change description","Change mail nickname") -BackLabel "Done"
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename to '$v'?")) {
                $ok = Invoke-Action -Description ("Rename group '{0}' -> '{1}'" -f $group.DisplayName, $v) -Action { Update-MgGroup -GroupId $group.Id -DisplayName $v -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Updated." } } }
        1 { $v = Read-UserInput "New description (or 'clear')"; $sv = if ($v -eq 'clear') { "" } else { $v }; if (Confirm-Action "Update description?") {
                $ok = Invoke-Action -Description ("Update group '{0}' description" -f $group.DisplayName) -Action { Update-MgGroup -GroupId $group.Id -Description $sv -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Updated." } } }
        2 { $v = Read-UserInput "New mail nickname"; if ($v -and (Confirm-Action "Change to '$v'?")) {
                $ok = Invoke-Action -Description ("Update group '{0}' mailNickname -> '{1}'" -f $group.DisplayName, $v) -Action { Update-MgGroup -GroupId $group.Id -MailNickname $v -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Updated." } } }
    }
    Pause-ForUser
}

function Remove-SecurityGroupFlow {
    $group = Find-SecurityGroup; if ($null -eq $group) { Pause-ForUser; return }
    Write-StatusLine "Group" $group.DisplayName "White"
    Write-Warn "This is irreversible!"
    if (Confirm-Action "DELETE '$($group.DisplayName)'?") {
        $check = Read-UserInput "Type the group name to confirm"
        if ($check -eq $group.DisplayName) {
            $ok = Invoke-Action -Description ("DELETE security group '{0}'" -f $group.DisplayName) -Action {
                Remove-MgGroup -GroupId $group.Id -ErrorAction Stop; $true
            }
            if ($ok -and -not (Get-PreviewMode)) { Write-Success "Deleted." }
        } else { Write-Warn "Name mismatch. Cancelled." }
    }
    Pause-ForUser
}

function Find-SecurityGroup {
    $s = Read-UserInput "Search security group by name"
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try {
        $groups = @(Get-MgGroup -Search "displayName:$s" -ConsistencyLevel eventual -ErrorAction Stop | Where-Object { $_.SecurityEnabled })
        if ($groups.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($groups.Count -eq 1) { Write-Success "Found: $($groups[0].DisplayName)"; return $groups[0] }
        $labels = $groups | ForEach-Object { $_.DisplayName }
        $sel = Show-Menu -Title "Select" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $groups[$sel]
    } catch { Write-ErrorMsg "Search error: $_"; return $null }
}

function Show-GroupMembers {
    param([string]$GroupId, [string]$GroupName)
    try {
        $members = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop)
        Write-InfoMsg "Members of '$GroupName' ($($members.Count)):"
        if ($members.Count -eq 0) { Write-InfoMsg "  (none)" }
        else { $members | ForEach-Object { Write-Host "    - $($_.AdditionalProperties['displayName']) ($($_.AdditionalProperties['userPrincipalName']))" -ForegroundColor White } }
    } catch { Write-Warn "Could not read members: $_" }
}

function Add-MembersLoop {
    param([string]$GroupId, [string]$GroupName)
    while ($true) {
        $ui = Read-UserInput "User name or email to add (or 'done')"
        if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }
                if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { continue }; $f[$sel]
                }
            }
            if (Confirm-Action "Add '$($tu.DisplayName)' to '$GroupName'?") {
                $ok = Invoke-Action -Description ("Add {0} to security group '{1}'" -f $tu.UserPrincipalName, $GroupName) -Action {
                    try { New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $tu.Id -ErrorAction Stop; $true }
                    catch { if ($_.Exception.Message -match 'already exist') { 'already' } else { throw } }
                }
                if (-not (Get-PreviewMode)) {
                    if ($ok -eq 'already') { Write-Warn "Already a member." }
                    elseif ($ok)           { Write-Success "Added." }
                }
            }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
}
