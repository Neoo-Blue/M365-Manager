# ============================================================
#  DistributionList.ps1 - DL Management (MS Graph + EXO)
# ============================================================

function Start-DistributionListManagement {
    Write-SectionHeader "Distribution List Management"
    if (-not (Connect-ForTask "DistributionList")) { Pause-ForUser; return }

    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Create a new distribution list","Add / remove members",
        "View / edit DL properties","Delete a distribution list"
    ) -BackLabel "Back to Main Menu"

    switch ($action) {
        0 { New-DistributionListFlow }
        1 { Edit-DistributionListMembers }
        2 { Edit-DistributionListProperties }
        3 { Remove-DistributionListFlow }
    }
}

function New-DistributionListFlow {
    Write-SectionHeader "Create New Distribution List"
    $name = Read-UserInput "Display name"; if ([string]::IsNullOrWhiteSpace($name)) { Pause-ForUser; return }
    $alias = Read-UserInput "Email alias (e.g. sales-team)"; if ([string]::IsNullOrWhiteSpace($alias)) { $alias = ($name -replace '[^a-zA-Z0-9-]','').ToLower() }
    $smtp = Read-UserInput "Full primary email (e.g. sales@contoso.com)"; if ([string]::IsNullOrWhiteSpace($smtp)) { Pause-ForUser; return }
    $owner = Read-UserInput "Owner email (or Enter to skip)"
    $ra = Show-Menu -Title "Who can send?" -Options @("Anyone","Internal only") -BackLabel "Cancel"; if ($ra -eq -1) { return }

    if (Confirm-Action "Create DL '$name' ($smtp)?") {
        try {
            $p = @{ Name = $name; DisplayName = $name; Alias = $alias; PrimarySmtpAddress = $smtp; Type = "Distribution"; RequireSenderAuthenticationEnabled = ($ra -eq 1) }
            if ($owner) { $p["ManagedBy"] = $owner }
            New-DistributionGroup @p -ErrorAction Stop | Out-Null; Write-Success "Created."
            $add = Read-UserInput "Add members now? (y/n)"; if ($add -match '^[Yy]') { Add-DLMembersLoop -Id $smtp -Name $name }
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Edit-DistributionListMembers {
    $user = Resolve-UserIdentity; if ($null -eq $user) { Pause-ForUser; return }
    $upn = $user.UserPrincipalName
    $action = Show-Menu -Title "Action" -Options @("Add to DL(s)","Remove from DL(s)") -BackLabel "Cancel"; if ($action -eq -1) { return }

    if ($action -eq 0) {
        $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
        if (Confirm-Action "Add to '$($dl.DisplayName)'?") {
            $ok = Invoke-Action -Description ("Add {0} to DL '{1}'" -f $upn, $dl.DisplayName) -Action {
                try { Add-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $upn -ErrorAction Stop; $true }
                catch { if ($_.Exception.Message -match 'already') { 'already' } else { throw } }
            }
            if (-not (Get-PreviewMode)) {
                if ($ok -eq 'already') { Write-Warn "Already a member." }
                elseif ($ok)           { Write-Success "Added." }
            }
        }
        $pc = Show-Menu -Title "Send permissions?" -Options @("Send As","Send on Behalf","Both","None") -BackLabel "Skip"
        if ($pc -ne -1 -and $pc -ne 3) {
            if ($pc -eq 0 -or $pc -eq 2) {
                if (Confirm-Action "Grant Send As?") {
                    $ok = Invoke-Action -Description ("Grant {0} SendAs on DL '{1}'" -f $upn, $dl.DisplayName) -Action {
                        Add-RecipientPermission -Identity $dl.PrimarySmtpAddress -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop; $true
                    }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Granted." }
                }
            }
            if ($pc -eq 1 -or $pc -eq 2) {
                if (Confirm-Action "Grant Send on Behalf?") {
                    $ok = Invoke-Action -Description ("Grant {0} SendOnBehalf on DL '{1}'" -f $upn, $dl.DisplayName) -Action {
                        Set-DistributionGroup -Identity $dl.PrimarySmtpAddress -GrantSendOnBehalfTo @{Add=$upn} -ErrorAction Stop; $true
                    }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Granted." }
                }
            }
        }
    } else {
        Write-InfoMsg "Finding DL memberships..."
        try {
            $allDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
            $memberOf = @(); foreach ($dl in $allDLs) {
                $ms = Get-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -ErrorAction SilentlyContinue
                if ($ms | Where-Object { $_.PrimarySmtpAddress -eq $upn }) { $memberOf += $dl }
            }
            if ($memberOf.Count -eq 0) { Write-InfoMsg "Not in any DLs."; Pause-ForUser; return }
            $labels = $memberOf | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }
            $sel = Show-MultiSelect -Title "Remove from" -Options $labels
            foreach ($idx in $sel) {
                $dl = $memberOf[$idx]
                if (Confirm-Action "Remove from '$($dl.DisplayName)'?") {
                    $ok = Invoke-Action -Description ("Remove {0} from DL '{1}'" -f $upn, $dl.DisplayName) -Action {
                        Remove-DistributionGroupMember -Identity $dl.PrimarySmtpAddress -Member $upn -Confirm:$false -ErrorAction Stop; $true
                    }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "Removed." }
                }
            }
        } catch { Write-ErrorMsg "$_" }
    }
    Pause-ForUser
}

function Edit-DistributionListProperties {
    $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
    try { $dl = Get-DistributionGroup -Identity $dl.PrimarySmtpAddress -ErrorAction Stop } catch {}
    Write-StatusLine "Name" $dl.DisplayName "White"; Write-StatusLine "Email" $dl.PrimarySmtpAddress "White"
    Write-StatusLine "Description" $(if ($dl.Description) { $dl.Description } else { "(none)" }) "White"
    Write-StatusLine "Managed By" ($dl.ManagedBy -join "; ") "White"
    Write-StatusLine "Auth Required" "$($dl.RequireSenderAuthenticationEnabled)" "White"
    Write-StatusLine "Hidden" "$($dl.HiddenFromAddressListsEnabled)" "White"

    $ec = Show-Menu -Title "Edit" -Options @("Change name","Change description","Change owner","Toggle sender auth","Toggle hidden") -BackLabel "Done"
    $dlId = $dl.PrimarySmtpAddress
    $dlName = $dl.DisplayName
    switch ($ec) {
        0 { $v = Read-UserInput "New name"; if ($v -and (Confirm-Action "Rename?")) {
                $ok = Invoke-Action -Description ("Rename DL '{0}' -> '{1}'" -f $dlName, $v) -Action { Set-DistributionGroup -Identity $dlId -DisplayName $v -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        1 { $v = Read-UserInput "New description (or 'clear')"; if (Confirm-Action "Update?") {
                $nv = if ($v -eq 'clear') { "" } else { $v }
                $ok = Invoke-Action -Description ("Update DL '{0}' description" -f $dlName) -Action { Set-DistributionGroup -Identity $dlId -Description $nv -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        2 { $v = Read-UserInput "New owner email"; if ($v -and (Confirm-Action "Set owner?")) {
                $ok = Invoke-Action -Description ("Set DL '{0}' owner -> {1}" -f $dlName, $v) -Action { Set-DistributionGroup -Identity $dlId -ManagedBy $v -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        3 { $nv = -not $dl.RequireSenderAuthenticationEnabled; if (Confirm-Action "Set to $nv?") {
                $ok = Invoke-Action -Description ("Set DL '{0}' RequireSenderAuthenticationEnabled = {1}" -f $dlName, $nv) -Action { Set-DistributionGroup -Identity $dlId -RequireSenderAuthenticationEnabled $nv -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
        4 { $nv = -not $dl.HiddenFromAddressListsEnabled; if (Confirm-Action "Set hidden to $nv?") {
                $ok = Invoke-Action -Description ("Set DL '{0}' HiddenFromAddressListsEnabled = {1}" -f $dlName, $nv) -Action { Set-DistributionGroup -Identity $dlId -HiddenFromAddressListsEnabled $nv -ErrorAction Stop; $true }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Done." } } }
    }
    Pause-ForUser
}

function Remove-DistributionListFlow {
    $dl = Find-DistributionList; if ($null -eq $dl) { Pause-ForUser; return }
    Write-Warn "Irreversible!"
    if (Confirm-Action "DELETE '$($dl.DisplayName)'?") {
        $check = Read-UserInput "Type the DL email to confirm"
        if ($check -eq $dl.PrimarySmtpAddress) {
            $ok = Invoke-Action -Description ("DELETE DL '{0}' <{1}>" -f $dl.DisplayName, $dl.PrimarySmtpAddress) -Action {
                Remove-DistributionGroup -Identity $dl.PrimarySmtpAddress -Confirm:$false -ErrorAction Stop; $true
            }
            if ($ok -and -not (Get-PreviewMode)) { Write-Success "Deleted." }
        } else { Write-Warn "Mismatch. Cancelled." }
    }
    Pause-ForUser
}

function Find-DistributionList {
    $sm = Show-Menu -Title "Find DL by" -Options @("Name","Email") -BackLabel "Cancel"; if ($sm -eq -1) { return $null }
    $si = Read-UserInput $(if ($sm -eq 0) { "DL name" } else { "DL email" }); if ([string]::IsNullOrWhiteSpace($si)) { return $null }
    try {
        $dls = @(if ($sm -eq 0) { Get-DistributionGroup -Filter "DisplayName -like '*$si*'" -ResultSize 50 } else { Get-DistributionGroup -Filter "PrimarySmtpAddress -like '*$si*'" -ResultSize 50 })
        if ($dls.Count -eq 0) { Write-ErrorMsg "None found."; return $null }
        if ($dls.Count -eq 1) { return $dls[0] }
        $sel = Show-Menu -Title "Select" -Options ($dls | ForEach-Object { "$($_.DisplayName) ($($_.PrimarySmtpAddress))" }) -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }; return $dls[$sel]
    } catch { Write-ErrorMsg "$_"; return $null }
}

function Add-DLMembersLoop {
    param([string]$Id, [string]$Name)
    while ($true) {
        $ui = Read-UserInput "User to add (or 'done')"; if ($ui -match '^done$') { break }
        try {
            $tu = if ($ui -match '@') { Get-MgUser -UserId $ui -ErrorAction Stop } else {
                $f = @(Get-MgUser -Search "displayName:$ui" -ConsistencyLevel eventual -ErrorAction Stop)
                if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; continue }; if ($f.Count -eq 1) { $f[0] } else {
                    $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                    if ($sel -eq -1) { continue }; $f[$sel]
                }
            }
            if (Confirm-Action "Add '$($tu.DisplayName)'?") {
                $ok = Invoke-Action -Description ("Add {0} to DL '{1}'" -f $tu.UserPrincipalName, $Name) -Action {
                    Add-DistributionGroupMember -Identity $Id -Member $tu.UserPrincipalName -ErrorAction Stop; $true
                }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Added." }
            }
        } catch { if ($_.Exception.Message -match "already") { Write-Warn "Already a member." } else { Write-ErrorMsg "$_" } }
    }
}
