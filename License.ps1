# ============================================================
#  License.ps1 - License Management (Microsoft Graph)
#  Shows friendly names, detects group-assigned licenses
# ============================================================

function Start-LicenseManagement {
    Write-SectionHeader "License Management"
    if (-not (Connect-ForTask "License")) { Pause-ForUser; return }

    $user = Resolve-UserIdentity
    if ($null -eq $user) { Pause-ForUser; return }

    # ---- Get license assignment states (direct vs group) ----
    $fullUser = $null
    try {
        $fullUser = Get-MgUser -UserId $user.Id -Property "Id,DisplayName,UserPrincipalName,LicenseAssignmentStates" -ErrorAction Stop
    } catch {
        Write-Warn "Could not read assignment states: $_"
    }

    # Build a lookup: SkuId -> assignment info
    $assignmentInfo = @{}
    if ($fullUser -and $fullUser.LicenseAssignmentStates) {
        foreach ($state in $fullUser.LicenseAssignmentStates) {
            $skuId = "$($state.SkuId)"
            $info = @{ Direct = $false; Groups = @() }
            if ($assignmentInfo.ContainsKey($skuId)) { $info = $assignmentInfo[$skuId] }

            if ($null -eq $state.AssignedByGroup -or $state.AssignedByGroup -eq "") {
                $info.Direct = $true
            } else {
                $info.Groups += $state.AssignedByGroup
            }
            $assignmentInfo[$skuId] = $info
        }
    }

    # ---- Show current licenses ----
    Write-SectionHeader "Current Licenses for $($user.DisplayName)"

    try { $currentLics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop) } catch { $currentLics = @() }

    if ($currentLics.Count -eq 0) {
        Write-InfoMsg "No licenses assigned."
    } else {
        for ($i = 0; $i -lt $currentLics.Count; $i++) {
            $lic = $currentLics[$i]
            $label = Format-LicenseLabel $lic.SkuPartNumber

            # Check assignment type
            $skuId = "$($lic.SkuId)"
            $assignTag = ""
            if ($assignmentInfo.ContainsKey($skuId)) {
                $ai = $assignmentInfo[$skuId]
                if ($ai.Direct -and $ai.Groups.Count -gt 0) {
                    $assignTag = " [Direct + Group]"
                } elseif ($ai.Groups.Count -gt 0) {
                    $assignTag = " [Group-assigned]"
                } else {
                    $assignTag = " [Direct]"
                }
            }

            Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host ($i + 1) -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host $label -NoNewline -ForegroundColor White
            if ($assignTag -match "Group") {
                Write-Host $assignTag -ForegroundColor $script:Colors.Warning
            } else {
                Write-Host $assignTag -ForegroundColor $script:Colors.Info
            }
        }
    }
    Write-Host ""

    # ---- Add or Remove ----
    $action = Show-Menu -Title "Action" -Options @("Add license(s)", "Remove license(s)") -BackLabel "Cancel"
    if ($action -eq -1) { return }

    if ($action -eq 0) {
        # ---- ADD ----
        Write-SectionHeader "Available Tenant Licenses"
        try {
            $skus = Get-MgSubscribedSku -ErrorAction Stop
            $available = $skus | ForEach-Object {
                $t = $_.PrepaidUnits.Enabled; $u = $_.ConsumedUnits
                [PSCustomObject]@{
                    SkuPartNumber = $_.SkuPartNumber
                    SkuId         = $_.SkuId
                    FriendlyName  = Format-LicenseLabel $_.SkuPartNumber
                    Total         = $t
                    Used          = $u
                    Free          = $t - $u
                }
            }

            $labels = $available | ForEach-Object {
                "$($_.FriendlyName)  [Total: $($_.Total) | Used: $($_.Used) | Free: $($_.Free)]"
            }

            $selected = Show-MultiSelect -Title "Select license(s) to add" -Options $labels
            foreach ($idx in $selected) {
                $sku = $available[$idx]
                if ($sku.Free -le 0) {
                    Write-Warn "$(Get-SkuFriendlyName $sku.SkuPartNumber) has no available seats. Skipping."
                    continue
                }
                if (Confirm-Action "Assign '$(Get-SkuFriendlyName $sku.SkuPartNumber)' to $($user.UserPrincipalName)?") {
                    $ok = Invoke-Action -Description ("Assign license '{0}' to {1}" -f $sku.SkuPartNumber, $user.UserPrincipalName) -Action {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @() -ErrorAction Stop; $true
                    }
                    if ($ok -and -not (Get-PreviewMode)) { Write-Success "$(Get-SkuFriendlyName $sku.SkuPartNumber) assigned." }
                }
            }
        } catch { Write-ErrorMsg "License error: $_" }
    }
    else {
        # ---- REMOVE ----
        if ($currentLics.Count -eq 0) {
            Write-Warn "No licenses to remove."
            Pause-ForUser; return
        }

        $labels = @()
        for ($i = 0; $i -lt $currentLics.Count; $i++) {
            $lic = $currentLics[$i]
            $label = Format-LicenseLabel $lic.SkuPartNumber
            $skuId = "$($lic.SkuId)"
            if ($assignmentInfo.ContainsKey($skuId) -and $assignmentInfo[$skuId].Groups.Count -gt 0 -and -not $assignmentInfo[$skuId].Direct) {
                $label += "  !! GROUP-ASSIGNED"
            }
            $labels += $label
        }

        $selected = Show-MultiSelect -Title "Select license(s) to remove" -Options $labels

        foreach ($idx in $selected) {
            $lic = $currentLics[$idx]
            $friendlyName = Get-SkuFriendlyName $lic.SkuPartNumber
            $skuId = "$($lic.SkuId)"

            # Check if group-assigned only
            if ($assignmentInfo.ContainsKey($skuId)) {
                $ai = $assignmentInfo[$skuId]
                if ($ai.Groups.Count -gt 0 -and -not $ai.Direct) {
                    Write-Host ""
                    Write-ErrorMsg "'$friendlyName' is assigned via a group, not directly."
                    Write-Warn "You cannot remove group-assigned licenses from individual users."
                    Write-InfoMsg "To remove this license:"
                    Write-InfoMsg "  1. Remove the user from the licensing group, OR"
                    Write-InfoMsg "  2. Remove this license from the group itself."

                    # Try to resolve group name
                    foreach ($gId in $ai.Groups) {
                        try {
                            $grp = Get-MgGroup -GroupId $gId -Property "DisplayName" -ErrorAction Stop
                            Write-InfoMsg "  Assigned by group: $($grp.DisplayName) ($gId)"
                        } catch {
                            Write-InfoMsg "  Assigned by group ID: $gId"
                        }
                    }
                    Write-Host ""
                    continue
                }
            }

            if (Confirm-Action "Remove '$friendlyName' from $($user.UserPrincipalName)?") {
                $ok = Invoke-Action -Description ("Remove license '{0}' from {1}" -f $lic.SkuPartNumber, $user.UserPrincipalName) -Action {
                    try {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; $true
                    } catch {
                        $errMsg = $_.Exception.Message
                        if ($errMsg -match "group-based|inherited|cannot remove") {
                            'inherited'
                        } else { throw }
                    }
                }
                if (-not (Get-PreviewMode)) {
                    if ($ok -eq 'inherited') {
                        Write-ErrorMsg "Cannot remove: this license is inherited from a group."
                        Write-InfoMsg "Remove the user from the licensing group instead."
                    } elseif ($ok) {
                        Write-Success "$friendlyName removed."
                    }
                }
            }
        }
    }

    Write-Success "License management complete."
    Pause-ForUser
}
