# ============================================================
#  UserProfile.ps1 - View / Edit User Profile
# ============================================================

function Start-UserProfileManagement {
    Write-SectionHeader "User Profile Management"

    if (-not (Connect-ForTask "UserProfile")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ---- Identify user ----
    $user = Resolve-UserIdentity -PromptText "Enter user name or email"
    if ($null -eq $user) { Pause-ForUser; return }

    # ---- Fetch full profile ----
    try {
        $profile = Get-AzureADUser -ObjectId $user.ObjectId -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Could not retrieve user profile: $_"
        Pause-ForUser; return
    }

    # ---- Define the editable field map ----
    # Each entry: display label -> AzureAD property name
    $fieldMap = [ordered]@{
        "Display Name"          = "DisplayName"
        "First Name"            = "GivenName"
        "Last Name"             = "Surname"
        "Job Title"             = "JobTitle"
        "Department"            = "Department"
        "Company Name"          = "CompanyName"
        "Employee ID"           = "ExtensionProperty"
        "Manager"               = "_Manager"
        "Office Location"       = "PhysicalDeliveryOfficeName"
        "Street Address"        = "StreetAddress"
        "City"                  = "City"
        "State"                 = "State"
        "Postal Code"          = "PostalCode"
        "Country"               = "Country"
        "Usage Location"        = "UsageLocation"
        "Business Phone"        = "TelephoneNumber"
        "Mobile Phone"          = "Mobile"
        "Fax Number"            = "FacsimileTelephoneNumber"
        "Mail"                  = "Mail"
        "UPN"                   = "UserPrincipalName"
        "Mail Nickname"         = "MailNickName"
        "Other Emails"          = "OtherMails"
        "Proxy Addresses"       = "ProxyAddresses"
        "Account Enabled"       = "AccountEnabled"
        "User Type"             = "UserType"
        "Creation Type"         = "CreationType"
        "Object ID"             = "ObjectId"
        "Last Dir Sync"         = "DirSyncEnabled"
    }

    # ---- Fields that can be edited with Set-AzureADUser ----
    $editableFields = @(
        "Display Name", "First Name", "Last Name", "Job Title",
        "Department", "Company Name", "Office Location",
        "Street Address", "City", "State", "Postal Code",
        "Country", "Usage Location", "Business Phone",
        "Mobile Phone", "Fax Number", "Mail Nickname"
    )

    # ---- Build display data ----
    function Get-ProfileValue {
        param($Profile, $PropertyName)

        if ($PropertyName -eq "_Manager") {
            try {
                $mgr = Get-AzureADUserManager -ObjectId $Profile.ObjectId -ErrorAction SilentlyContinue
                if ($mgr) { return "$($mgr.DisplayName) ($($mgr.UserPrincipalName))" }
                else      { return "(not set)" }
            } catch { return "(not set)" }
        }

        $val = $Profile.$PropertyName
        if ($null -eq $val)          { return "(not set)" }
        if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
            $joined = ($val | ForEach-Object { "$_" }) -join "; "
            if ([string]::IsNullOrWhiteSpace($joined)) { return "(not set)" }
            return $joined
        }
        if ([string]::IsNullOrWhiteSpace("$val")) { return "(not set)" }
        return "$val"
    }

    # ---- Show full profile ----
    function Show-FullProfile {
        param($Profile)

        $b = $script:Box
        Write-Host ""
        Write-Host ("  " + $b.DTL + [string]::new($b.DH, 2) + " Profile: $($Profile.DisplayName) " + [string]::new($b.DH, 20) + $b.DTR) -ForegroundColor $script:Colors.Title
        Write-Host ""

        $idx = 1
        foreach ($label in $fieldMap.Keys) {
            $propName = $fieldMap[$label]
            $value    = Get-ProfileValue -Profile $Profile -PropertyName $propName
            $isEditable = $editableFields -contains $label
            $numTag     = if ($isEditable) { ("{0,3}" -f $idx) } else { "   " }

            Write-Host "  " -NoNewline
            if ($isEditable) {
                Write-Host "[" -NoNewline -ForegroundColor $script:Colors.Accent
                Write-Host $numTag.Trim() -NoNewline -ForegroundColor $script:Colors.Highlight
                Write-Host "]" -NoNewline -ForegroundColor $script:Colors.Accent
            } else {
                Write-Host "   " -NoNewline
            }

            $paddedLabel = " {0,-20}" -f $label
            Write-Host $paddedLabel -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host ": " -NoNewline

            if ($value -eq "(not set)") {
                Write-Host $value -ForegroundColor "DarkGray"
            } else {
                Write-Host $value -ForegroundColor White
            }

            if ($isEditable) { $idx++ }
        }

        Write-Host ""
        Write-Host ("  " + $b.DBL + [string]::new($b.DH, 60) + $b.DBR) -ForegroundColor $script:Colors.Title
        Write-Host ""
        Write-InfoMsg "Numbered fields are editable. Unnumbered fields are read-only."
    }

    # ---- Main profile loop ----
    $keepGoing = $true
    while ($keepGoing) {

        Show-FullProfile -Profile $profile

        $action = Show-Menu -Title "Profile Actions" -Options @(
            "Edit a field",
            "Change manager",
            "Refresh profile view"
        ) -BackLabel "Done"

        switch ($action) {
            0 {
                # ---- Edit a field ----
                $fieldNum = Read-UserInput "Enter the field number to edit"

                if ($fieldNum -match '^\d+$') {
                    $fi = [int]$fieldNum
                    if ($fi -ge 1 -and $fi -le $editableFields.Count) {
                        $chosenLabel = $editableFields[$fi - 1]
                        $propName    = $fieldMap[$chosenLabel]
                        $currentVal  = Get-ProfileValue -Profile $profile -PropertyName $propName

                        Write-Host ""
                        Write-StatusLine "Field" $chosenLabel "Cyan"
                        Write-StatusLine "Current value" $currentVal "White"

                        $newVal = Read-UserInput "Enter new value (or 'clear' to blank it)"

                        if ($newVal -eq 'clear') { $newVal = $null }

                        $details = "Field: $chosenLabel`nOld: $currentVal`nNew: $(if ($null -eq $newVal) { '(cleared)' } else { $newVal })"
                        if (Confirm-Action "Update this field?" $details) {
                            try {
                                $params = @{ ObjectId = $profile.ObjectId }
                                if ($null -eq $newVal -or $newVal -eq '') {
                                    # Set to empty/null
                                    $params[$propName] = $null
                                } else {
                                    $params[$propName] = $newVal
                                }
                                Set-AzureADUser @params -ErrorAction Stop
                                Write-Success "'$chosenLabel' updated."

                                # Refresh local copy
                                $profile = Get-AzureADUser -ObjectId $profile.ObjectId -ErrorAction Stop
                            } catch {
                                Write-ErrorMsg "Failed to update: $_"
                            }
                        }
                    } else {
                        Write-ErrorMsg "Invalid field number. Must be 1-$($editableFields.Count)."
                    }
                } else {
                    Write-ErrorMsg "Please enter a number."
                }
                Pause-ForUser
            }
            1 {
                # ---- Change manager ----
                Write-Host ""
                try {
                    $currentMgr = Get-AzureADUserManager -ObjectId $profile.ObjectId -ErrorAction SilentlyContinue
                    if ($currentMgr) {
                        Write-StatusLine "Current Manager" "$($currentMgr.DisplayName) ($($currentMgr.UserPrincipalName))" "White"
                    } else {
                        Write-StatusLine "Current Manager" "(not set)" "DarkGray"
                    }
                } catch {
                    Write-StatusLine "Current Manager" "(not set)" "DarkGray"
                }

                $mgrAction = Show-Menu -Title "Manager Action" -Options @(
                    "Set / change manager",
                    "Remove manager"
                ) -BackLabel "Cancel"

                if ($mgrAction -eq 0) {
                    $mgrInput = Read-UserInput "Enter new manager name or email"
                    if (-not [string]::IsNullOrWhiteSpace($mgrInput)) {
                        try {
                            if ($mgrInput -match '@') {
                                $newMgr = Get-AzureADUser -ObjectId $mgrInput -ErrorAction Stop
                            } else {
                                $found = Get-AzureADUser -SearchString $mgrInput -ErrorAction Stop
                                if ($found.Count -eq 0) {
                                    Write-ErrorMsg "No user found matching '$mgrInput'."
                                    Pause-ForUser; continue
                                }
                                if ($found.Count -eq 1) {
                                    $newMgr = $found[0]
                                } else {
                                    $names = $found | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                                    $sel = Show-Menu -Title "Select Manager" -Options $names -BackLabel "Cancel"
                                    if ($sel -eq -1) { Pause-ForUser; continue }
                                    $newMgr = $found[$sel]
                                }
                            }

                            Write-StatusLine "New Manager" "$($newMgr.DisplayName) ($($newMgr.UserPrincipalName))" "White"
                            if (Confirm-Action "Set $($newMgr.DisplayName) as manager for $($profile.DisplayName)?") {
                                Set-AzureADUserManager -ObjectId $profile.ObjectId `
                                    -RefObjectId $newMgr.ObjectId -ErrorAction Stop
                                Write-Success "Manager updated to $($newMgr.DisplayName)."
                            }
                        } catch {
                            Write-ErrorMsg "Failed to set manager: $_"
                        }
                    }
                }
                elseif ($mgrAction -eq 1) {
                    if (Confirm-Action "Remove manager from $($profile.DisplayName)?") {
                        try {
                            Remove-AzureADUserManager -ObjectId $profile.ObjectId -ErrorAction Stop
                            Write-Success "Manager removed."
                        } catch {
                            Write-ErrorMsg "Failed to remove manager: $_"
                        }
                    }
                }
                Pause-ForUser
            }
            2 {
                # ---- Refresh ----
                try {
                    $profile = Get-AzureADUser -ObjectId $profile.ObjectId -ErrorAction Stop
                    Write-Success "Profile refreshed."
                } catch {
                    Write-ErrorMsg "Could not refresh profile: $_"
                }
                Pause-ForUser
            }
            -1 {
                $keepGoing = $false
            }
        }
    }

    Write-Success "Profile management complete."
}
