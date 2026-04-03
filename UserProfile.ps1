# ============================================================
#  UserProfile.ps1 - View / Edit User Profile (Microsoft Graph)
# ============================================================

function Start-UserProfileManagement {
    Write-SectionHeader "User Profile Management"
    if (-not (Connect-ForTask "UserProfile")) { Pause-ForUser; return }

    $user = Resolve-UserIdentity
    if ($null -eq $user) { Pause-ForUser; return }

    # Editable fields: label -> Graph property
    $editableFields = [ordered]@{
        "Display Name"    = "DisplayName"
        "First Name"      = "GivenName"
        "Last Name"       = "Surname"
        "Job Title"       = "JobTitle"
        "Department"      = "Department"
        "Company Name"    = "CompanyName"
        "Office Location" = "OfficeLocation"
        "Street Address"  = "StreetAddress"
        "City"            = "City"
        "State"           = "State"
        "Postal Code"     = "PostalCode"
        "Country"         = "Country"
        "Usage Location"  = "UsageLocation"
        "Mobile Phone"    = "MobilePhone"
        "Mail Nickname"   = "MailNickname"
    }

    # Read-only fields
    $readOnlyFields = [ordered]@{
        "UPN"             = "UserPrincipalName"
        "Mail"            = "Mail"
        "Business Phones" = "BusinessPhones"
        "Account Enabled" = "AccountEnabled"
        "User Type"       = "UserType"
        "Object ID"       = "Id"
    }

    $keepGoing = $true
    while ($keepGoing) {
        # Refresh profile
        try {
            $profile = Get-MgUser -UserId $user.Id -Property ($editableFields.Values + $readOnlyFields.Values + @("Id")) -ErrorAction Stop
        } catch { Write-ErrorMsg "Could not load profile: $_"; Pause-ForUser; return }

        # Manager
        $mgrDisplay = "(not set)"
        try {
            $mgr = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
            if ($mgr) { $mgrDisplay = "$($mgr.AdditionalProperties['displayName']) ($($mgr.AdditionalProperties['userPrincipalName']))" }
        } catch {}

        # Display
        $b = $script:Box
        Write-Host ""
        Write-Host ("  " + $b.DTL + [string]::new($b.DH, 2) + " Profile: $($profile.DisplayName) " + [string]::new($b.DH, 20) + $b.DTR) -ForegroundColor $script:Colors.Title
        Write-Host ""

        $idx = 1
        foreach ($label in $editableFields.Keys) {
            $prop = $editableFields[$label]
            $val = $profile.$prop
            $valStr = if ($null -eq $val -or ([string]::IsNullOrWhiteSpace("$val"))) { "(not set)" } else { "$val" }
            Write-Host "  [" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host ("{0,2}" -f $idx) -NoNewline -ForegroundColor $script:Colors.Highlight
            Write-Host "]" -NoNewline -ForegroundColor $script:Colors.Accent
            Write-Host (" {0,-18}" -f $label) -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host ": " -NoNewline
            Write-Host $valStr -ForegroundColor $(if ($valStr -eq "(not set)") { "DarkGray" } else { "White" })
            $idx++
        }

        # Manager row (editable via separate menu)
        Write-Host "  [" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host " M" -NoNewline -ForegroundColor $script:Colors.Highlight
        Write-Host "]" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host (" {0,-18}" -f "Manager") -NoNewline -ForegroundColor $script:Colors.Info
        Write-Host ": " -NoNewline
        Write-Host $mgrDisplay -ForegroundColor $(if ($mgrDisplay -eq "(not set)") { "DarkGray" } else { "White" })

        Write-Host ""
        # Read-only
        foreach ($label in $readOnlyFields.Keys) {
            $prop = $readOnlyFields[$label]
            $val = $profile.$prop
            if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) { $valStr = ($val -join "; ") }
            else { $valStr = if ($null -eq $val -or ([string]::IsNullOrWhiteSpace("$val"))) { "(not set)" } else { "$val" } }
            Write-Host "     " -NoNewline
            Write-Host (" {0,-18}" -f $label) -NoNewline -ForegroundColor $script:Colors.Info
            Write-Host ": " -NoNewline
            Write-Host $valStr -ForegroundColor "Gray"
        }

        Write-Host ""
        Write-Host ("  " + $b.DBL + [string]::new($b.DH, 60) + $b.DBR) -ForegroundColor $script:Colors.Title
        Write-InfoMsg "Numbered fields are editable. Enter number, 'M' for manager, or 'done'."

        $input = Read-UserInput "Action"

        if ($input -match '^done$') { $keepGoing = $false; continue }

        if ($input -match '^[Mm]$') {
            # Manager management
            Write-StatusLine "Current Manager" $mgrDisplay "White"
            $ma = Show-Menu -Title "Manager" -Options @("Set / change","Remove") -BackLabel "Cancel"
            if ($ma -eq 0) {
                $mi = Read-UserInput "New manager name or email"
                if ($mi) {
                    try {
                        $nm = if ($mi -match '@') { Get-MgUser -UserId $mi -ErrorAction Stop } else {
                            $f = @(Get-MgUser -Search "displayName:$mi" -ConsistencyLevel eventual -ErrorAction Stop)
                            if ($f.Count -eq 0) { Write-ErrorMsg "Not found."; Pause-ForUser; continue }
                            if ($f.Count -eq 1) { $f[0] } else {
                                $sel = Show-Menu -Title "Select" -Options ($f | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }) -BackLabel "Cancel"
                                if ($sel -eq -1) { Pause-ForUser; continue }; $f[$sel]
                            }
                        }
                        if (Confirm-Action "Set manager to $($nm.DisplayName)?") {
                            Set-MgUserManagerByRef -UserId $user.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($nm.Id)" } -ErrorAction Stop
                            Write-Success "Manager set."
                        }
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            } elseif ($ma -eq 1) {
                if (Confirm-Action "Remove manager?") {
                    try { Remove-MgUserManagerByRef -UserId $user.Id -ErrorAction Stop; Write-Success "Removed." }
                    catch { Write-ErrorMsg "Failed: $_" }
                }
            }
            Pause-ForUser; continue
        }

        if ($input -match '^\d+$') {
            $fi = [int]$input
            $editKeys = @($editableFields.Keys)
            if ($fi -ge 1 -and $fi -le $editKeys.Count) {
                $label = $editKeys[$fi - 1]
                $prop = $editableFields[$label]
                $curVal = $profile.$prop
                Write-StatusLine "Field" $label "Cyan"
                Write-StatusLine "Current" $(if ($curVal) { "$curVal" } else { "(not set)" }) "White"
                $newVal = Read-UserInput "New value (or 'clear')"
                if ($newVal -eq 'clear') { $newVal = "" }
                if (Confirm-Action "Update '$label'?" "Old: $(if ($curVal) { $curVal } else { '(empty)' })`nNew: $(if ($newVal) { $newVal } else { '(cleared)' })") {
                    try {
                        $body = @{}; $body[$prop] = if ([string]::IsNullOrWhiteSpace($newVal)) { $null } else { $newVal }
                        Update-MgUser -UserId $user.Id -BodyParameter $body -ErrorAction Stop
                        Write-Success "'$label' updated."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            } else { Write-ErrorMsg "Invalid number." }
            Pause-ForUser; continue
        }

        Write-ErrorMsg "Enter a field number, 'M', or 'done'."
        Pause-ForUser
    }
    Write-Success "Profile management complete."
}
