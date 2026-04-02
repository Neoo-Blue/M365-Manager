# ============================================================
#  Onboard.ps1 - New User Onboarding Workflow
# ============================================================

function Start-Onboard {
    Write-SectionHeader "User Onboarding"

    # ---- Profile fields ----
    $fields = @(
        "FirstName","LastName","DisplayName","UserPrincipalName",
        "JobTitle","Department","CompanyName","OfficeLocation",
        "StreetAddress","City","State","PostalCode","Country",
        "UsageLocation","BusinessPhone","MobilePhone","Manager"
    )

    # ---- Replicate-related state ----
    $replicateSource     = $null
    $replicateLicenses   = @()
    $replicateGroups     = @()

    $choice = Show-Menu -Title "How to provide user data?" -Options @(
        "Parse from a text file",
        "Enter manually",
        "Replicate from an existing user"
    ) -BackLabel "Cancel"

    if ($choice -eq -1) { return }

    $userData = @{}

    # ------------------------------------------------------------------
    #  Option 1: Parse from file
    # ------------------------------------------------------------------
    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) {
            Write-ErrorMsg "File not found: $filePath"
            Pause-ForUser; return
        }

        Write-InfoMsg "Parsing file..."
        $lines = Get-Content $filePath
        foreach ($line in $lines) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key   = $Matches[1].Trim()
                $value = $Matches[2].Trim()
                switch -Regex ($key) {
                    'first\s*name'              { $userData["FirstName"]         = $value }
                    'last\s*name'               { $userData["LastName"]          = $value }
                    'display\s*name'            { $userData["DisplayName"]       = $value }
                    'upn|email|user\s*principal' { $userData["UserPrincipalName"] = $value }
                    'title|job'                 { $userData["JobTitle"]          = $value }
                    'depart'                    { $userData["Department"]        = $value }
                    'company'                   { $userData["CompanyName"]       = $value }
                    'office'                    { $userData["OfficeLocation"]    = $value }
                    'street'                    { $userData["StreetAddress"]     = $value }
                    'city'                      { $userData["City"]              = $value }
                    'state'                     { $userData["State"]             = $value }
                    'postal|zip'                { $userData["PostalCode"]        = $value }
                    'country'                   { $userData["Country"]           = $value }
                    'usage\s*location'          { $userData["UsageLocation"]     = $value }
                    'business\s*phone'          { $userData["BusinessPhone"]     = $value }
                    'mobile'                    { $userData["MobilePhone"]       = $value }
                    'location|usage'            { $userData["UsageLocation"]     = $value }
                    'manager'                   { $userData["Manager"]           = $value }
                    default                     { $userData[$key]                = $value }
                }
            }
        }

        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }

        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and
            $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }

        Write-Success "Parsed data:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    # ------------------------------------------------------------------
    #  Option 2: Manual entry
    # ------------------------------------------------------------------
    elseif ($choice -eq 1) {
        foreach ($f in $fields) {
            $userData[$f] = Read-UserInput "Enter $f"
        }
        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and
            $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }
        Write-InfoMsg "Review the information:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    # ------------------------------------------------------------------
    #  Option 3: Replicate from existing user
    # ------------------------------------------------------------------
    elseif ($choice -eq 2) {

        # Need AAD/MSOL early to look up the source user
        if (-not (Connect-ForTask "Onboard")) {
            Write-ErrorMsg "Could not connect to required services."
            Pause-ForUser; return
        }

        Write-SectionHeader "Select User to Replicate From"
        $replicateSource = Resolve-UserIdentity -PromptText "Enter the existing user's name or email"
        if ($null -eq $replicateSource) {
            Write-ErrorMsg "No source user selected."
            Pause-ForUser; return
        }

        $src = $replicateSource

        # ---- Copy profile fields ----
        Write-InfoMsg "Copying profile from $($src.DisplayName)..."

        $userData["FirstName"]         = ""
        $userData["LastName"]          = ""
        $userData["DisplayName"]       = ""
        $userData["UserPrincipalName"] = ""
        $userData["JobTitle"]          = if ($src.JobTitle)                       { $src.JobTitle } else { "" }
        $userData["Department"]        = if ($src.Department)                     { $src.Department } else { "" }
        $userData["CompanyName"]       = if ($src.CompanyName)                    { $src.CompanyName } else { "" }
        $userData["OfficeLocation"]    = if ($src.PhysicalDeliveryOfficeName)     { $src.PhysicalDeliveryOfficeName } else { "" }
        $userData["StreetAddress"]     = if ($src.StreetAddress)                  { $src.StreetAddress } else { "" }
        $userData["City"]              = if ($src.City)                           { $src.City } else { "" }
        $userData["State"]             = if ($src.State)                          { $src.State } else { "" }
        $userData["PostalCode"]        = if ($src.PostalCode)                     { $src.PostalCode } else { "" }
        $userData["Country"]           = if ($src.Country)                        { $src.Country } else { "" }
        $userData["UsageLocation"]     = if ($src.UsageLocation)                  { $src.UsageLocation } else { "" }
        $userData["BusinessPhone"]     = if ($src.TelephoneNumber)               { $src.TelephoneNumber } else { "" }
        $userData["MobilePhone"]       = if ($src.Mobile)                        { $src.Mobile } else { "" }

        # Manager
        try {
            $srcMgr = Get-AzureADUserManager -ObjectId $src.ObjectId -ErrorAction SilentlyContinue
            if ($srcMgr) {
                $userData["Manager"] = $srcMgr.UserPrincipalName
            } else { $userData["Manager"] = "" }
        } catch { $userData["Manager"] = "" }

        # ---- Collect source licenses ----
        Write-InfoMsg "Reading licenses..."
        try {
            $srcMsol = Get-MsolUser -UserPrincipalName $src.UserPrincipalName -ErrorAction Stop
            $replicateLicenses = @($srcMsol.Licenses)
        } catch {
            Write-Warn "Could not read licenses from source: $_"
            $replicateLicenses = @()
        }

        # ---- Collect source security groups ----
        Write-InfoMsg "Reading security group memberships..."
        try {
            $replicateGroups = @(Get-AzureADUserMembership -ObjectId $src.ObjectId -All $true |
                Where-Object { $_.ObjectType -eq "Group" -and $_.SecurityEnabled -eq $true })
        } catch {
            Write-Warn "Could not read security groups from source: $_"
            $replicateGroups = @()
        }

        # ---- Display replicated summary ----
        Write-SectionHeader "Replicated from $($src.DisplayName)"

        Write-Success "Profile fields copied (name/email left blank for you to fill in)."
        Write-Host ""

        if ($replicateLicenses.Count -gt 0) {
            Write-InfoMsg "Licenses to replicate:"
            foreach ($lic in $replicateLicenses) {
                Write-Host "    - $($lic.AccountSkuId)" -ForegroundColor White
            }
        } else {
            Write-InfoMsg "No licenses to replicate."
        }
        Write-Host ""

        if ($replicateGroups.Count -gt 0) {
            Write-InfoMsg "Security groups to replicate:"
            foreach ($grp in $replicateGroups) {
                Write-Host "    - $($grp.DisplayName)" -ForegroundColor White
            }
        } else {
            Write-InfoMsg "No security groups to replicate."
        }
        Write-Host ""

        # ---- Let user fill in name/email and review everything ----
        Write-InfoMsg "Now fill in the new user's identity fields and review all data:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }

    # ---- Validate required fields ----
    $required = @("FirstName","LastName","UserPrincipalName","UsageLocation")
    foreach ($r in $required) {
        if ([string]::IsNullOrWhiteSpace($userData[$r])) {
            Write-ErrorMsg "'$r' is required but empty."
            Pause-ForUser; return
        }
    }

    # Auto-generate DisplayName if still empty
    if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and
        $userData["FirstName"] -and $userData["LastName"]) {
        $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
    }

    # ---- Connect services (if not already connected by replicate path) ----
    if (-not (Connect-ForTask "Onboard")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ==================================================================
    #  STEP 1 : Create Account
    # ==================================================================
    Write-SectionHeader "Step 1 - Create Account"

    $password = -join ((48..57) + (65..90) + (97..122) + (33,35,36,37,38) |
        Get-Random -Count 16 | ForEach-Object { [char]$_ })

    $details = @"
  UPN          : $($userData['UserPrincipalName'])
  Display Name : $($userData['DisplayName'])
  Department   : $($userData['Department'])
  Job Title    : $($userData['JobTitle'])
  Company      : $($userData['CompanyName'])
  Office       : $($userData['OfficeLocation'])
"@

    if (-not (Confirm-Action "Create the following user account?" $details)) {
        Write-Warn "Onboarding cancelled."; Pause-ForUser; return
    }

    try {
        $secPwd = ConvertTo-SecureString $password -AsPlainText -Force

        $params = @{
            UserPrincipalName   = $userData["UserPrincipalName"]
            DisplayName         = $userData["DisplayName"]
            FirstName           = $userData["FirstName"]
            LastName            = $userData["LastName"]
            Password            = $secPwd
            UsageLocation       = $userData["UsageLocation"]
            ForceChangePassword = $true
        }
        if ($userData["Department"])    { $params["Department"]  = $userData["Department"] }
        if ($userData["JobTitle"])      { $params["Title"]       = $userData["JobTitle"] }
        if ($userData["CompanyName"])   { $params["Company"]     = $userData["CompanyName"] }
        if ($userData["City"])          { $params["City"]        = $userData["City"] }
        if ($userData["State"])         { $params["State"]       = $userData["State"] }
        if ($userData["PostalCode"])    { $params["PostalCode"]  = $userData["PostalCode"] }
        if ($userData["Country"])       { $params["Country"]     = $userData["Country"] }
        if ($userData["MobilePhone"])   { $params["MobilePhone"] = $userData["MobilePhone"] }
        if ($userData["BusinessPhone"]) { $params["PhoneNumber"] = $userData["BusinessPhone"] }
        if ($userData["StreetAddress"]) { $params["StreetAddress"] = $userData["StreetAddress"] }

        New-MsolUser @params -ErrorAction Stop | Out-Null
        Write-Success "Account created: $($userData['UserPrincipalName'])"
    }
    catch {
        Write-ErrorMsg "Failed to create account: $_"
        Pause-ForUser; return
    }

    # Set extended AAD properties that New-MsolUser doesn't cover
    Start-Sleep -Seconds 3   # brief wait for AAD sync
    try {
        $newAadUser = Get-AzureADUser -ObjectId $userData["UserPrincipalName"] -ErrorAction Stop
        $aadParams = @{ ObjectId = $newAadUser.ObjectId }
        if ($userData["OfficeLocation"]) { $aadParams["PhysicalDeliveryOfficeName"] = $userData["OfficeLocation"] }
        if ($aadParams.Keys.Count -gt 1) {
            Set-AzureADUser @aadParams -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Warn "Could not set extended AAD properties: $_"
    }

    # ==================================================================
    #  STEP 2 : Assign Licenses
    # ==================================================================
    Write-SectionHeader "Step 2 - Assign Licenses"

    if ($replicateLicenses.Count -gt 0) {
        # ---- Replicate mode: auto-assign same licenses ----
        Write-InfoMsg "Replicating licenses from source user..."
        foreach ($lic in $replicateLicenses) {
            $skuName = $lic.AccountSkuId
            if (Confirm-Action "Assign license '$skuName'?") {
                try {
                    Set-MsolUserLicense -UserPrincipalName $userData["UserPrincipalName"] `
                        -AddLicenses $skuName -ErrorAction Stop
                    Write-Success "License '$skuName' assigned."
                } catch {
                    Write-ErrorMsg "Failed to assign '$skuName': $_"
                }
            }
        }

        # Offer to add more
        $addMore = Read-UserInput "Add additional licenses? (y/n)"
        if ($addMore -match '^[Yy]') {
            Select-AndAssignLicenses -UPN $userData["UserPrincipalName"]
        }
    }
    else {
        # ---- Normal mode: pick from tenant ----
        Select-AndAssignLicenses -UPN $userData["UserPrincipalName"]
    }

    # ==================================================================
    #  STEP 3 : Security Groups
    # ==================================================================
    Write-SectionHeader "Step 3 - Assign Security Groups"

    # Refresh the new AAD user object
    try { $newAadUser = Get-AzureADUser -ObjectId $userData["UserPrincipalName"] -ErrorAction Stop } catch {}

    if ($replicateGroups.Count -gt 0) {
        # ---- Replicate mode: auto-add to same groups ----
        Write-InfoMsg "Replicating security group memberships..."
        foreach ($grp in $replicateGroups) {
            if (Confirm-Action "Add to group '$($grp.DisplayName)'?") {
                try {
                    Add-AzureADGroupMember -ObjectId $grp.ObjectId `
                        -RefObjectId $newAadUser.ObjectId -ErrorAction Stop
                    Write-Success "Added to '$($grp.DisplayName)'."
                } catch {
                    if ($_.Exception.Message -match "already exist") {
                        Write-Warn "Already a member of '$($grp.DisplayName)'."
                    } else {
                        Write-ErrorMsg "Failed: $_"
                    }
                }
            }
        }

        # Offer to add more
        $addMore = Read-UserInput "Add additional security groups? (y/n)"
        if ($addMore -match '^[Yy]') {
            Search-AndAssignGroups -UserObjectId $newAadUser.ObjectId
        }
    }
    else {
        Search-AndAssignGroups -UserObjectId $newAadUser.ObjectId
    }

    # ==================================================================
    #  STEP 4 : Enforce MFA
    # ==================================================================
    Write-SectionHeader "Step 4 - Enforce MFA"

    if (Confirm-Action "Enable and enforce MFA for $($userData['UserPrincipalName'])?") {
        try {
            $mfaState = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
            $mfaState.RelyingParty = "*"
            $mfaState.State = "Enforced"
            $mfaReqs = @($mfaState)

            Set-MsolUser -UserPrincipalName $userData["UserPrincipalName"] `
                -StrongAuthenticationRequirements $mfaReqs -ErrorAction Stop
            Write-Success "MFA enforced for $($userData['UserPrincipalName'])."
        }
        catch {
            Write-ErrorMsg "MFA enforcement failed: $_"
        }
    }

    # ==================================================================
    #  STEP 5 : Set Manager
    # ==================================================================
    if ($userData["Manager"]) {
        Write-SectionHeader "Step 5 - Set Manager"
        try {
            $mgr = $null
            if ($userData["Manager"] -match '@') {
                $mgr = Get-AzureADUser -ObjectId $userData["Manager"] -ErrorAction Stop
            } else {
                $mgr = Get-AzureADUser -SearchString $userData["Manager"] -ErrorAction Stop |
                    Select-Object -First 1
            }
            if ($mgr) {
                if (Confirm-Action "Set manager to $($mgr.DisplayName) ($($mgr.UserPrincipalName))?") {
                    $newAadUser = Get-AzureADUser -ObjectId $userData["UserPrincipalName"]
                    Set-AzureADUserManager -ObjectId $newAadUser.ObjectId `
                        -RefObjectId $mgr.ObjectId -ErrorAction Stop
                    Write-Success "Manager set to $($mgr.DisplayName)."
                }
            } else {
                Write-Warn "Manager '$($userData['Manager'])' not found in AAD."
            }
        } catch {
            Write-Warn "Could not set manager: $_"
        }
    }

    # ==================================================================
    #  STEP 6 : Output password
    # ==================================================================
    Write-SectionHeader "Step 6 - Temporary Password"
    Write-Host ""
    $b = $script:Box
    Write-Host ("  " + $b.TL + [string]::new($b.H, 48) + $b.TR) -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + $b.V + "  User : $($userData['UserPrincipalName'])") -ForegroundColor White
    Write-Host ("  " + $b.V + "  Pass : $password") -ForegroundColor $script:Colors.Success
    Write-Host ("  " + $b.V + "  (User will be forced to change on first login)") -ForegroundColor $script:Colors.Info
    Write-Host ("  " + $b.BL + [string]::new($b.H, 48) + $b.BR) -ForegroundColor $script:Colors.Highlight
    Write-Host ""
    Write-Warn "Please securely deliver these credentials to the user."

    if ($replicateSource) {
        Write-Host ""
        Write-Success "Onboarding complete! (Replicated from $($replicateSource.DisplayName))"
    } else {
        Write-Success "Onboarding complete!"
    }
    Pause-ForUser
}

# ==================================================================
#  Helper: Select licenses from tenant and assign
# ==================================================================
function Select-AndAssignLicenses {
    param([string]$UPN)

    try {
        $skus = Get-MsolAccountSku | Where-Object { $_.ActiveUnits - $_.ConsumedUnits -gt 0 }
        if ($skus.Count -eq 0) {
            Write-Warn "No licenses with available seats found."
            return
        }

        $skuLabels = $skus | ForEach-Object {
            "$($_.SkuPartNumber)  (Available: $($_.ActiveUnits - $_.ConsumedUnits))"
        }
        $selected = Show-MultiSelect -Title "Select license(s) to assign" -Options $skuLabels
        foreach ($i in $selected) {
            $sku = $skus[$i]
            if (Confirm-Action "Assign license '$($sku.SkuPartNumber)' to $UPN?") {
                Set-MsolUserLicense -UserPrincipalName $UPN `
                    -AddLicenses $sku.AccountSkuId -ErrorAction Stop
                Write-Success "License '$($sku.SkuPartNumber)' assigned."
            }
        }
    }
    catch {
        Write-ErrorMsg "License assignment error: $_"
    }
}

# ==================================================================
#  Helper: Search and assign security groups
# ==================================================================
function Search-AndAssignGroups {
    param([string]$UserObjectId)

    try {
        $searchInput = Read-UserInput "Search for a security group (or 'skip')"
        if ($searchInput -eq 'skip') { return }

        $groups = Get-AzureADGroup -SearchString $searchInput -All $true |
            Where-Object { $_.SecurityEnabled -eq $true }

        if ($groups.Count -eq 0) {
            Write-Warn "No groups matching '$searchInput'."
            return
        }

        $gLabels = $groups | ForEach-Object { $_.DisplayName }
        $gSel = Show-MultiSelect -Title "Select group(s)" -Options $gLabels

        foreach ($gi in $gSel) {
            $grp = $groups[$gi]
            if (Confirm-Action "Add user to group '$($grp.DisplayName)'?") {
                try {
                    Add-AzureADGroupMember -ObjectId $grp.ObjectId `
                        -RefObjectId $UserObjectId -ErrorAction Stop
                    Write-Success "Added to '$($grp.DisplayName)'."
                } catch {
                    if ($_.Exception.Message -match "already exist") {
                        Write-Warn "Already a member of '$($grp.DisplayName)'."
                    } else {
                        Write-ErrorMsg "Failed: $_"
                    }
                }
            }
        }
    }
    catch {
        Write-ErrorMsg "Security group error: $_"
    }
}
