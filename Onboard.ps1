# ============================================================
#  Onboard.ps1 - New User Onboarding (Microsoft Graph)
# ============================================================

function Start-Onboard {
    Write-SectionHeader "User Onboarding"

    $fields = @(
        "FirstName","LastName","DisplayName","UserPrincipalName",
        "JobTitle","Department","CompanyName","OfficeLocation",
        "StreetAddress","City","State","PostalCode","Country",
        "UsageLocation","BusinessPhone","MobilePhone","Manager"
    )

    $replicateSource   = $null
    $replicateLicenses = @()
    $replicateGroups   = @()

    $choice = Show-Menu -Title "How to provide user data?" -Options @(
        "Parse from a text file",
        "Enter manually",
        "Replicate from an existing user"
    ) -BackLabel "Cancel"
    if ($choice -eq -1) { return }

    $userData = @{}

    if ($choice -eq 0) {
        $filePath = Read-UserInput "Enter full path to the text file"
        if (-not (Test-Path $filePath)) { Write-ErrorMsg "File not found: $filePath"; Pause-ForUser; return }
        Write-InfoMsg "Parsing file..."
        $lines = Get-Content $filePath
        foreach ($line in $lines) {
            if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.+)$') {
                $key = $Matches[1].Trim(); $value = $Matches[2].Trim()
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
                    'manager'                   { $userData["Manager"]           = $value }
                    default                     { $userData[$key]                = $value }
                }
            }
        }
        foreach ($f in $fields) { if (-not $userData.ContainsKey($f)) { $userData[$f] = "" } }
        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }
        Write-Success "Parsed data:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    elseif ($choice -eq 1) {
        foreach ($f in $fields) { $userData[$f] = Read-UserInput "Enter $f" }
        if ([string]::IsNullOrWhiteSpace($userData["DisplayName"]) -and $userData["FirstName"] -and $userData["LastName"]) {
            $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
        }
        Write-InfoMsg "Review the information:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    elseif ($choice -eq 2) {
        if (-not (Connect-ForTask "Onboard")) { Write-ErrorMsg "Could not connect."; Pause-ForUser; return }

        Write-SectionHeader "Select User to Replicate From"
        $replicateSource = Resolve-UserIdentity -PromptText "Enter the existing user's name or email"
        if ($null -eq $replicateSource) { Write-ErrorMsg "No source user."; Pause-ForUser; return }

        $src = $replicateSource
        $userData["FirstName"] = ""; $userData["LastName"] = ""; $userData["DisplayName"] = ""; $userData["UserPrincipalName"] = ""
        $userData["JobTitle"]       = if ($src.JobTitle)       { $src.JobTitle } else { "" }
        $userData["Department"]     = if ($src.Department)     { $src.Department } else { "" }
        $userData["CompanyName"]    = if ($src.CompanyName)    { $src.CompanyName } else { "" }
        $userData["OfficeLocation"] = if ($src.OfficeLocation) { $src.OfficeLocation } else { "" }
        $userData["StreetAddress"]  = if ($src.StreetAddress)  { $src.StreetAddress } else { "" }
        $userData["City"]           = if ($src.City)           { $src.City } else { "" }
        $userData["State"]          = if ($src.State)          { $src.State } else { "" }
        $userData["PostalCode"]     = if ($src.PostalCode)     { $src.PostalCode } else { "" }
        $userData["Country"]        = if ($src.Country)        { $src.Country } else { "" }
        $userData["UsageLocation"]  = if ($src.UsageLocation)  { $src.UsageLocation } else { "" }
        $userData["BusinessPhone"]  = if ($src.BusinessPhones.Count -gt 0) { $src.BusinessPhones[0] } else { "" }
        $userData["MobilePhone"]    = if ($src.MobilePhone)    { $src.MobilePhone } else { "" }

        try {
            $srcMgr = Get-MgUserManager -UserId $src.Id -ErrorAction SilentlyContinue
            $userData["Manager"] = if ($srcMgr) { $srcMgr.AdditionalProperties["userPrincipalName"] } else { "" }
        } catch { $userData["Manager"] = "" }

        Write-InfoMsg "Reading licenses..."
        try { $replicateLicenses = @(Get-MgUserLicenseDetail -UserId $src.Id -ErrorAction Stop) } catch { $replicateLicenses = @() }

        Write-InfoMsg "Reading security groups..."
        try {
            $allMemberships = @(Get-MgUserMemberOf -UserId $src.Id -All -ErrorAction Stop)
            $replicateGroups = @($allMemberships | Where-Object {
                $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" -and
                $_.AdditionalProperties["securityEnabled"] -eq $true
            })
        } catch { $replicateGroups = @() }

        Write-SectionHeader "Replicated from $($src.DisplayName)"
        if ($replicateLicenses.Count -gt 0) {
            Write-InfoMsg "Licenses:"; $replicateLicenses | ForEach-Object { Write-Host "    - $(Format-LicenseLabel $_.SkuPartNumber)" -ForegroundColor White }
        }
        if ($replicateGroups.Count -gt 0) {
            Write-InfoMsg "Security groups:"; $replicateGroups | ForEach-Object { Write-Host "    - $($_.AdditionalProperties['displayName'])" -ForegroundColor White }
        }
        Write-Host ""
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }

    # Validate
    foreach ($r in @("FirstName","LastName","UserPrincipalName","UsageLocation")) {
        if ([string]::IsNullOrWhiteSpace($userData[$r])) { Write-ErrorMsg "'$r' is required."; Pause-ForUser; return }
    }
    if ([string]::IsNullOrWhiteSpace($userData["DisplayName"])) {
        $userData["DisplayName"] = "$($userData['FirstName']) $($userData['LastName'])"
    }

    if (-not (Connect-ForTask "Onboard")) { Pause-ForUser; return }

    # ---- Create Account ----
    Write-SectionHeader "Step 1 - Create Account"
    $password = -join ((48..57) + (65..90) + (97..122) + (33,35,36,37,38) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
    $mailNick = ($userData["UserPrincipalName"] -split '@')[0]

    $details = "UPN: $($userData['UserPrincipalName'])`nName: $($userData['DisplayName'])`nDept: $($userData['Department'])"
    if (-not (Confirm-Action "Create this user account?" $details)) { Pause-ForUser; return }

    try {
        $body = @{
            AccountEnabled    = $true
            DisplayName       = $userData["DisplayName"]
            GivenName         = $userData["FirstName"]
            Surname           = $userData["LastName"]
            UserPrincipalName = $userData["UserPrincipalName"]
            MailNickname      = $mailNick
            UsageLocation     = $userData["UsageLocation"]
            PasswordProfile   = @{ ForceChangePasswordNextSignIn = $true; Password = $password }
        }
        if ($userData["JobTitle"])       { $body["JobTitle"]       = $userData["JobTitle"] }
        if ($userData["Department"])     { $body["Department"]     = $userData["Department"] }
        if ($userData["CompanyName"])    { $body["CompanyName"]    = $userData["CompanyName"] }
        if ($userData["OfficeLocation"]) { $body["OfficeLocation"] = $userData["OfficeLocation"] }
        if ($userData["StreetAddress"])  { $body["StreetAddress"]  = $userData["StreetAddress"] }
        if ($userData["City"])           { $body["City"]           = $userData["City"] }
        if ($userData["State"])          { $body["State"]          = $userData["State"] }
        if ($userData["PostalCode"])     { $body["PostalCode"]     = $userData["PostalCode"] }
        if ($userData["Country"])        { $body["Country"]        = $userData["Country"] }
        if ($userData["MobilePhone"])    { $body["MobilePhone"]    = $userData["MobilePhone"] }
        if ($userData["BusinessPhone"])  { $body["BusinessPhones"] = @($userData["BusinessPhone"]) }

        $newUser = New-MgUser -BodyParameter $body -ErrorAction Stop
        Write-Success "Account created: $($userData['UserPrincipalName'])"
    } catch { Write-ErrorMsg "Failed to create account: $_"; Pause-ForUser; return }

    # ---- Licenses ----
    Write-SectionHeader "Step 2 - Assign Licenses"
    Start-Sleep -Seconds 5
    if ($replicateLicenses.Count -gt 0) {
        foreach ($lic in $replicateLicenses) {
            if (Confirm-Action "Assign license '$(Format-LicenseLabel $lic.SkuPartNumber)'?") {
                try { Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{SkuId = $lic.SkuId}) -RemoveLicenses @() -ErrorAction Stop; Write-Success "'$(Get-SkuFriendlyName $lic.SkuPartNumber)' assigned." }
                catch { Write-ErrorMsg "Failed: $_" }
            }
        }
        $more = Read-UserInput "Add additional licenses? (y/n)"
        if ($more -match '^[Yy]') { Select-AndAssignLicenses -UserId $newUser.Id }
    } else { Select-AndAssignLicenses -UserId $newUser.Id }

    # ---- Security Groups ----
    Write-SectionHeader "Step 3 - Assign Security Groups"
    if ($replicateGroups.Count -gt 0) {
        foreach ($grp in $replicateGroups) {
            $gName = $grp.AdditionalProperties["displayName"]
            if (Confirm-Action "Add to group '$gName'?") {
                try { New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $newUser.Id -ErrorAction Stop; Write-Success "Added to '$gName'." }
                catch { if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." } else { Write-ErrorMsg "Failed: $_" } }
            }
        }
        $more = Read-UserInput "Add additional groups? (y/n)"
        if ($more -match '^[Yy]') { Search-AndAssignGroups -UserId $newUser.Id }
    } else { Search-AndAssignGroups -UserId $newUser.Id }

    # ---- Manager ----
    if ($userData["Manager"]) {
        Write-SectionHeader "Step 4 - Set Manager"
        try {
            $mgr = Get-MgUser -UserId $userData["Manager"] -ErrorAction Stop
            if (Confirm-Action "Set manager to $($mgr.DisplayName)?") {
                Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } -ErrorAction Stop
                Write-Success "Manager set."
            }
        } catch { Write-Warn "Could not set manager: $_" }
    }

    # ---- Password ----
    Write-SectionHeader "Step 5 - Temporary Password"
    $b = $script:Box
    Write-Host ""; Write-Host ("  " + $b.TL + [string]::new($b.H, 48) + $b.TR) -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + $b.V + "  User : $($userData['UserPrincipalName'])") -ForegroundColor White
    Write-Host ("  " + $b.V + "  Pass : $password") -ForegroundColor $script:Colors.Success
    Write-Host ("  " + $b.V + "  (User must change on first login)") -ForegroundColor $script:Colors.Info
    Write-Host ("  " + $b.BL + [string]::new($b.H, 48) + $b.BR) -ForegroundColor $script:Colors.Highlight
    Write-Warn "Securely deliver these credentials to the user."
    Write-Success "Onboarding complete!"
    Pause-ForUser
}

function Select-AndAssignLicenses {
    param([string]$UserId)
    try {
        $skus = Get-MgSubscribedSku -ErrorAction Stop
        $available = $skus | ForEach-Object {
            $total = $_.PrepaidUnits.Enabled; $used = $_.ConsumedUnits
            [PSCustomObject]@{ SkuPartNumber = $_.SkuPartNumber; SkuId = $_.SkuId; FriendlyName = Format-LicenseLabel $_.SkuPartNumber; Total = $total; Used = $used; Free = $total - $used }
        }
        $labels = $available | ForEach-Object { "$($_.FriendlyName)  [Total: $($_.Total) | Used: $($_.Used) | Free: $($_.Free)]" }
        $selected = Show-MultiSelect -Title "Select license(s)" -Options $labels
        foreach ($idx in $selected) {
            $sku = $available[$idx]
            if ($sku.Free -le 0) { Write-Warn "$(Get-SkuFriendlyName $sku.SkuPartNumber) has no seats."; continue }
            if (Confirm-Action "Assign '$(Get-SkuFriendlyName $sku.SkuPartNumber)'?") {
                Set-MgUserLicense -UserId $UserId -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @() -ErrorAction Stop
                Write-Success "'$(Get-SkuFriendlyName $sku.SkuPartNumber)' assigned."
            }
        }
    } catch { Write-ErrorMsg "License error: $_" }
}

function Search-AndAssignGroups {
    param([string]$UserId)
    $searchInput = Read-UserInput "Search for a security group (or 'skip')"
    if ($searchInput -eq 'skip') { return }
    try {
        $groups = @(Get-MgGroup -Search "displayName:$searchInput" -ConsistencyLevel eventual -ErrorAction Stop | Where-Object { $_.SecurityEnabled })
        if ($groups.Count -eq 0) { Write-Warn "No groups found."; return }
        $gLabels = $groups | ForEach-Object { $_.DisplayName }
        $gSel = Show-MultiSelect -Title "Select group(s)" -Options $gLabels
        foreach ($gi in $gSel) {
            $grp = $groups[$gi]
            if (Confirm-Action "Add to '$($grp.DisplayName)'?") {
                try { New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop; Write-Success "Added." }
                catch { if ($_.Exception.Message -match "already exist") { Write-Warn "Already a member." } else { Write-ErrorMsg "Failed: $_" } }
            }
        }
    } catch { Write-ErrorMsg "Group error: $_" }
}
