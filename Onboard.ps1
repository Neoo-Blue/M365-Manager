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

    # ---- Template selection (Phase 1) ----
    # Offer a role template up front. If picked, skip the file/manual/
    # replicate picker entirely -- the operator only fills user-specific
    # fields and the template provides the rest.
    $template = $null
    $templateList = @(Get-OnboardTemplates)
    if ($templateList.Count -gt 0 -and (Confirm-Action "Apply a role-based onboarding template?")) {
        $labels = $templateList | ForEach-Object { "{0} -- {1}" -f $_.Name, $_.Description }
        $tsel = Show-Menu -Title "Select Template" -Options $labels -BackLabel "Cancel (no template)"
        if ($tsel -ge 0) {
            try {
                $template = Get-OnboardTemplate -Key $templateList[$tsel].Key
                Write-Success "Template '$($template.name)' selected -- $($template.description)"
            } catch {
                Write-ErrorMsg "Could not load template: $_"; Pause-ForUser; return
            }
        }
    }

    $userData = @{}
    foreach ($f in $fields) { $userData[$f] = "" }
    $choice = -1

    if ($template) {
        Write-SectionHeader "User Details"
        $userData["FirstName"] = Read-UserInput "First name"
        $userData["LastName"]  = Read-UserInput "Last name"
        $defDisp = ("{0} {1}" -f $userData["FirstName"], $userData["LastName"]).Trim()
        $valDN   = Read-UserInput "Display name (Enter for '$defDisp')"
        $userData["DisplayName"]       = if ([string]::IsNullOrWhiteSpace($valDN)) { $defDisp } else { $valDN }
        $userData["UserPrincipalName"] = Read-UserInput "User principal name (email)"
        $userData["JobTitle"]          = Read-UserInput "Job title (optional)"
        $userData["Manager"]           = Read-UserInput "Manager UPN (optional)"

        $userData = Resolve-OnboardTemplate -Template $template -UserData $userData

        Write-Host ""
        Write-InfoMsg "Template '$($template.name)' applied. Review/edit any field before continuing:"
        $userData = Edit-UserDataTable -Data $userData -FieldOrder $fields
    }
    else {
        $choice = Show-Menu -Title "How to provide user data?" -Options @(
            "Parse from a text file",
            "Enter manually",
            "Replicate from an existing user"
        ) -BackLabel "Cancel"
        if ($choice -eq -1) { return }
    }

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

    $stubUser = [PSCustomObject]@{
        Id                = "preview-$([guid]::NewGuid())"
        UserPrincipalName = $userData['UserPrincipalName']
        DisplayName       = $userData['DisplayName']
    }
    $newUser = Invoke-Action `
        -Description ("Create user {0}" -f $userData['UserPrincipalName']) `
        -ActionType 'CreateUser' `
        -Target @{ userUpn = $userData['UserPrincipalName']; displayName = $userData['DisplayName'] } `
        -NoUndoReason 'User creation cannot be cleanly undone (deletion is destructive and removes audit history)' `
        -Critical -StubReturn $stubUser -Action {
            New-MgUser -BodyParameter $body -ErrorAction Stop
        }
    if (-not $newUser) { Write-ErrorMsg "Failed to create account."; Pause-ForUser; return }
    if (-not (Get-PreviewMode)) { Write-Success "Account created: $($userData['UserPrincipalName'])" }

    # ---- Licenses ----
    Write-SectionHeader "Step 2 - Assign Licenses"
    Start-Sleep -Seconds 5
    if ($template -and @($template.licenseSKUs).Count -gt 0) {
        Apply-TemplateLicenses -UserId $newUser.Id -Template $template
        $more = Read-UserInput "Add additional licenses? (y/n)"
        if ($more -match '^[Yy]') { Select-AndAssignLicenses -UserId $newUser.Id }
    }
    elseif ($replicateLicenses.Count -gt 0) {
        foreach ($lic in $replicateLicenses) {
            if (Confirm-Action "Assign license '$(Format-LicenseLabel $lic.SkuPartNumber)'?") {
                $ok = Invoke-Action -Description ("Assign license '{0}' to {1}" -f $lic.SkuPartNumber, $userData['UserPrincipalName']) -Action {
                    Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{SkuId = $lic.SkuId}) -RemoveLicenses @() -ErrorAction Stop; $true
                }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "'$(Get-SkuFriendlyName $lic.SkuPartNumber)' assigned." }
            }
        }
        $more = Read-UserInput "Add additional licenses? (y/n)"
        if ($more -match '^[Yy]') { Select-AndAssignLicenses -UserId $newUser.Id }
    } else { Select-AndAssignLicenses -UserId $newUser.Id }

    # ---- Security Groups ----
    Write-SectionHeader "Step 3 - Assign Security Groups"
    if ($template -and @($template.securityGroups).Count -gt 0) {
        Apply-TemplateSecurityGroups -UserId $newUser.Id -Template $template
        $more = Read-UserInput "Add additional groups? (y/n)"
        if ($more -match '^[Yy]') { Search-AndAssignGroups -UserId $newUser.Id }
    }
    elseif ($replicateGroups.Count -gt 0) {
        foreach ($grp in $replicateGroups) {
            $gName = $grp.AdditionalProperties["displayName"]
            if (Confirm-Action "Add to group '$gName'?") {
                $ok = Invoke-Action -Description ("Add {0} to security group '{1}'" -f $userData['UserPrincipalName'], $gName) -Action {
                    try {
                        New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $newUser.Id -ErrorAction Stop; $true
                    } catch {
                        if ($_.Exception.Message -match 'already exist') { 'already' } else { throw }
                    }
                }
                if (-not (Get-PreviewMode)) {
                    if ($ok -eq 'already') { Write-Warn "Already a member." }
                    elseif ($ok)           { Write-Success "Added to '$gName'." }
                }
            }
        }
        $more = Read-UserInput "Add additional groups? (y/n)"
        if ($more -match '^[Yy]') { Search-AndAssignGroups -UserId $newUser.Id }
    } else { Search-AndAssignGroups -UserId $newUser.Id }

    # ---- Distribution Lists (template only) ----
    if ($template -and @($template.distributionLists).Count -gt 0) {
        Apply-TemplateDistributionLists -Upn $userData["UserPrincipalName"] -Template $template
    }

    # ---- Shared Mailboxes (template only) ----
    if ($template -and @($template.sharedMailboxes).Count -gt 0) {
        Apply-TemplateSharedMailboxes -Upn $userData["UserPrincipalName"] -Template $template
    }

    # ---- Manager ----
    if ($userData["Manager"]) {
        Write-SectionHeader "Step 4 - Set Manager"
        try {
            $mgr = if (Get-PreviewMode) { [PSCustomObject]@{ Id = "preview-mgr"; DisplayName = $userData["Manager"]; UserPrincipalName = $userData["Manager"] } } else { Get-MgUser -UserId $userData["Manager"] -ErrorAction Stop }
            if (Confirm-Action "Set manager to $($mgr.DisplayName)?") {
                $ok = Invoke-Action -Description ("Set manager of {0} to {1}" -f $userData['UserPrincipalName'], $mgr.UserPrincipalName) -Action {
                    Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } -ErrorAction Stop; $true
                }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "Manager set." }
            }
        } catch { Write-Warn "Could not set manager: $_" }
    }

    # ---- Contractor expiry (template only) ----
    if ($template -and $template.contractorExpiryDays) {
        Apply-TemplateContractorExpiry -UserId $newUser.Id -Template $template
    }

    # ---- Password ----
    # Don't print the temp password to the host (lingers in scrollback /
    # transcript files). Push it to the clipboard, prompt the operator to
    # deliver it via a secure channel, then scrub clipboard + clear screen.
    Write-SectionHeader "Step 5 - Temporary Password"
    $b   = $script:Box
    $upn = $userData['UserPrincipalName']

    $clipped = $false
    if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
        try { Set-Clipboard -Value $password -ErrorAction Stop; $clipped = $true } catch {}
    }

    Write-Host ""
    Write-Host ("  " + $b.TL + [string]::new($b.H, 56) + $b.TR) -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + $b.V + "  User : $upn") -ForegroundColor White
    if ($clipped) {
        Write-Host ("  " + $b.V + "  Pass : <copied to clipboard>") -ForegroundColor $script:Colors.Success
        Write-Host ("  " + $b.V + "  Paste it now -- it will NOT be shown again.") -ForegroundColor $script:Colors.Warning
    } else {
        Write-Host ("  " + $b.V + "  Pass : $password") -ForegroundColor $script:Colors.Success
        Write-Host ("  " + $b.V + "  Clipboard unavailable -- copy now, screen will clear.") -ForegroundColor $script:Colors.Warning
    }
    Write-Host ("  " + $b.V + "  User must change on first sign-in.") -ForegroundColor $script:Colors.Info
    Write-Host ("  " + $b.BL + [string]::new($b.H, 56) + $b.BR) -ForegroundColor $script:Colors.Highlight
    Write-Warn "Deliver via a secure channel (password manager / encrypted message)."
    Write-Host ""
    $null = Read-UserInput "Press Enter when the password has been delivered (screen will clear)"

    # ---- Scrub ----
    # Overwrite the clipboard so the password doesn't sit on the OS pasteboard.
    # PowerShell strings are immutable so we can't truly zero $password in
    # memory, but we drop the reference so GC can reclaim it sooner.
    if ($clipped) {
        try { Set-Clipboard -Value " " -ErrorAction SilentlyContinue } catch {}
    }
    Remove-Variable password -ErrorAction SilentlyContinue

    Clear-Host
    Write-Success "Onboarding complete for $upn."
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
                $ok = Invoke-Action -Description ("Assign license '{0}' to user {1}" -f $sku.SkuPartNumber, $UserId) -Action {
                    Set-MgUserLicense -UserId $UserId -AddLicenses @(@{SkuId = $sku.SkuId}) -RemoveLicenses @() -ErrorAction Stop; $true
                }
                if ($ok -and -not (Get-PreviewMode)) { Write-Success "'$(Get-SkuFriendlyName $sku.SkuPartNumber)' assigned." }
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
                $ok = Invoke-Action -Description ("Add user {0} to security group '{1}'" -f $UserId, $grp.DisplayName) -Action {
                    try {
                        New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop; $true
                    } catch {
                        if ($_.Exception.Message -match 'already exist') { 'already' } else { throw }
                    }
                }
                if (-not (Get-PreviewMode)) {
                    if ($ok -eq 'already') { Write-Warn "Already a member." }
                    elseif ($ok)           { Write-Success "Added." }
                }
            }
        }
    } catch { Write-ErrorMsg "Group error: $_" }
}

# ============================================================
#  Apply-Template* — consume a Resolve-OnboardTemplate result to
#  apply each piece of the role bundle. Each helper continues on
#  per-item failure and logs a warning; the overall onboard does
#  not abort. Unknown SKUs / groups / DLs / mailboxes are skipped
#  with a warning per the templates/README.md trust-model.
# ============================================================

function Apply-TemplateLicenses {
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][hashtable]$Template)
    $skus = @($Template['licenseSKUs'])
    if ($skus.Count -eq 0) { return }
    Write-InfoMsg "Applying $($skus.Count) license SKU(s) from template..."

    $allSkus = @()
    if (-not (Get-PreviewMode)) {
        try { $allSkus = @(Get-MgSubscribedSku -ErrorAction Stop) }
        catch { Write-ErrorMsg "Cannot read tenant SKUs: $_"; return }
    }

    foreach ($skuPart in $skus) {
        if (-not (Get-PreviewMode)) {
            $sku = $allSkus | Where-Object { $_.SkuPartNumber -eq $skuPart } | Select-Object -First 1
            if (-not $sku) {
                Write-Warn "Skipping unknown SKU '$skuPart' (not in this tenant)."
                continue
            }
            $free = ($sku.PrepaidUnits.Enabled - $sku.ConsumedUnits)
            if ($free -le 0) {
                Write-Warn "Skipping '$skuPart' -- no seats available."
                continue
            }
        } else {
            $sku = [PSCustomObject]@{ SkuId = "preview-sku-$skuPart"; SkuPartNumber = $skuPart }
        }
        $ok = Invoke-Action `
            -Description ("Assign license '{0}' to user {1}" -f $skuPart, $UserId) `
            -ActionType 'AssignLicense' `
            -Target @{ userId = $UserId; skuId = [string]$sku.SkuId; skuPart = [string]$skuPart } `
            -ReverseType 'RemoveLicense' `
            -ReverseDescription ("Remove license '{0}' from user {1}" -f $skuPart, $UserId) `
            -Action {
                Set-MgUserLicense -UserId $UserId -AddLicenses @(@{ SkuId = $sku.SkuId }) -RemoveLicenses @() -ErrorAction Stop; $true
            }
        if ($ok -and -not (Get-PreviewMode)) { Write-Success "Assigned '$(Format-LicenseLabel $skuPart)'." }
    }
}

function Apply-TemplateSecurityGroups {
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][hashtable]$Template)
    $groups = @($Template['securityGroups'])
    if ($groups.Count -eq 0) { return }
    Write-InfoMsg "Applying $($groups.Count) security group(s) from template..."

    foreach ($name in $groups) {
        $grp = $null
        if (-not (Get-PreviewMode)) {
            try {
                $escaped = $name -replace "'", "''"
                $grp = @(Get-MgGroup -Filter "displayName eq '$escaped'" -ConsistencyLevel eventual -ErrorAction Stop) | Select-Object -First 1
                if (-not $grp) {
                    Write-Warn "Skipping unknown security group '$name'."; continue
                }
            } catch { Write-Warn "Lookup for '$name' failed: $_"; continue }
        } else {
            $grp = [PSCustomObject]@{ Id = "preview-grp-$name"; DisplayName = $name }
        }
        $ok = Invoke-Action `
            -Description ("Add user {0} to security group '{1}'" -f $UserId, $name) `
            -ActionType 'AddToGroup' `
            -Target @{ userId = $UserId; groupId = [string]$grp.Id; groupName = [string]$name } `
            -ReverseType 'RemoveFromGroup' `
            -ReverseDescription ("Remove user {0} from security group '{1}'" -f $UserId, $name) `
            -Action {
                try {
                    New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop; $true
                } catch {
                    if ($_.Exception.Message -match 'already exist') { 'already' } else { throw }
                }
            }
        if (-not (Get-PreviewMode)) {
            if ($ok -eq 'already') { Write-Warn "Already a member of '$name'." }
            elseif ($ok)           { Write-Success "Added to '$name'." }
        }
    }
}

function Apply-TemplateDistributionLists {
    param([Parameter(Mandatory)][string]$Upn, [Parameter(Mandatory)][hashtable]$Template)
    $dls = @($Template['distributionLists'])
    if ($dls.Count -eq 0) { return }
    Write-SectionHeader "Step 3a - Distribution Lists (from template)"

    foreach ($name in $dls) {
        $dlId = $name
        if (-not (Get-PreviewMode)) {
            try {
                $dl = Get-DistributionGroup -Identity $name -ErrorAction Stop
                $dlId = $dl.Identity
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match "couldn't be found|not found|Couldn't find") {
                    Write-Warn "Skipping unknown DL '$name'."; continue
                }
                Write-ErrorMsg "Lookup for DL '$name' failed: $_"; continue
            }
        }
        $ok = Invoke-Action `
            -Description ("Add {0} to distribution list '{1}'" -f $Upn, $name) `
            -ActionType 'AddToDistributionList' `
            -Target @{ upn = $Upn; dlIdentity = [string]$dlId; dlName = [string]$name } `
            -ReverseType 'RemoveFromDistributionList' `
            -ReverseDescription ("Remove {0} from distribution list '{1}'" -f $Upn, $name) `
            -Action {
                try {
                    Add-DistributionGroupMember -Identity $dlId -Member $Upn -BypassSecurityGroupManagerCheck -ErrorAction Stop; $true
                } catch {
                    if ($_.Exception.Message -match 'already a member|already exists') { 'already' } else { throw }
                }
            }
        if (-not (Get-PreviewMode)) {
            if ($ok -eq 'already') { Write-Warn "Already a member of DL '$name'." }
            elseif ($ok)           { Write-Success "Added to DL '$name'." }
        }
    }
}

function Apply-TemplateSharedMailboxes {
    param([Parameter(Mandatory)][string]$Upn, [Parameter(Mandatory)][hashtable]$Template)
    $sms = @($Template['sharedMailboxes'])
    if ($sms.Count -eq 0) { return }
    Write-SectionHeader "Step 3b - Shared Mailboxes (from template)"

    foreach ($sm in $sms) {
        $id = if ($sm -is [hashtable]) { $sm['identity'] } else { $sm.identity }
        $access = if ($sm -is [hashtable]) { $sm['access'] } else { $sm.access }
        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        if (-not (Get-PreviewMode)) {
            try { Get-Mailbox -Identity $id -ErrorAction Stop | Out-Null }
            catch { Write-Warn "Skipping unknown shared mailbox '$id'."; continue }
        }

        if ($access -eq 'Full' -or $access -eq 'FullSendAs') {
            $ok = Invoke-Action `
                -Description ("Grant {0} FullAccess on shared mailbox '{1}'" -f $Upn, $id) `
                -ActionType 'GrantMailboxFullAccess' `
                -Target @{ mailbox = [string]$id; user = $Upn } `
                -ReverseType 'RevokeMailboxFullAccess' `
                -ReverseDescription ("Revoke {0} FullAccess on shared mailbox '{1}'" -f $Upn, $id) `
                -Action {
                    try {
                        Add-MailboxPermission -Identity $id -User $Upn -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop; $true
                    } catch {
                        if ($_.Exception.Message -match 'already exist|already a member') { 'already' } else { throw }
                    }
                }
            if (-not (Get-PreviewMode)) {
                if ($ok -eq 'already') { Write-Warn "Already has FullAccess on '$id'." }
                elseif ($ok)           { Write-Success "FullAccess on '$id' granted." }
            }
        }
        if ($access -eq 'SendAs' -or $access -eq 'FullSendAs') {
            $ok = Invoke-Action `
                -Description ("Grant {0} SendAs on shared mailbox '{1}'" -f $Upn, $id) `
                -ActionType 'GrantMailboxSendAs' `
                -Target @{ mailbox = [string]$id; user = $Upn } `
                -ReverseType 'RevokeMailboxSendAs' `
                -ReverseDescription ("Revoke {0} SendAs on shared mailbox '{1}'" -f $Upn, $id) `
                -Action {
                    try {
                        Add-RecipientPermission -Identity $id -Trustee $Upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop; $true
                    } catch {
                        if ($_.Exception.Message -match 'already exist|already a member') { 'already' } else { throw }
                    }
                }
            if (-not (Get-PreviewMode)) {
                if ($ok -eq 'already') { Write-Warn "Already has SendAs on '$id'." }
                elseif ($ok)           { Write-Success "SendAs on '$id' granted." }
            }
        }
    }
}

function Apply-TemplateContractorExpiry {
    <#
        Records a planned end date in employeeLeaveDateTime. This alone
        does NOT disable the account on that date -- needs Entra
        lifecycle workflows or a scheduled job. We log the intent.
    #>
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][hashtable]$Template)
    $days = [int]$Template['contractorExpiryDays']
    if ($days -le 0) { return }
    $when = (Get-Date).ToUniversalTime().AddDays($days)
    $iso  = $when.ToString("yyyy-MM-ddTHH:mm:ssZ")
    Write-SectionHeader "Step 4a - Contractor Expiry (from template)"
    $ok = Invoke-Action -Description ("Set employeeLeaveDateTime for user {0} to {1} (+{2} days)" -f $UserId, $iso, $days) -Action {
        Update-MgUser -UserId $UserId -BodyParameter @{ employeeLeaveDateTime = $iso } -ErrorAction Stop; $true
    }
    if ($ok -and -not (Get-PreviewMode)) {
        Write-Success "employeeLeaveDateTime set to $iso (+$days days)."
        Write-Warn "Auto-disable requires Entra lifecycle workflows or a scheduled task -- this only records the intent."
    }
}
