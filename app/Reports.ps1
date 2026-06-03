# ============================================================
#  Reports.ps1 - Interactive Reporting & CSV Export
#  Features: scope picker, active/inactive filter, TUI sorting
# ============================================================

function Start-ReportingMenu {
    Write-SectionHeader "Reporting"

    if (-not (Connect-ForTask "Report")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $keepGoing = $true
    while ($keepGoing) {
        $sel = Show-Menu -Title "Select Report" -Options @(
            "License Report",
            "Mailbox Size Report",
            "Archive Mailbox Report",
            "User Account Status Report",
            "Group Membership Report",
            "Shared Mailbox Report",
            "Inactive Users Report"
        ) -BackLabel "Back to Main Menu"

        switch ($sel) {
            0 { Invoke-LicenseReport }
            1 { Invoke-MailboxSizeReport }
            2 { Invoke-ArchiveReport }
            3 { Invoke-UserStatusReport }
            4 { Invoke-GroupMembershipReport }
            5 { Invoke-SharedMailboxReport }
            6 { Invoke-InactiveUsersReport }
            -1 { $keepGoing = $false }
        }
    }
}

# ==================================================================
#  Shared: Scope picker, active/inactive filter, sorting, export
# ==================================================================

function Select-ReportScope {
    $scope = Show-Menu -Title "Report Scope" -Options @(
        "Entire organization",
        "Single user",
        "Members of a group"
    ) -BackLabel "Cancel"
    switch ($scope) {
        0 { return "Org" }
        1 { return "User" }
        2 { return "Group" }
        -1 { return $null }
    }
}

function Select-ActiveFilter {
    <# Returns: "All", "ActiveOnly", "InactiveOnly" #>
    $f = Show-Menu -Title "Include which users?" -Options @(
        "All users",
        "Active users only (sign-in enabled)",
        "Inactive/disabled users only"
    ) -BackLabel "Cancel"
    switch ($f) {
        0 { return "All" }
        1 { return "ActiveOnly" }
        2 { return "InactiveOnly" }
        -1 { return $null }
    }
}

function Apply-ActiveFilter {
    param(
        [array]$Users,
        [string]$Filter
    )
    if ($Filter -eq "ActiveOnly") {
        $filtered = @($Users | Where-Object { $_.AccountEnabled -eq $true })
        Write-InfoMsg "Filtered to $($filtered.Count) active users (from $($Users.Count))."
        return $filtered
    }
    elseif ($Filter -eq "InactiveOnly") {
        $filtered = @($Users | Where-Object { $_.AccountEnabled -eq $false })
        Write-InfoMsg "Filtered to $($filtered.Count) disabled users (from $($Users.Count))."
        return $filtered
    }
    return $Users
}

function Get-ScopeUsers {
    param([string]$Scope, [string]$ActiveFilter = "All")

    $userProps = "Id,DisplayName,UserPrincipalName,AccountEnabled,JobTitle,Department,UsageLocation,CreatedDateTime,SignInActivity"

    switch ($Scope) {
        "Org" {
            Write-InfoMsg "Fetching all users (this may take a moment)..."
            $users = @(Get-MgUser -All -Property $userProps -ErrorAction Stop)
            return Apply-ActiveFilter -Users $users -Filter $ActiveFilter
        }
        "User" {
            $user = Resolve-UserIdentity
            if ($null -eq $user) { return $null }
            $full = Get-MgUser -UserId $user.Id -Property $userProps -ErrorAction Stop
            return @($full)
        }
        "Group" {
            $searchInput = Read-UserInput "Search for group by name"
            if ([string]::IsNullOrWhiteSpace($searchInput)) { return $null }
            $groups = @(Get-MgGroup -Search "displayName:$searchInput" -ConsistencyLevel eventual -ErrorAction Stop)
            if ($groups.Count -eq 0) { Write-ErrorMsg "No groups found."; return $null }
            $grp = if ($groups.Count -eq 1) { $groups[0] } else {
                $sel = Show-Menu -Title "Select Group" -Options ($groups | ForEach-Object { $_.DisplayName }) -BackLabel "Cancel"
                if ($sel -eq -1) { return $null }; $groups[$sel]
            }
            Write-InfoMsg "Fetching members of '$($grp.DisplayName)'..."
            $members = @(Get-MgGroupMember -GroupId $grp.Id -All -ErrorAction Stop)
            $users = @()
            foreach ($m in $members) {
                if ($m.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user") {
                    try { $users += Get-MgUser -UserId $m.Id -Property $userProps -ErrorAction Stop } catch {}
                }
            }
            Write-InfoMsg "Found $($users.Count) user members."
            return Apply-ActiveFilter -Users $users -Filter $ActiveFilter
        }
    }
    return $null
}

function Select-SortColumn {
    <# Asks user which column to sort by and direction. Returns hashtable with Column and Descending. #>
    param([string[]]$Columns)

    Write-Host ""
    Write-InfoMsg "Sort options:"
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        Write-Host "    [" -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host ($i + 1) -NoNewline -ForegroundColor $script:Colors.Highlight
        Write-Host "] " -NoNewline -ForegroundColor $script:Colors.Accent
        Write-Host $Columns[$i] -ForegroundColor $script:Colors.Menu
    }
    Write-Host ""
    $colInput = Read-UserInput "Sort by column # (or Enter for default)"

    $sortCol = $null
    if ($colInput -match '^\d+$') {
        $idx = [int]$colInput
        if ($idx -ge 1 -and $idx -le $Columns.Count) {
            $sortCol = $Columns[$idx - 1]
        }
    }

    if ($null -eq $sortCol) {
        return @{ Column = $null; Descending = $false }
    }

    $dir = Show-Menu -Title "Sort direction for '$sortCol'" -Options @(
        "Ascending (A-Z, 0-9, oldest first)",
        "Descending (Z-A, 9-0, newest first)"
    ) -BackLabel "Default (ascending)"

    $desc = ($dir -eq 1)

    return @{ Column = $sortCol; Descending = $desc }
}

function Sort-ReportData {
    param(
        [array]$Data,
        [string]$Column,
        [bool]$Descending = $false
    )

    if ($null -eq $Column -or $Data.Count -eq 0) { return $Data }

    # Try numeric sort first, fall back to string
    $isNumeric = $true
    foreach ($row in ($Data | Select-Object -First 10)) {
        $val = "$($row.$Column)"
        if ($val -and $val -ne "N/A" -and $val -ne "Never" -and $val -notmatch '^\d+$') {
            $isNumeric = $false; break
        }
    }

    if ($isNumeric) {
        if ($Descending) {
            return @($Data | Sort-Object { $v = "$($_.$Column)"; if ($v -match '^\d+$') { [int64]$v } else { -1 } } -Descending)
        } else {
            return @($Data | Sort-Object { $v = "$($_.$Column)"; if ($v -match '^\d+$') { [int64]$v } else { [int64]::MaxValue } })
        }
    } else {
        if ($Descending) {
            return @($Data | Sort-Object $Column -Descending)
        } else {
            return @($Data | Sort-Object $Column)
        }
    }
}

function Show-ReportTable {
    param(
        [string]$Title,
        [array]$Data,
        [string[]]$Columns,
        [hashtable]$ColumnWidths = @{},
        [switch]$SkipSort
    )

    if ($Data.Count -eq 0) { Write-InfoMsg "No data to display."; return $Data }

    # ---- Interactive sort ----
    if (-not $SkipSort -and $Data.Count -gt 1) {
        $sortChoice = Select-SortColumn -Columns $Columns
        if ($sortChoice.Column) {
            $Data = Sort-ReportData -Data $Data -Column $sortChoice.Column -Descending $sortChoice.Descending
            Write-Success "Sorted by '$($sortChoice.Column)' $(if ($sortChoice.Descending) { '(descending)' } else { '(ascending)' })"
        }
    }

    Write-SectionHeader "$Title ($($Data.Count) records)"

    # Build header
    $headerLine = "  "
    foreach ($col in $Columns) {
        $w = if ($ColumnWidths.ContainsKey($col)) { $ColumnWidths[$col] } else { 20 }
        $headerLine += ("{0,-$w}" -f $col)
    }
    Write-Host $headerLine -ForegroundColor $script:Colors.Highlight
    Write-Host ("  " + ("-" * ($headerLine.Length - 2))) -ForegroundColor $script:Colors.Accent

    # Display rows (cap at 50)
    $displayCount = [math]::Min($Data.Count, 50)
    for ($i = 0; $i -lt $displayCount; $i++) {
        $row = $Data[$i]
        $line = "  "
        foreach ($col in $Columns) {
            $w = if ($ColumnWidths.ContainsKey($col)) { $ColumnWidths[$col] } else { 20 }
            $val = "$($row.$col)"
            if ($val.Length -gt ($w - 1)) { $val = $val.Substring(0, $w - 3) + ".." }
            $line += ("{0,-$w}" -f $val)
        }
        Write-Host $line -ForegroundColor White
    }

    if ($Data.Count -gt 50) {
        Write-Host ""
        Write-Warn "Showing first 50 of $($Data.Count) records. Export to CSV for full data."
    }

    return $Data
}

function Export-ReportCsv {
    param([string]$ReportName, [array]$Data)

    if ($Data.Count -eq 0) { Write-Warn "No data to export."; return }

    $exportChoice = Show-Menu -Title "Export to CSV?" -Options @(
        "Yes, export to C:\Temp",
        "Yes, choose a folder",
        "No, skip export"
    ) -BackLabel "Skip"

    if ($exportChoice -eq 2 -or $exportChoice -eq -1) { return }

    $folder = "C:\Temp"
    if ($exportChoice -eq 1) {
        $folder = Read-UserInput "Enter folder path"
        if ([string]::IsNullOrWhiteSpace($folder)) { $folder = "C:\Temp" }
    }

    if (-not (Test-Path $folder)) {
        try { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
        catch { Write-ErrorMsg "Could not create folder: $_"; return }
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $fileName = "${ReportName}_${timestamp}.csv"
    $filePath = Join-Path $folder $fileName

    try {
        $Data | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Success "Report exported: $filePath"
        Write-StatusLine "Records" "$($Data.Count)" "White"
        Write-StatusLine "Size" "$([math]::Round((Get-Item $filePath).Length / 1KB, 1)) KB" "White"
    } catch { Write-ErrorMsg "Export failed: $_" }
}

# ==================================================================
#  1. License Report
# ==================================================================
function Invoke-LicenseReport {
    Write-SectionHeader "License Report"

    $type = Show-Menu -Title "License Report Type" -Options @(
        "Tenant license summary (all SKUs)",
        "Per-user license assignments"
    ) -BackLabel "Cancel"
    if ($type -eq -1) { return }

    if ($type -eq 0) {
        Write-InfoMsg "Fetching tenant license inventory..."
        try {
            $skus = Get-MgSubscribedSku -ErrorAction Stop
            $data = $skus | ForEach-Object {
                $t = $_.PrepaidUnits.Enabled; $u = $_.ConsumedUnits
                [PSCustomObject]@{
                    License   = Get-SkuFriendlyName $_.SkuPartNumber
                    SKU       = $_.SkuPartNumber
                    Total     = $t
                    Assigned  = $u
                    Available = $t - $u
                    UsagePct  = if ($t -gt 0) { [math]::Round(($u / $t) * 100) } else { 0 }
                }
            }

            $data = Show-ReportTable -Title "Tenant License Summary" -Data @($data) `
                -Columns @("License","SKU","Total","Assigned","Available","UsagePct") `
                -ColumnWidths @{ License = 35; SKU = 25; Total = 8; Assigned = 10; Available = 10; UsagePct = 8 }

            Export-ReportCsv -ReportName "LicenseSummary" -Data $data
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    else {
        $scope = Select-ReportScope
        if ($null -eq $scope) { Pause-ForUser; return }

        $activeFilter = "All"
        if ($scope -ne "User") {
            $activeFilter = Select-ActiveFilter
            if ($null -eq $activeFilter) { Pause-ForUser; return }
        }

        $users = Get-ScopeUsers -Scope $scope -ActiveFilter $activeFilter
        if ($null -eq $users -or $users.Count -eq 0) { Write-Warn "No users found."; Pause-ForUser; return }

        Write-InfoMsg "Fetching license details for $($users.Count) user(s)..."
        $data = @()
        foreach ($u in $users) {
            try {
                $lics = @(Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction SilentlyContinue)
                if ($lics.Count -eq 0) {
                    $data += [PSCustomObject]@{
                        DisplayName = $u.DisplayName; UPN = $u.UserPrincipalName
                        Enabled = $u.AccountEnabled; Department = $u.Department
                        License = "(none)"; SKU = ""
                    }
                } else {
                    foreach ($lic in $lics) {
                        $data += [PSCustomObject]@{
                            DisplayName = $u.DisplayName; UPN = $u.UserPrincipalName
                            Enabled = $u.AccountEnabled; Department = $u.Department
                            License = Get-SkuFriendlyName $lic.SkuPartNumber; SKU = $lic.SkuPartNumber
                        }
                    }
                }
            } catch {}
        }

        $data = Show-ReportTable -Title "User License Assignments" -Data @($data) `
            -Columns @("DisplayName","UPN","Enabled","Department","License") `
            -ColumnWidths @{ DisplayName = 20; UPN = 28; Enabled = 9; Department = 14; License = 28 }

        Export-ReportCsv -ReportName "UserLicenses" -Data $data
    }
    Pause-ForUser
}

# ==================================================================
#  2. Mailbox Size Report
# ==================================================================
function Invoke-MailboxSizeReport {
    Write-SectionHeader "Mailbox Size Report"

    $scope = Select-ReportScope
    if ($null -eq $scope) { Pause-ForUser; return }

    $activeFilter = "All"
    if ($scope -ne "User") {
        $activeFilter = Select-ActiveFilter
        if ($null -eq $activeFilter) { Pause-ForUser; return }
    }

    $users = Get-ScopeUsers -Scope $scope -ActiveFilter $activeFilter
    if ($null -eq $users -or $users.Count -eq 0) { Write-Warn "No users."; Pause-ForUser; return }

    Write-InfoMsg "Fetching mailbox statistics for $($users.Count) user(s)..."
    $data = @()
    $errors = 0
    $i = 0
    foreach ($u in $users) {
        $i++; if ($i % 25 -eq 0) { Write-InfoMsg "  Processing $i of $($users.Count)..." }

        $upn = $u.UserPrincipalName
        $stats = $null

        # First verify the mailbox exists
        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
        } catch {
            $errors++
            if ($users.Count -le 5) { Write-Warn "No mailbox for $upn : $_" }
            continue
        }

        # Try Get-EXOMailboxStatistics first (modern), fallback to Get-MailboxStatistics
        try {
            $stats = Get-EXOMailboxStatistics -Identity $upn -ErrorAction Stop
        } catch {
            try {
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop
            } catch {
                $errors++
                if ($users.Count -le 5) { Write-Warn "Stats failed for $upn : $_" }
                continue
            }
        }

        if ($stats) {
            $sizeStr = "N/A"
            $sizeBytes = 0
            if ($stats.TotalItemSize) {
                $sizeStr = $stats.TotalItemSize.ToString()
                try { $sizeBytes = $stats.TotalItemSize.Value.ToBytes() } catch {
                    # Try parsing from string if .Value is not available
                    if ($sizeStr -match '([\d,]+)\s*bytes') {
                        $sizeBytes = [int64]($Matches[1] -replace ',','')
                    }
                }
            }
            $lastAct = "N/A"
            if ($stats.LastInteractionTime) {
                $lastAct = $stats.LastInteractionTime.ToString("yyyy-MM-dd")
            } elseif ($stats.LastLogonTime) {
                $lastAct = $stats.LastLogonTime.ToString("yyyy-MM-dd")
            }

            $data += [PSCustomObject]@{
                DisplayName  = $u.DisplayName
                UPN          = $upn
                Enabled      = $u.AccountEnabled
                Department   = $u.Department
                MailboxType  = $mbx.RecipientTypeDetails
                ItemCount    = $stats.ItemCount
                TotalSize    = $sizeStr
                SizeBytes    = $sizeBytes
                LastActivity = $lastAct
            }
        }
    }

    if ($errors -gt 0) {
        Write-Warn "$errors user(s) had no mailbox or stats could not be retrieved."
    }

    $data = Show-ReportTable -Title "Mailbox Sizes" -Data @($data) `
        -Columns @("DisplayName","UPN","MailboxType","ItemCount","TotalSize","LastActivity") `
        -ColumnWidths @{ DisplayName = 18; UPN = 26; MailboxType = 14; ItemCount = 10; TotalSize = 22; LastActivity = 12 }

    $csvData = $data | Select-Object DisplayName, UPN, Enabled, Department, MailboxType, ItemCount, TotalSize, LastActivity
    Export-ReportCsv -ReportName "MailboxSizes" -Data $csvData
    Pause-ForUser
}

# ==================================================================
#  3. Archive Mailbox Report
# ==================================================================
function Invoke-ArchiveReport {
    Write-SectionHeader "Archive Mailbox Report"

    $scope = Select-ReportScope
    if ($null -eq $scope) { Pause-ForUser; return }

    $activeFilter = "All"
    if ($scope -ne "User") {
        $activeFilter = Select-ActiveFilter
        if ($null -eq $activeFilter) { Pause-ForUser; return }
    }

    $users = Get-ScopeUsers -Scope $scope -ActiveFilter $activeFilter
    if ($null -eq $users -or $users.Count -eq 0) { Write-Warn "No users."; Pause-ForUser; return }

    Write-InfoMsg "Fetching archive info for $($users.Count) user(s)..."
    $data = @()
    $errors = 0
    $i = 0
    foreach ($u in $users) {
        $i++; if ($i % 25 -eq 0) { Write-InfoMsg "  Processing $i of $($users.Count)..." }
        $upn = $u.UserPrincipalName

        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
        } catch {
            $errors++
            if ($users.Count -le 5) { Write-Warn "No mailbox for $upn : $_" }
            continue
        }

        $archEnabled = $mbx.ArchiveStatus -eq "Active"
        $archSize = "N/A"; $archItems = "N/A"

        if ($archEnabled) {
            try {
                $archStats = $null
                try { $archStats = Get-EXOMailboxStatistics -Identity $upn -Archive -ErrorAction Stop }
                catch { $archStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction Stop }

                if ($archStats) {
                    $archSize = $archStats.TotalItemSize.ToString()
                    $archItems = $archStats.ItemCount
                }
            } catch {
                if ($users.Count -le 5) { Write-Warn "Archive stats failed for $upn : $_" }
            }
        }

        $data += [PSCustomObject]@{
            DisplayName = $u.DisplayName
            UPN         = $upn
            Enabled     = $u.AccountEnabled
            Archive     = $archEnabled
            Policy      = if ($mbx.RetentionPolicy) { $mbx.RetentionPolicy } else { "(default)" }
            ArchSize    = $archSize
            ArchItems   = $archItems
        }
    }

    if ($errors -gt 0) {
        Write-Warn "$errors user(s) had no mailbox."
    }

    $data = Show-ReportTable -Title "Archive Mailbox Status" -Data @($data) `
        -Columns @("DisplayName","UPN","Enabled","Archive","Policy","ArchSize","ArchItems") `
        -ColumnWidths @{ DisplayName = 18; UPN = 26; Enabled = 9; Archive = 9; Policy = 20; ArchSize = 16; ArchItems = 10 }

    Export-ReportCsv -ReportName "ArchiveReport" -Data $data
    Pause-ForUser
}

# ==================================================================
#  4. User Account Status Report
# ==================================================================
function Invoke-UserStatusReport {
    Write-SectionHeader "User Account Status Report"

    $scope = Select-ReportScope
    if ($null -eq $scope) { Pause-ForUser; return }

    $activeFilter = "All"
    if ($scope -ne "User") {
        $activeFilter = Select-ActiveFilter
        if ($null -eq $activeFilter) { Pause-ForUser; return }
    }

    $users = Get-ScopeUsers -Scope $scope -ActiveFilter $activeFilter
    if ($null -eq $users -or $users.Count -eq 0) { Write-Warn "No users."; Pause-ForUser; return }

    Write-InfoMsg "Building status report for $($users.Count) user(s)..."
    $data = @()
    $i = 0
    foreach ($u in $users) {
        $i++; if ($i % 50 -eq 0) { Write-InfoMsg "  Processing $i of $($users.Count)..." }
        $lastSignIn = "N/A"
        if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
            $lastSignIn = $u.SignInActivity.LastSignInDateTime.ToString("yyyy-MM-dd HH:mm")
        }
        $created = if ($u.CreatedDateTime) { $u.CreatedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
        $licCount = 0
        try { $licCount = @(Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction SilentlyContinue).Count } catch {}

        $data += [PSCustomObject]@{
            DisplayName = $u.DisplayName; UPN = $u.UserPrincipalName
            Enabled = $u.AccountEnabled; Department = $u.Department
            JobTitle = $u.JobTitle; Licenses = $licCount
            Created = $created; LastSignIn = $lastSignIn
        }
    }

    $data = Show-ReportTable -Title "User Account Status" -Data @($data) `
        -Columns @("DisplayName","UPN","Enabled","Department","Licenses","LastSignIn") `
        -ColumnWidths @{ DisplayName = 20; UPN = 26; Enabled = 9; Department = 14; Licenses = 9; LastSignIn = 18 }

    Export-ReportCsv -ReportName "UserStatus" -Data $data
    Pause-ForUser
}

# ==================================================================
#  5. Group Membership Report
# ==================================================================
function Invoke-GroupMembershipReport {
    Write-SectionHeader "Group Membership Report"

    $type = Show-Menu -Title "Report Type" -Options @(
        "All groups a user belongs to",
        "All members of a specific group",
        "All security groups with member counts"
    ) -BackLabel "Cancel"
    if ($type -eq -1) { return }

    $data = @()

    if ($type -eq 0) {
        $user = Resolve-UserIdentity
        if ($null -eq $user) { Pause-ForUser; return }
        Write-InfoMsg "Fetching group memberships..."
        try {
            $memberships = @(Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction Stop)
            foreach ($m in $memberships) {
                if ($m.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                    $data += [PSCustomObject]@{
                        User = $user.DisplayName; UPN = $user.UserPrincipalName
                        GroupName = $m.AdditionalProperties["displayName"]
                        GroupType = if ($m.AdditionalProperties["securityEnabled"]) { "Security" } else { "M365/DL" }
                        Mail = $m.AdditionalProperties["mail"]
                    }
                }
            }
        } catch { Write-ErrorMsg "Failed: $_" }

        $data = Show-ReportTable -Title "Group Memberships for $($user.DisplayName)" -Data @($data) `
            -Columns @("GroupName","GroupType","Mail") `
            -ColumnWidths @{ GroupName = 35; GroupType = 12; Mail = 35 }
    }
    elseif ($type -eq 1) {
        $searchInput = Read-UserInput "Search group by name"
        if ([string]::IsNullOrWhiteSpace($searchInput)) { Pause-ForUser; return }
        try {
            $groups = @(Get-MgGroup -Search "displayName:$searchInput" -ConsistencyLevel eventual -ErrorAction Stop)
            if ($groups.Count -eq 0) { Write-ErrorMsg "No groups found."; Pause-ForUser; return }
            $grp = if ($groups.Count -eq 1) { $groups[0] } else {
                $sel = Show-Menu -Title "Select" -Options ($groups | ForEach-Object { $_.DisplayName }) -BackLabel "Cancel"
                if ($sel -eq -1) { Pause-ForUser; return }; $groups[$sel]
            }

            # Active/inactive filter for group members
            $activeFilter = Select-ActiveFilter

            Write-InfoMsg "Fetching members of '$($grp.DisplayName)'..."
            $members = @(Get-MgGroupMember -GroupId $grp.Id -All -ErrorAction Stop)
            foreach ($m in $members) {
                $mType = ($m.AdditionalProperties["@odata.type"] -replace '#microsoft.graph.','')
                $mEnabled = $m.AdditionalProperties["accountEnabled"]

                # Apply filter for user members
                if ($mType -eq "user" -and $null -ne $activeFilter -and $activeFilter -ne "All") {
                    if ($activeFilter -eq "ActiveOnly" -and $mEnabled -ne $true) { continue }
                    if ($activeFilter -eq "InactiveOnly" -and $mEnabled -ne $false) { continue }
                }

                $data += [PSCustomObject]@{
                    GroupName = $grp.DisplayName
                    MemberName = $m.AdditionalProperties["displayName"]
                    MemberUPN = $m.AdditionalProperties["userPrincipalName"]
                    MemberType = $mType
                    Enabled = if ($null -ne $mEnabled) { $mEnabled } else { "N/A" }
                }
            }
        } catch { Write-ErrorMsg "Failed: $_" }

        $data = Show-ReportTable -Title "Members of '$($grp.DisplayName)'" -Data @($data) `
            -Columns @("MemberName","MemberUPN","MemberType","Enabled") `
            -ColumnWidths @{ MemberName = 25; MemberUPN = 32; MemberType = 12; Enabled = 9 }
    }
    elseif ($type -eq 2) {
        Write-InfoMsg "Fetching all security groups..."
        try {
            $allGroups = @(Get-MgGroup -All -Filter "securityEnabled eq true" -Property "Id,DisplayName,Mail,MailEnabled,MembershipRule" -ErrorAction Stop)
            Write-InfoMsg "Found $($allGroups.Count) security groups. Counting members..."
            $i = 0
            foreach ($grp in $allGroups) {
                $i++; if ($i % 25 -eq 0) { Write-InfoMsg "  Processing $i of $($allGroups.Count)..." }
                $count = 0
                try { $count = @(Get-MgGroupMember -GroupId $grp.Id -All -ErrorAction SilentlyContinue).Count } catch {}
                $data += [PSCustomObject]@{
                    GroupName = $grp.DisplayName; Mail = $grp.Mail
                    MailEnabled = $grp.MailEnabled
                    Type = if ($grp.MembershipRule) { "Dynamic" } else { "Static" }
                    Members = $count
                }
            }
        } catch { Write-ErrorMsg "Failed: $_" }

        $data = Show-ReportTable -Title "Security Groups" -Data @($data) `
            -Columns @("GroupName","Mail","Type","Members") `
            -ColumnWidths @{ GroupName = 35; Mail = 30; Type = 10; Members = 8 }
    }

    Export-ReportCsv -ReportName "GroupReport" -Data $data
    Pause-ForUser
}

# ==================================================================
#  6. Shared Mailbox Report
# ==================================================================
function Invoke-SharedMailboxReport {
    Write-SectionHeader "Shared Mailbox Report"

    Write-InfoMsg "Fetching all shared mailboxes..."
    try {
        $boxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
        Write-InfoMsg "Found $($boxes.Count) shared mailboxes. Gathering stats..."

        $data = @()
        $i = 0
        foreach ($box in $boxes) {
            $i++; if ($i % 10 -eq 0) { Write-InfoMsg "  Processing $i of $($boxes.Count)..." }
            $size = "N/A"; $items = 0
            try {
                $stats = $null
                try { $stats = Get-EXOMailboxStatistics -Identity $box.PrimarySmtpAddress -ErrorAction Stop }
                catch { $stats = Get-MailboxStatistics -Identity $box.PrimarySmtpAddress -ErrorAction Stop }
                if ($stats) { $size = $stats.TotalItemSize.ToString(); $items = $stats.ItemCount }
            } catch {}

            $accessCount = 0
            try {
                $perms = @(Get-MailboxPermission -Identity $box.PrimarySmtpAddress -ErrorAction SilentlyContinue |
                    Where-Object { $_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-*" -and -not $_.IsInherited })
                $accessCount = $perms.Count
            } catch {}

            $fwd = if ($box.ForwardingSmtpAddress) { $box.ForwardingSmtpAddress -replace 'smtp:','' } else { "" }

            $data += [PSCustomObject]@{
                DisplayName = $box.DisplayName; Email = $box.PrimarySmtpAddress
                ItemCount = $items; TotalSize = $size
                UsersAccess = $accessCount; Forwarding = $fwd
                Hidden = $box.HiddenFromAddressListsEnabled
            }
        }

        $data = Show-ReportTable -Title "Shared Mailboxes" -Data @($data) `
            -Columns @("DisplayName","Email","ItemCount","TotalSize","UsersAccess","Hidden") `
            -ColumnWidths @{ DisplayName = 22; Email = 26; ItemCount = 10; TotalSize = 18; UsersAccess = 8; Hidden = 8 }

        Export-ReportCsv -ReportName "SharedMailboxes" -Data $data
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

# ==================================================================
#  7. Inactive Users Report
# ==================================================================
function Invoke-InactiveUsersReport {
    Write-SectionHeader "Inactive Users Report"

    $daysInput = Read-UserInput "Show users inactive for more than how many days? (default: 90)"
    $days = 90
    if ($daysInput -match '^\d+$') { $days = [int]$daysInput }
    $cutoff = (Get-Date).AddDays(-$days)

    $accountFilter = Show-Menu -Title "Which accounts to check?" -Options @(
        "Enabled accounts only (licensed users still consuming seats)",
        "Disabled accounts only (check if cleanup needed)",
        "All accounts"
    ) -BackLabel "Cancel"
    if ($accountFilter -eq -1) { Pause-ForUser; return }

    Write-InfoMsg "Fetching users..."
    try {
        $filterStr = switch ($accountFilter) {
            0 { "accountEnabled eq true" }
            1 { "accountEnabled eq false" }
            2 { $null }
        }

        $fetchParams = @{
            All      = $true
            Property = "Id,DisplayName,UserPrincipalName,Department,JobTitle,AccountEnabled,SignInActivity,CreatedDateTime"
        }
        if ($filterStr) { $fetchParams["Filter"] = $filterStr }

        $allUsers = @(Get-MgUser @fetchParams -ErrorAction Stop)
        Write-InfoMsg "Checking sign-in activity for $($allUsers.Count) users..."

        $data = @()
        $i = 0
        foreach ($u in $allUsers) {
            $i++; if ($i % 100 -eq 0) { Write-InfoMsg "  Processing $i of $($allUsers.Count)..." }
            $lastSignIn = $null
            if ($u.SignInActivity -and $u.SignInActivity.LastSignInDateTime) {
                $lastSignIn = $u.SignInActivity.LastSignInDateTime
            }

            $isInactive = ($null -eq $lastSignIn) -or ($lastSignIn -lt $cutoff)

            if ($isInactive) {
                $daysSince = if ($lastSignIn) { [math]::Round(((Get-Date) - $lastSignIn).TotalDays) } else { 9999 }
                $created = if ($u.CreatedDateTime) { $u.CreatedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
                $licCount = 0
                try { $licCount = @(Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction SilentlyContinue).Count } catch {}

                $data += [PSCustomObject]@{
                    DisplayName  = $u.DisplayName; UPN = $u.UserPrincipalName
                    Enabled = $u.AccountEnabled; Department = $u.Department
                    Licenses = $licCount; Created = $created
                    LastSignIn = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd") } else { "Never" }
                    DaysInactive = if ($daysSince -eq 9999) { "Never" } else { $daysSince }
                }
            }
        }

        Write-Host ""
        Write-StatusLine "Inactive threshold" "$days days" "Yellow"
        Write-StatusLine "Total users checked" "$($allUsers.Count)" "White"
        Write-StatusLine "Inactive users" "$($data.Count)" "Red"
        if ($allUsers.Count -gt 0) {
            Write-StatusLine "Inactive rate" ("{0:P1}" -f ($data.Count / $allUsers.Count)) "Yellow"
        }
        Write-Host ""

        $data = Show-ReportTable -Title "Inactive Users (${days}+ days)" -Data @($data) `
            -Columns @("DisplayName","UPN","Enabled","Department","Licenses","LastSignIn","DaysInactive") `
            -ColumnWidths @{ DisplayName = 18; UPN = 26; Enabled = 9; Department = 12; Licenses = 9; LastSignIn = 12; DaysInactive = 8 }

        Export-ReportCsv -ReportName "InactiveUsers" -Data $data
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}
