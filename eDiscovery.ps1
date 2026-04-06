# ============================================================
#  eDiscovery.ps1 - eDiscovery Management
#  Simple mode: quick content searches
#  Advanced mode: full case, hold, search, export management
# ============================================================

function Start-eDiscoveryMenu {
    Write-SectionHeader "eDiscovery"

    # If SCC is already connected without search session, reconnect properly
    if ($script:SessionState.ComplianceCenter) {
        Write-InfoMsg "Verifying SCC search session..."
        # Quick test - if it fails with the search session error, reconnect
        try {
            Get-ComplianceSearch -Identity "___test_nonexistent___" -ErrorAction Stop 2>$null
        } catch {
            $testErr = $_.Exception.Message
            if ($testErr -match "EnableSearchOnlySession|search initialization|Connect-IPPSSession") {
                Write-Warn "Current SCC session lacks search capability. Reconnecting..."
                try {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    $script:SessionState.ComplianceCenter = $false
                    $script:SessionState.ExchangeOnline = $false
                } catch {}
            }
        }
    }

    if (-not (Connect-ForTask "eDiscovery")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    $keepGoing = $true
    while ($keepGoing) {
        $mode = Show-Menu -Title "eDiscovery Mode" -Options @(
            "Simple - Quick Content Search",
            "Advanced - Case Management"
        ) -BackLabel "Back to Main Menu"

        switch ($mode) {
            0 { Start-SimpleSearch }
            1 { Start-AdvancedDiscovery }
            -1 { $keepGoing = $false }
        }
    }
}

# ============================================================
#  SIMPLE MODE - Quick Content Search (no case required)
# ============================================================

function Start-SimpleSearch {
    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Simple Content Search" -Options @(
            "New quick search",
            "View existing searches",
            "View search results / statistics",
            "Delete a search"
        ) -BackLabel "Back"

        switch ($action) {
            0 { New-QuickSearch }
            1 { Show-AllSearches }
            2 { View-SearchResults }
            3 { Remove-SearchFlow }
            -1 { $keepGoing = $false }
        }
    }
}

function New-QuickSearch {
    Write-SectionHeader "New Quick Content Search"

    $name = Read-UserInput "Search name (descriptive label)"
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    # ---- Build search query ----
    Write-SectionHeader "Search Criteria"

    $keywords = Read-UserInput "Keywords (e.g. 'confidential project alpha') or press Enter for all"

    $dateFrom = Read-UserInput "Date from (yyyy-MM-dd) or Enter to skip"
    $dateTo   = Read-UserInput "Date to (yyyy-MM-dd) or Enter to skip"

    $sender    = Read-UserInput "From sender (email) or Enter to skip"
    $recipient = Read-UserInput "To recipient (email) or Enter to skip"

    $msgType = Show-Menu -Title "Message type filter" -Options @(
        "All types",
        "Email only",
        "Documents only",
        "Instant messages (Teams/Skype)",
        "Meetings"
    ) -BackLabel "All types"

    # ---- Build KQL query ----
    $kqlParts = @()
    if ($keywords) { $kqlParts += $keywords }
    if ($dateFrom -match '^\d{4}-\d{2}-\d{2}$') { $kqlParts += "sent>=$dateFrom" }
    if ($dateTo -match '^\d{4}-\d{2}-\d{2}$')   { $kqlParts += "sent<=$dateTo" }
    if ($sender)    { $kqlParts += "from:$sender" }
    if ($recipient) { $kqlParts += "to:$recipient" }

    $typeMap = @{ 1 = "email"; 2 = "documents"; 3 = "im"; 4 = "meetings" }
    if ($msgType -ge 1 -and $msgType -le 4) {
        $kqlParts += "kind:$($typeMap[$msgType])"
    }

    $kqlQuery = if ($kqlParts.Count -gt 0) { $kqlParts -join " AND " } else { "*" }

    Write-Host ""
    Write-StatusLine "KQL Query" $kqlQuery "Cyan"
    Write-Host ""

    # ---- Scope: which mailboxes ----
    $scopeChoice = Show-Menu -Title "Search scope" -Options @(
        "All mailboxes in the organization",
        "Specific mailbox(es)"
    ) -BackLabel "Cancel"
    if ($scopeChoice -eq -1) { return }

    $locations = $null
    if ($scopeChoice -eq 1) {
        $mailboxes = @()
        $adding = $true
        while ($adding) {
            $mbxInput = Read-UserInput "Enter mailbox email (or 'done')"
            if ($mbxInput -match '^done$') { break }
            if ($mbxInput) { $mailboxes += $mbxInput }
        }
        if ($mailboxes.Count -eq 0) { Write-Warn "No mailboxes specified."; return }
        $locations = $mailboxes -join ","
    }

    # ---- Confirm and create ----
    $details = "Name: $name`nQuery: $kqlQuery`nScope: $(if ($locations) { $locations } else { 'All mailboxes' })"
    if (-not (Confirm-Action "Create and start this content search?" $details)) { return }

    try {
        $searchParams = @{
            Name            = $name
            ContentMatchQuery = $kqlQuery
            ExchangeLocation = if ($locations) { $mailboxes } else { "All" }
        }

        New-ComplianceSearch @searchParams -ErrorAction Stop | Out-Null
        Write-Success "Search '$name' created."

        Write-InfoMsg "Starting search..."
        Invoke-ComplianceSearchStart -SearchName $name | Out-Null
        Write-InfoMsg "Use 'View search results' to check status and see results."
    } catch {
        Write-ErrorMsg "Failed to create search: $_"
    }
    Pause-ForUser
}

function Show-AllSearches {
    Write-SectionHeader "All Content Searches"
    try {
        $searches = @(Get-ComplianceSearch -ErrorAction Stop)
        if ($searches.Count -eq 0) {
            Write-InfoMsg "No content searches found."
            Pause-ForUser; return
        }

        foreach ($s in $searches) {
            $b = $script:Box
            Write-Host ""
            Write-Host ("  " + $b.TL + [string]::new($b.H, 56) + $b.TR) -ForegroundColor $script:Colors.Accent
            Write-StatusLine "  Name" $s.Name "Cyan"
            Write-StatusLine "  Status" $s.Status $(if ($s.Status -eq "Completed") { "Green" } else { "Yellow" })
            Write-StatusLine "  Items" "$($s.Items)" "White"
            Write-StatusLine "  Size" "$($s.Size)" "White"
            Write-StatusLine "  Created" "$($s.CreatedTime)" "Gray"
            $queryDisplay = if ($s.ContentMatchQuery) { $s.ContentMatchQuery } else { "(all content)" }
            if ($queryDisplay.Length -gt 60) { $queryDisplay = $queryDisplay.Substring(0, 57) + "..." }
            Write-StatusLine "  Query" $queryDisplay "White"
            Write-Host ("  " + $b.BL + [string]::new($b.H, 56) + $b.BR) -ForegroundColor $script:Colors.Accent
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function View-SearchResults {
    Write-SectionHeader "Search Results"

    $search = Find-ComplianceSearchByName
    if ($null -eq $search) { Pause-ForUser; return }

    try {
        $s = Get-ComplianceSearch -Identity $search.Name -ErrorAction Stop

        Write-StatusLine "Name" $s.Name "Cyan"
        Write-StatusLine "Status" $s.Status $(if ($s.Status -eq "Completed") { "Green" } else { "Yellow" })
        Write-StatusLine "Items Found" "$($s.Items)" "White"
        Write-StatusLine "Total Size" "$($s.Size)" "White"
        Write-StatusLine "Query" $(if ($s.ContentMatchQuery) { $s.ContentMatchQuery } else { "(all)" }) "White"
        Write-StatusLine "Created" "$($s.CreatedTime)" "Gray"
        Write-StatusLine "Completed" "$($s.CompletedTime)" "Gray"
        Write-Host ""

        if ($s.Status -ne "Completed") {
            Write-Warn "Search is still $($s.Status). Results may be incomplete."
        }

        # Show per-location breakdown if available
        if ($s.SuccessResults) {
            Write-SectionHeader "Results by Location"
            $resultLines = $s.SuccessResults -split "`n"
            $locData = @()
            foreach ($line in $resultLines) {
                if ($line -match 'Location:\s*(.+?),\s*Item count:\s*(\d+),\s*Total size:\s*(\d+)') {
                    $locData += [PSCustomObject]@{
                        Location  = $Matches[1].Trim()
                        ItemCount = [int]$Matches[2]
                        SizeBytes = [int64]$Matches[3]
                        SizeMB    = [math]::Round([int64]$Matches[3] / 1MB, 2)
                    }
                }
            }
            if ($locData.Count -gt 0) {
                $locData = @($locData | Sort-Object ItemCount -Descending)
                $displayCount = [math]::Min($locData.Count, 30)
                for ($i = 0; $i -lt $displayCount; $i++) {
                    $loc = $locData[$i]
                    Write-Host "    $($loc.Location): $($loc.ItemCount) items ($($loc.SizeMB) MB)" -ForegroundColor White
                }
                if ($locData.Count -gt 30) { Write-Warn "Showing top 30 of $($locData.Count) locations." }
            }
        }

        Write-Host ""

        # Actions
        $action = Show-Menu -Title "Actions" -Options @(
            "Preview results (create preview action)",
            "Export results (create export action)",
            "Re-run this search",
            "View export status"
        ) -BackLabel "Done"

        switch ($action) {
            0 {
                if (Confirm-Action "Create a preview action for '$($s.Name)'?") {
                    try {
                        New-ComplianceSearchAction -SearchName $s.Name -Preview -ErrorAction Stop | Out-Null
                        Write-Success "Preview action created. Check status with 'View export status'."
                    } catch { Write-ErrorMsg "Preview failed: $_" }
                }
            }
            1 {
                $exportFormat = Show-Menu -Title "Export format" -Options @(
                    "PST files",
                    "Individual messages",
                    "Both"
                ) -BackLabel "Cancel"

                if ($exportFormat -ne -1) {
                    $formatMap = @{ 0 = "SoftDelete"; 1 = "SingleMsg"; 2 = "HardDelete" }
                    if (Confirm-Action "Create export action for '$($s.Name)'?") {
                        try {
                            New-ComplianceSearchAction -SearchName $s.Name -Export -Format $formatMap[$exportFormat] -ErrorAction Stop | Out-Null
                            Write-Success "Export action created."
                            Write-InfoMsg "Download the export from the Compliance portal:"
                            Write-InfoMsg "  https://compliance.microsoft.com > Content Search > Export tab"
                        } catch { Write-ErrorMsg "Export failed: $_" }
                    }
                }
            }
            2 {
                if (Confirm-Action "Re-run search '$($s.Name)'?") {
                    Invoke-ComplianceSearchStart -SearchName $s.Name | Out-Null
                }
            }
            3 { Show-ExportStatus }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function Remove-SearchFlow {
    Write-SectionHeader "Delete Content Search"
    $search = Find-ComplianceSearchByName
    if ($null -eq $search) { Pause-ForUser; return }

    Write-Warn "This will delete the search and all its results."
    if (Confirm-Action "DELETE search '$($search.Name)'?") {
        $check = Read-UserInput "Type the search name to confirm"
        if ($check -eq $search.Name) {
            try {
                # Remove any actions first
                $actions = @(Get-ComplianceSearchAction -ErrorAction SilentlyContinue | Where-Object { $_.SearchName -eq $search.Name })
                foreach ($a in $actions) {
                    Remove-ComplianceSearchAction -Identity $a.Name -Confirm:$false -ErrorAction SilentlyContinue
                }
                Remove-ComplianceSearch -Identity $search.Name -Confirm:$false -ErrorAction Stop
                Write-Success "Search '$($search.Name)' deleted."
            } catch { Write-ErrorMsg "Failed: $_" }
        } else { Write-Warn "Name mismatch. Cancelled." }
    }
    Pause-ForUser
}

# ============================================================
#  ADVANCED MODE - Full Case Management
# ============================================================

function Start-AdvancedDiscovery {
    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Advanced eDiscovery" -Options @(
            "Case Management (create, view, close, delete)",
            "Search Management (searches within a case)",
            "Hold Management (legal holds)",
            "Export Management (export status)"
        ) -BackLabel "Back"

        switch ($action) {
            0 { Start-CaseManagement }
            1 { Start-CaseSearchManagement }
            2 { Start-HoldManagement }
            3 { Show-ExportStatus }
            -1 { $keepGoing = $false }
        }
    }
}

# ---- Case Management ----

function Start-CaseManagement {
    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Case Management" -Options @(
            "Create new case",
            "List all cases",
            "View case details",
            "Add member to case",
            "Close a case",
            "Reopen a case",
            "Delete a case"
        ) -BackLabel "Back"

        switch ($action) {
            0 { New-DiscoveryCase }
            1 { Show-AllCases }
            2 { View-CaseDetails }
            3 { Add-CaseMember }
            4 { Close-DiscoveryCase }
            5 { Reopen-DiscoveryCase }
            6 { Remove-DiscoveryCase }
            -1 { $keepGoing = $false }
        }
    }
}

function New-DiscoveryCase {
    Write-SectionHeader "Create New eDiscovery Case"

    $name = Read-UserInput "Case name"
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    $desc = Read-UserInput "Description (or Enter to skip)"

    $caseType = Show-Menu -Title "Case type" -Options @(
        "eDiscovery Standard",
        "eDiscovery Premium (Advanced)"
    ) -BackLabel "Cancel"
    if ($caseType -eq -1) { return }

    $details = "Name: $name`nType: $(if ($caseType -eq 0) { 'Standard' } else { 'Premium' })`nDescription: $(if ($desc) { $desc } else { '(none)' })"
    if (-not (Confirm-Action "Create this eDiscovery case?" $details)) { return }

    try {
        $params = @{ Name = $name }
        if ($desc) { $params["Description"] = $desc }
        if ($caseType -eq 1) { $params["CaseType"] = "AdvancedEdiscovery" }

        New-ComplianceCase @params -ErrorAction Stop | Out-Null
        Write-Success "Case '$name' created."
    } catch { Write-ErrorMsg "Failed to create case: $_" }
    Pause-ForUser
}

function Show-AllCases {
    Write-SectionHeader "All eDiscovery Cases"
    try {
        $cases = @(Get-ComplianceCase -ErrorAction Stop)
        if ($cases.Count -eq 0) { Write-InfoMsg "No cases found."; Pause-ForUser; return }

        foreach ($c in $cases) {
            $statusColor = switch ($c.Status) {
                "Active"   { "Green" }
                "Closed"   { "Red" }
                default    { "Yellow" }
            }
            Write-Host ""
            Write-StatusLine "Name" $c.Name "Cyan"
            Write-StatusLine "Status" $c.Status $statusColor
            Write-StatusLine "Type" $c.CaseType "White"
            Write-StatusLine "Created" "$($c.CreatedDateTime)" "Gray"
            if ($c.Description) { Write-StatusLine "Desc" $c.Description "Gray" }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function View-CaseDetails {
    $case = Find-Case
    if ($null -eq $case) { Pause-ForUser; return }

    Write-SectionHeader "Case: $($case.Name)"
    Write-StatusLine "Status" $case.Status $(if ($case.Status -eq "Active") { "Green" } else { "Red" })
    Write-StatusLine "Type" $case.CaseType "White"
    Write-StatusLine "Created" "$($case.CreatedDateTime)" "Gray"
    if ($case.Description) { Write-StatusLine "Description" $case.Description "White" }

    # Searches in this case
    Write-Host ""
    Write-InfoMsg "Searches in this case:"
    try {
        $searches = @(Get-ComplianceSearch -Case $case.Name -ErrorAction Stop)
        if ($searches.Count -eq 0) { Write-InfoMsg "  (none)" }
        else { $searches | ForEach-Object { Write-Host "    - $($_.Name) [$($_.Status)] ($($_.Items) items)" -ForegroundColor White } }
    } catch { Write-Warn "Could not list searches: $_" }

    # Holds in this case
    Write-Host ""
    Write-InfoMsg "Holds in this case:"
    try {
        $holds = @(Get-CaseHoldPolicy -Case $case.Name -ErrorAction Stop)
        if ($holds.Count -eq 0) { Write-InfoMsg "  (none)" }
        else { $holds | ForEach-Object { Write-Host "    - $($_.Name) [Enabled: $($_.Enabled)]" -ForegroundColor White } }
    } catch { Write-Warn "Could not list holds: $_" }

    Pause-ForUser
}

function Add-CaseMember {
    $case = Find-Case
    if ($null -eq $case) { Pause-ForUser; return }

    $memberEmail = Read-UserInput "Enter member email to add to case '$($case.Name)'"
    if ([string]::IsNullOrWhiteSpace($memberEmail)) { return }

    if (Confirm-Action "Add '$memberEmail' as member of case '$($case.Name)'?") {
        try {
            Add-ComplianceCaseMember -Case $case.Name -Member $memberEmail -ErrorAction Stop
            Write-Success "Member added."
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Close-DiscoveryCase {
    $case = Find-Case -StatusFilter "Active"
    if ($null -eq $case) { Pause-ForUser; return }

    Write-Warn "Closing a case releases all holds and makes it read-only."
    if (Confirm-Action "Close case '$($case.Name)'?") {
        try {
            Set-ComplianceCase -Identity $case.Name -Status "Closed" -ErrorAction Stop
            Write-Success "Case '$($case.Name)' closed."
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Reopen-DiscoveryCase {
    $case = Find-Case -StatusFilter "Closed"
    if ($null -eq $case) { Pause-ForUser; return }

    if (Confirm-Action "Reopen case '$($case.Name)'?") {
        try {
            Set-ComplianceCase -Identity $case.Name -Status "Active" -ErrorAction Stop
            Write-Success "Case '$($case.Name)' reopened."
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

function Remove-DiscoveryCase {
    $case = Find-Case -StatusFilter "Closed"
    if ($null -eq $case) {
        Write-Warn "Only closed cases can be deleted. Close the case first."
        Pause-ForUser; return
    }

    Write-Warn "This permanently deletes the case and ALL its searches, holds, and exports!"
    if (Confirm-Action "DELETE case '$($case.Name)'?") {
        $check = Read-UserInput "Type the case name to confirm"
        if ($check -eq $case.Name) {
            try {
                Remove-ComplianceCase -Identity $case.Name -Confirm:$false -ErrorAction Stop
                Write-Success "Case deleted."
            } catch { Write-ErrorMsg "Failed: $_" }
        } else { Write-Warn "Mismatch. Cancelled." }
    }
    Pause-ForUser
}

# ---- Search Management (within a case) ----

function Start-CaseSearchManagement {
    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Search Management" -Options @(
            "Create search in a case",
            "List searches in a case",
            "Run / re-run a search",
            "View search results",
            "Delete a search from a case"
        ) -BackLabel "Back"

        switch ($action) {
            0 { New-CaseSearch }
            1 { List-CaseSearches }
            2 { Run-CaseSearch }
            3 { View-SearchResults }
            4 { Remove-CaseSearchFlow }
            -1 { $keepGoing = $false }
        }
    }
}

function New-CaseSearch {
    Write-SectionHeader "Create Search in Case"

    $case = Find-Case -StatusFilter "Active"
    if ($null -eq $case) { Pause-ForUser; return }

    $name = Read-UserInput "Search name"
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    Write-SectionHeader "Build Search Query"

    $queryMode = Show-Menu -Title "Query mode" -Options @(
        "Guided (answer prompts to build query)",
        "Raw KQL (type your own KQL query)"
    ) -BackLabel "Cancel"
    if ($queryMode -eq -1) { return }

    $kqlQuery = ""

    if ($queryMode -eq 0) {
        # Guided
        $keywords   = Read-UserInput "Keywords (or Enter for all)"
        $dateFrom   = Read-UserInput "Date from (yyyy-MM-dd) or Enter to skip"
        $dateTo     = Read-UserInput "Date to (yyyy-MM-dd) or Enter to skip"
        $sender     = Read-UserInput "From sender (email) or Enter to skip"
        $recipient  = Read-UserInput "To recipient (email) or Enter to skip"
        $subject    = Read-UserInput "Subject contains (or Enter to skip)"
        $hasAttach  = Read-UserInput "Has attachments? (y/n/Enter to skip)"
        $fileExt    = Read-UserInput "Attachment file type (e.g. docx, pdf) or Enter to skip"
        $msgType = Show-Menu -Title "Message type" -Options @("All","Email","Documents","IM (Teams)","Meetings") -BackLabel "All"

        $parts = @()
        if ($keywords)    { $parts += "($keywords)" }
        if ($dateFrom -match '^\d{4}-\d{2}-\d{2}$') { $parts += "sent>=$dateFrom" }
        if ($dateTo -match '^\d{4}-\d{2}-\d{2}$')   { $parts += "sent<=$dateTo" }
        if ($sender)      { $parts += "from:$sender" }
        if ($recipient)   { $parts += "to:$recipient" }
        if ($subject)     { $parts += "subject:`"$subject`"" }
        if ($hasAttach -match '^[Yy]') { $parts += "hasattachment:true" }
        if ($fileExt)     { $parts += "filetype:$fileExt" }
        $typeMap = @{ 1 = "email"; 2 = "documents"; 3 = "im"; 4 = "meetings" }
        if ($msgType -ge 1 -and $msgType -le 4) { $parts += "kind:$($typeMap[$msgType])" }

        $kqlQuery = if ($parts.Count -gt 0) { $parts -join " AND " } else { "*" }
    }
    else {
        # Raw KQL
        Write-InfoMsg "Common KQL operators:"
        Write-Host "    from:user@domain.com       - Sender" -ForegroundColor Gray
        Write-Host "    to:user@domain.com         - Recipient" -ForegroundColor Gray
        Write-Host "    subject:`"quarterly report`" - Subject" -ForegroundColor Gray
        Write-Host "    sent>=2024-01-01           - Date from" -ForegroundColor Gray
        Write-Host "    kind:email / im / meetings - Type" -ForegroundColor Gray
        Write-Host "    hasattachment:true          - Has attachments" -ForegroundColor Gray
        Write-Host "    filetype:pdf               - File type" -ForegroundColor Gray
        Write-Host "    participants:user@domain    - Any participant" -ForegroundColor Gray
        Write-Host "    cc:user@domain.com         - CC recipient" -ForegroundColor Gray
        Write-Host "    bcc:user@domain.com        - BCC recipient" -ForegroundColor Gray
        Write-Host "    size>1000000               - Size > 1MB" -ForegroundColor Gray
        Write-Host ""
        $kqlQuery = Read-UserInput "Enter KQL query"
        if ([string]::IsNullOrWhiteSpace($kqlQuery)) { $kqlQuery = "*" }
    }

    Write-Host ""
    Write-StatusLine "Query" $kqlQuery "Cyan"
    Write-Host ""

    # Scope
    $scopeChoice = Show-Menu -Title "Search scope" -Options @(
        "All mailboxes",
        "Specific mailbox(es)",
        "All SharePoint sites",
        "All mailboxes + all SharePoint"
    ) -BackLabel "Cancel"
    if ($scopeChoice -eq -1) { return }

    $exLoc = $null; $spLoc = $null
    switch ($scopeChoice) {
        0 { $exLoc = "All" }
        1 {
            $mailboxes = @()
            while ($true) {
                $mb = Read-UserInput "Mailbox email (or 'done')"
                if ($mb -match '^done$') { break }
                if ($mb) { $mailboxes += $mb }
            }
            if ($mailboxes.Count -eq 0) { Write-Warn "No mailboxes."; return }
            $exLoc = $mailboxes
        }
        2 { $spLoc = "All" }
        3 { $exLoc = "All"; $spLoc = "All" }
    }

    $details = "Case: $($case.Name)`nSearch: $name`nQuery: $kqlQuery`nExchange: $(if ($exLoc) { $exLoc -join ', ' } else { 'none' })`nSharePoint: $(if ($spLoc) { $spLoc } else { 'none' })"
    if (-not (Confirm-Action "Create this search?" $details)) { return }

    try {
        $params = @{
            Name              = $name
            Case              = $case.Name
            ContentMatchQuery = $kqlQuery
        }
        if ($exLoc) { $params["ExchangeLocation"] = $exLoc }
        if ($spLoc) { $params["SharePointLocation"] = $spLoc }

        New-ComplianceSearch @params -ErrorAction Stop | Out-Null
        Write-Success "Search '$name' created in case '$($case.Name)'."

        $runNow = Read-UserInput "Start search now? (y/n)"
        if ($runNow -match '^[Yy]') {
            Invoke-ComplianceSearchStart -SearchName $name | Out-Null
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function List-CaseSearches {
    $case = Find-Case
    if ($null -eq $case) { Pause-ForUser; return }

    Write-SectionHeader "Searches in '$($case.Name)'"
    try {
        $searches = @(Get-ComplianceSearch -Case $case.Name -ErrorAction Stop)
        if ($searches.Count -eq 0) { Write-InfoMsg "No searches in this case." }
        else {
            foreach ($s in $searches) {
                Write-Host ""
                Write-StatusLine "Name" $s.Name "Cyan"
                Write-StatusLine "Status" $s.Status $(if ($s.Status -eq "Completed") { "Green" } else { "Yellow" })
                Write-StatusLine "Items" "$($s.Items)" "White"
                Write-StatusLine "Size" "$($s.Size)" "White"
                $qd = if ($s.ContentMatchQuery) { $s.ContentMatchQuery } else { "(all)" }
                Write-StatusLine "Query" $qd "Gray"
            }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function Run-CaseSearch {
    $search = Find-ComplianceSearchByName
    if ($null -eq $search) { Pause-ForUser; return }

    if (Confirm-Action "Run search '$($search.Name)'?") {
        Invoke-ComplianceSearchStart -SearchName $search.Name | Out-Null
    }
    Pause-ForUser
}

function Remove-CaseSearchFlow {
    $search = Find-ComplianceSearchByName
    if ($null -eq $search) { Pause-ForUser; return }

    if (Confirm-Action "DELETE search '$($search.Name)'?") {
        try {
            $actions = @(Get-ComplianceSearchAction -ErrorAction SilentlyContinue | Where-Object { $_.SearchName -eq $search.Name })
            foreach ($a in $actions) { Remove-ComplianceSearchAction -Identity $a.Name -Confirm:$false -ErrorAction SilentlyContinue }
            Remove-ComplianceSearch -Identity $search.Name -Confirm:$false -ErrorAction Stop
            Write-Success "Deleted."
        } catch { Write-ErrorMsg "Failed: $_" }
    }
    Pause-ForUser
}

# ---- Hold Management ----

function Start-HoldManagement {
    $keepGoing = $true
    while ($keepGoing) {
        $action = Show-Menu -Title "Hold Management" -Options @(
            "Create a hold in a case",
            "List holds in a case",
            "Modify a hold (add/remove locations)",
            "Remove a hold"
        ) -BackLabel "Back"

        switch ($action) {
            0 { New-HoldFlow }
            1 { List-CaseHolds }
            2 { Modify-HoldFlow }
            3 { Remove-HoldFlow }
            -1 { $keepGoing = $false }
        }
    }
}

function New-HoldFlow {
    Write-SectionHeader "Create Legal Hold"

    $case = Find-Case -StatusFilter "Active"
    if ($null -eq $case) { Pause-ForUser; return }

    $name = Read-UserInput "Hold name"
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    # Locations to hold
    Write-SectionHeader "Hold Locations"
    $mailboxes = @()
    Write-InfoMsg "Add mailboxes to place on hold (or 'done' when finished):"
    while ($true) {
        $mb = Read-UserInput "Mailbox email (or 'done')"
        if ($mb -match '^done$') { break }
        if ($mb) { $mailboxes += $mb }
    }

    if ($mailboxes.Count -eq 0) {
        Write-Warn "No mailboxes specified for hold."
        Pause-ForUser; return
    }

    # Optional KQL query for the hold
    $holdQuery = Read-UserInput "Hold query (KQL to hold only matching content, or Enter for all content)"

    $details = "Case: $($case.Name)`nHold: $name`nMailboxes: $($mailboxes -join ', ')`nQuery: $(if ($holdQuery) { $holdQuery } else { '(all content)' })"
    if (-not (Confirm-Action "Create this hold?" $details)) { return }

    try {
        # Create hold policy
        $policyParams = @{
            Name             = $name
            Case             = $case.Name
            ExchangeLocation = $mailboxes
            Enabled          = $true
        }
        New-CaseHoldPolicy @policyParams -ErrorAction Stop | Out-Null
        Write-Success "Hold policy '$name' created."

        # Create hold rule with query if provided
        if ($holdQuery) {
            $ruleParams = @{
                Name              = "${name}_Rule"
                Policy            = $name
                ContentMatchQuery = $holdQuery
            }
            New-CaseHoldRule @ruleParams -ErrorAction Stop | Out-Null
            Write-Success "Hold rule created with query filter."
        }

        Write-Success "Legal hold is now active."
        Write-Warn "Held mailboxes will retain all content matching the criteria."
    } catch { Write-ErrorMsg "Failed to create hold: $_" }
    Pause-ForUser
}

function List-CaseHolds {
    $case = Find-Case
    if ($null -eq $case) { Pause-ForUser; return }

    Write-SectionHeader "Holds in '$($case.Name)'"
    try {
        $holds = @(Get-CaseHoldPolicy -Case $case.Name -ErrorAction Stop)
        if ($holds.Count -eq 0) { Write-InfoMsg "No holds in this case."; Pause-ForUser; return }

        foreach ($h in $holds) {
            Write-Host ""
            Write-StatusLine "Name" $h.Name "Cyan"
            Write-StatusLine "Enabled" "$($h.Enabled)" $(if ($h.Enabled) { "Green" } else { "Red" })
            Write-StatusLine "Exchange" ($h.ExchangeLocation -join "; ") "White"
            if ($h.SharePointLocation) {
                Write-StatusLine "SharePoint" ($h.SharePointLocation -join "; ") "White"
            }

            # Show rule/query
            try {
                $rules = @(Get-CaseHoldRule -Policy $h.Name -ErrorAction SilentlyContinue)
                foreach ($r in $rules) {
                    if ($r.ContentMatchQuery) {
                        Write-StatusLine "Query" $r.ContentMatchQuery "Gray"
                    }
                }
            } catch {}
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function Modify-HoldFlow {
    Write-SectionHeader "Modify Hold"

    $case = Find-Case -StatusFilter "Active"
    if ($null -eq $case) { Pause-ForUser; return }

    try {
        $holds = @(Get-CaseHoldPolicy -Case $case.Name -ErrorAction Stop)
        if ($holds.Count -eq 0) { Write-InfoMsg "No holds."; Pause-ForUser; return }

        $holdLabels = $holds | ForEach-Object { "$($_.Name) [Enabled: $($_.Enabled)]" }
        $sel = Show-Menu -Title "Select hold to modify" -Options $holdLabels -BackLabel "Cancel"
        if ($sel -eq -1) { Pause-ForUser; return }
        $hold = $holds[$sel]

        $modAction = Show-Menu -Title "Modify '$($hold.Name)'" -Options @(
            "Add mailbox(es) to hold",
            "Remove mailbox(es) from hold",
            "Enable / disable hold",
            "Update hold query"
        ) -BackLabel "Cancel"

        switch ($modAction) {
            0 {
                $newMb = @()
                while ($true) {
                    $mb = Read-UserInput "Mailbox to add (or 'done')"
                    if ($mb -match '^done$') { break }; if ($mb) { $newMb += $mb }
                }
                if ($newMb.Count -gt 0 -and (Confirm-Action "Add $($newMb.Count) mailbox(es) to hold?")) {
                    try {
                        Set-CaseHoldPolicy -Identity $hold.Name -AddExchangeLocation $newMb -ErrorAction Stop
                        Write-Success "Mailboxes added."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
            1 {
                $currentLocs = @($hold.ExchangeLocation)
                if ($currentLocs.Count -eq 0) { Write-InfoMsg "No locations on hold."; break }
                $locLabels = $currentLocs | ForEach-Object { "$_" }
                $remSel = Show-MultiSelect -Title "Select to remove" -Options $locLabels
                $removeMb = $remSel | ForEach-Object { $currentLocs[$_] }
                if (Confirm-Action "Remove $($removeMb.Count) mailbox(es) from hold?") {
                    try {
                        Set-CaseHoldPolicy -Identity $hold.Name -RemoveExchangeLocation $removeMb -ErrorAction Stop
                        Write-Success "Mailboxes removed from hold."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
            2 {
                $newState = -not $hold.Enabled
                if (Confirm-Action "Set hold to Enabled=$newState?") {
                    try {
                        Set-CaseHoldPolicy -Identity $hold.Name -Enabled $newState -ErrorAction Stop
                        Write-Success "Hold is now $(if ($newState) { 'enabled' } else { 'disabled' })."
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
            3 {
                $newQuery = Read-UserInput "New KQL query (or 'clear' for all content)"
                if (Confirm-Action "Update hold query?") {
                    try {
                        $rules = @(Get-CaseHoldRule -Policy $hold.Name -ErrorAction Stop)
                        if ($rules.Count -gt 0) {
                            $qVal = if ($newQuery -eq 'clear') { "" } else { $newQuery }
                            Set-CaseHoldRule -Identity $rules[0].Name -ContentMatchQuery $qVal -ErrorAction Stop
                            Write-Success "Hold query updated."
                        } else {
                            if ($newQuery -and $newQuery -ne 'clear') {
                                New-CaseHoldRule -Name "$($hold.Name)_Rule" -Policy $hold.Name -ContentMatchQuery $newQuery -ErrorAction Stop
                                Write-Success "Hold rule created."
                            }
                        }
                    } catch { Write-ErrorMsg "Failed: $_" }
                }
            }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

function Remove-HoldFlow {
    Write-SectionHeader "Remove Hold"

    $case = Find-Case -StatusFilter "Active"
    if ($null -eq $case) { Pause-ForUser; return }

    try {
        $holds = @(Get-CaseHoldPolicy -Case $case.Name -ErrorAction Stop)
        if ($holds.Count -eq 0) { Write-InfoMsg "No holds."; Pause-ForUser; return }

        $holdLabels = $holds | ForEach-Object { "$($_.Name) [Enabled: $($_.Enabled)]" }
        $sel = Show-Menu -Title "Select hold to remove" -Options $holdLabels -BackLabel "Cancel"
        if ($sel -eq -1) { Pause-ForUser; return }
        $hold = $holds[$sel]

        Write-Warn "This releases all held content in the mailboxes."
        if (Confirm-Action "REMOVE hold '$($hold.Name)'?") {
            try {
                # Remove rules first
                $rules = @(Get-CaseHoldRule -Policy $hold.Name -ErrorAction SilentlyContinue)
                foreach ($r in $rules) { Remove-CaseHoldRule -Identity $r.Name -Confirm:$false -ErrorAction SilentlyContinue }
                Remove-CaseHoldPolicy -Identity $hold.Name -Confirm:$false -ErrorAction Stop
                Write-Success "Hold '$($hold.Name)' removed."
            } catch { Write-ErrorMsg "Failed: $_" }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

# ---- Export Status ----

function Show-ExportStatus {
    Write-SectionHeader "Export / Action Status"
    try {
        $actions = @(Get-ComplianceSearchAction -ErrorAction Stop)
        if ($actions.Count -eq 0) {
            Write-InfoMsg "No export or preview actions found."
            Pause-ForUser; return
        }

        foreach ($a in $actions) {
            $statusColor = switch ($a.Status) {
                "Completed" { "Green" }
                "InProgress" { "Yellow" }
                "Starting"   { "Yellow" }
                default      { "Gray" }
            }
            Write-Host ""
            Write-StatusLine "Name" $a.Name "Cyan"
            Write-StatusLine "Action" $a.Action "White"
            Write-StatusLine "Search" $a.SearchName "White"
            Write-StatusLine "Status" $a.Status $statusColor
            Write-StatusLine "Created" "$($a.CreatedTime)" "Gray"
            if ($a.Results -and $a.Action -eq "Preview") {
                $previewCount = ($a.Results -split "`n").Count
                Write-StatusLine "Preview Lines" "$previewCount" "White"
            }
        }
    } catch { Write-ErrorMsg "Failed: $_" }
    Pause-ForUser
}

# ============================================================
#  Shared Helpers
# ============================================================

function Find-ComplianceSearchByName {
    try {
        $searches = @(Get-ComplianceSearch -ErrorAction Stop)
        if ($searches.Count -eq 0) { Write-InfoMsg "No searches found."; return $null }
        $labels = $searches | ForEach-Object { "$($_.Name) [$($_.Status)] ($($_.Items) items)" }
        $sel = Show-Menu -Title "Select Search" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return $searches[$sel]
    } catch { Write-ErrorMsg "Failed to list searches: $_"; return $null }
}

function Find-Case {
    param([string]$StatusFilter = "")

    try {
        $cases = @(Get-ComplianceCase -ErrorAction Stop)
        if ($StatusFilter) {
            $cases = @($cases | Where-Object { $_.Status -eq $StatusFilter })
        }
        if ($cases.Count -eq 0) {
            if ($StatusFilter) { Write-InfoMsg "No $StatusFilter cases found." }
            else { Write-InfoMsg "No cases found." }
            return $null
        }

        $labels = $cases | ForEach-Object { "$($_.Name) [$($_.Status)] ($($_.CaseType))" }
        $sel = Show-Menu -Title "Select Case" -Options $labels -BackLabel "Cancel"
        if ($sel -eq -1) { return $null }
        return $cases[$sel]
    } catch { Write-ErrorMsg "Failed to list cases: $_"; return $null }
}

function Invoke-ComplianceSearchStart {
    <#
    .SYNOPSIS
        Wraps Start-ComplianceSearch with auto-reconnect if search session missing.
    #>
    param([string]$SearchName)

    try {
        Start-ComplianceSearch -Identity $SearchName -ErrorAction Stop
        Write-Success "Search '$SearchName' started."
        return $true
    } catch {
        $startErr = $_.Exception.Message
        if ($startErr -match "EnableSearchOnlySession|search initialization|Connect-IPPSSession") {
            Write-Warn "SCC session needs search-only mode. Reconnecting..."
            if (Invoke-SCCReconnect) {
                Write-InfoMsg "Retrying search start..."
                try {
                    Start-ComplianceSearch -Identity $SearchName -ErrorAction Stop
                    Write-Success "Search '$SearchName' started."
                    return $true
                } catch {
                    Write-ErrorMsg "Still failed after reconnect: $_"
                    return $false
                }
            } else {
                Write-ErrorMsg "Could not reconnect. Restart the tool and try again."
                return $false
            }
        } else {
            Write-ErrorMsg "Failed to start search: $_"
            return $false
        }
    }
}

function Invoke-SCCReconnect {
    <#
    .SYNOPSIS
        Disconnects and reconnects SCC with -EnableSearchOnlySession.
        Returns $true on success.
    #>
    Write-InfoMsg "Disconnecting current SCC session..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    } catch {}

    $script:SessionState.ComplianceCenter = $false
    $script:SessionState.ExchangeOnline = $false

    Write-InfoMsg "Reconnecting with search-only session..."
    $sccParams = @{
        ShowBanner              = $false
        EnableSearchOnlySession = $true
    }
    if ($script:SessionState.TenantMode -eq "Partner" -and $script:SessionState.TenantDomain) {
        $sccParams["DelegatedOrganization"] = $script:SessionState.TenantDomain
    }

    try {
        Connect-IPPSSession @sccParams -ErrorAction Stop
        $script:SessionState.ComplianceCenter = $true
        Write-Success "SCC reconnected with search session."
        return $true
    } catch {
        Write-ErrorMsg "SCC reconnect failed: $_"
        return $false
    }
}
