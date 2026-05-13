# ============================================================
#  AuditViewer.ps1 — read, filter, paginate, export audit logs
#
#  Two log streams on disk:
#    %LOCALAPPDATA%\M365Manager\audit\session-*.log  (general)
#    %LOCALAPPDATA%\M365Manager\audit\mark-*.log     (AI assistant)
#
#  Phase 2 onward, session-*.log is JSON-per-line. Older lines
#  (and all mark-*.log entries) are human-readable. The parser
#  below handles both transparently and normalizes to a uniform
#  PSCustomObject shape.
# ============================================================

function ConvertFrom-AuditLine {
    <#
        Parse one line of an audit log into a normalized record.
        Returns $null on unparseable input.
    #>
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }
    $trimmed = $Line.TrimStart()

    # 1. JSON line
    if ($trimmed.StartsWith('{')) {
        try {
            $obj = $trimmed | ConvertFrom-Json -ErrorAction Stop
            $ts = $null
            if ($obj.ts) { [DateTime]::TryParse($obj.ts, [ref]$ts) | Out-Null }
            return [PSCustomObject]@{
                ts           = $ts
                entryId      = $obj.entryId
                mode         = $obj.mode
                event        = $obj.event
                description  = $obj.description
                actionType   = $obj.actionType
                target       = $obj.target
                result       = $obj.result
                error        = $obj.error
                tenant       = $obj.tenant
                session      = $obj.session
                reverse      = $obj.reverse
                noUndoReason = $obj.noUndoReason
                source       = 'jsonl'
                raw          = $trimmed
            }
        } catch { return $null }
    }

    # 2. Legacy session line: [ts] [event] [MODE=X] detail
    if ($trimmed -match '^\[([^\]]+)\]\s*\[([^\]]+)\]\s*\[MODE=(\w+)\]\s*(.*)$') {
        $ts = $null; [DateTime]::TryParse($Matches[1], [ref]$ts) | Out-Null
        return [PSCustomObject]@{
            ts           = $ts
            entryId      = $null
            mode         = $Matches[3]
            event        = $Matches[2]
            description  = $Matches[4]
            actionType   = $null
            target       = $null
            result       = $null
            error        = $null
            tenant       = $null
            session      = $null
            reverse      = $null
            noUndoReason = $null
            source       = 'legacy-session'
            raw          = $trimmed
        }
    }

    # 3. AI assistant log: [ts] [event] detail (no MODE= tag)
    if ($trimmed -match '^\[([^\]]+)\]\s*\[([^\]]+)\]\s*(.*)$') {
        $ts = $null; [DateTime]::TryParse($Matches[1], [ref]$ts) | Out-Null
        return [PSCustomObject]@{
            ts           = $ts
            entryId      = $null
            mode         = $null
            event        = $Matches[2]
            description  = $Matches[3]
            actionType   = 'AICmd'
            target       = $null
            result       = $null
            error        = $null
            tenant       = $null
            session      = $null
            reverse      = $null
            noUndoReason = $null
            source       = 'ai-mark'
            raw          = $trimmed
        }
    }
    return $null
}

function Read-AuditEntries {
    <#
        Reads one or more audit log files and returns the parsed
        records sorted by timestamp ascending. When -Path is
        omitted, walks the default audit directory; pass
        -IncludeAiLog to also pull mark-*.log entries.
    #>
    param(
        [string[]]$Path,
        [switch]$IncludeAiLog
    )
    if (-not $Path -or $Path.Count -eq 0) {
        $dir = Get-AuditLogDirectory
        if (-not (Test-Path -LiteralPath $dir)) { return @() }
        $Path = @(Get-ChildItem -LiteralPath $dir -Filter 'session-*.log' -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
        if ($IncludeAiLog) {
            $Path += @(Get-ChildItem -LiteralPath $dir -Filter 'mark-*.log' -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
        }
    }
    $entries = New-Object System.Collections.ArrayList
    foreach ($p in $Path) {
        try {
            $lines = Get-Content -LiteralPath $p -ErrorAction Stop
            foreach ($line in $lines) {
                $e = ConvertFrom-AuditLine -Line $line
                if ($e) { [void]$entries.Add($e) }
            }
        } catch { Write-Warn "Could not read '$p': $_" }
    }
    return @($entries | Sort-Object { $_.ts })
}

function Test-AuditEntryMatchesUser {
    <#
        Match a UPN against an entry. Prefer structured target
        fields; fall back to regex scan of the description (covers
        legacy human-readable lines).
    #>
    param($Entry, [string]$Upn)
    if (-not $Upn) { return $true }
    $u = $Upn.ToLowerInvariant()
    if ($Entry.target) {
        foreach ($k in 'upn','userUpn','userPrincipalName','identity','member','trustee','manager') {
            if ($Entry.target.$k -and ([string]$Entry.target.$k).ToLowerInvariant() -eq $u) { return $true }
        }
    }
    if ($Entry.description -and $Entry.description.ToLowerInvariant().Contains($u)) { return $true }
    return $false
}

function Test-AuditEntryMatchesTarget {
    <#
        Match a free-form target string (group name, mailbox alias,
        sku part number, etc.) against an entry. Substring match
        on description; case-insensitive scan of target's values.
    #>
    param($Entry, [string]$TargetText)
    if (-not $TargetText) { return $true }
    $t = $TargetText.ToLowerInvariant()
    if ($Entry.description -and $Entry.description.ToLowerInvariant().Contains($t)) { return $true }
    if ($Entry.target) {
        foreach ($p in $Entry.target.PSObject.Properties) {
            if ($p.Value -is [string] -and $p.Value.ToLowerInvariant().Contains($t)) { return $true }
        }
    }
    return $false
}

function Filter-AuditEntries {
    <#
        Apply a filter hashtable to an entry array. Keys (any
        omitted/null means "no filter on that field"):
          User       : UPN (matches target or description)
          From       : DateTime  (UTC; entries strictly before are dropped)
          To         : DateTime  (UTC; entries strictly after are dropped)
          ActionType : exact match (case-insensitive)
          EventType  : exact match (case-insensitive)
          Mode       : LIVE | PREVIEW
          Result     : success | failure | preview | info
          Target     : substring on description / target values
    #>
    param([array]$Entries, [hashtable]$Filter)
    if (-not $Filter -or $Filter.Count -eq 0) { return $Entries }
    return @($Entries | Where-Object {
        $e = $_
        if ($Filter.From -and $e.ts -and $e.ts -lt $Filter.From)           { return $false }
        if ($Filter.To   -and $e.ts -and $e.ts -gt $Filter.To)             { return $false }
        if ($Filter.Mode       -and $e.mode       -and ($e.mode      -ne $Filter.Mode))         { return $false }
        if ($Filter.Mode       -and -not $e.mode  -and $Filter.Mode -ne 'ANY')                  { return $false }
        if ($Filter.EventType  -and $e.event      -and ($e.event     -notlike $Filter.EventType)) { return $false }
        if ($Filter.ActionType -and $e.actionType -and ($e.actionType -notlike $Filter.ActionType)) { return $false }
        if ($Filter.ActionType -and -not $e.actionType) { return $false }
        if ($Filter.Result     -and $e.result     -and ($e.result    -ne $Filter.Result))       { return $false }
        if (-not (Test-AuditEntryMatchesUser   -Entry $e -Upn $Filter.User))        { return $false }
        if (-not (Test-AuditEntryMatchesTarget -Entry $e -TargetText $Filter.Target)) { return $false }
        return $true
    })
}

function Format-AuditEntryRow {
    <#
        Returns a short string representation for table display.
        Truncates description to keep the row within ~120 chars.
    #>
    param($Entry, [int]$DescWidth = 60)
    $tsText = if ($Entry.ts) { $Entry.ts.ToString('yyyy-MM-dd HH:mm:ss') } else { '???' }
    $modeText = if ($Entry.mode) { $Entry.mode.PadRight(7) } else { '-------' }
    $evtText  = if ($Entry.event) { $Entry.event.PadRight(8) } else { '--------' }
    $resText  = if ($Entry.result) { $Entry.result.PadRight(8) } else { '--------' }
    $desc = if ($Entry.description) { $Entry.description } else { '' }
    if ($desc.Length -gt $DescWidth) { $desc = $desc.Substring(0, $DescWidth - 3) + '...' }
    return "{0} {1} {2} {3} {4}" -f $tsText, $modeText, $evtText, $resText, $desc
}

function Show-AuditPage {
    param([array]$Entries, [int]$Page = 0, [int]$PageSize = 20)
    $total = $Entries.Count
    if ($total -eq 0) { Write-InfoMsg "(no entries match the current filter)"; return }
    $start = $Page * $PageSize
    if ($start -ge $total) { $start = [Math]::Max(0, $total - $PageSize) }
    $end = [Math]::Min($start + $PageSize - 1, $total - 1)

    Write-Host ""
    Write-Host ("  Page {0} of {1}   (entries {2}-{3} of {4})" -f ($Page + 1), [Math]::Ceiling($total / $PageSize), ($start + 1), ($end + 1), $total) -ForegroundColor DarkGray
    Write-Host ("  " + ('-' * 110)) -ForegroundColor DarkGray
    Write-Host ("  TIMESTAMP            MODE    EVENT    RESULT   DESCRIPTION") -ForegroundColor DarkGray
    for ($i = $start; $i -le $end; $i++) {
        $row = Format-AuditEntryRow -Entry $Entries[$i]
        $colour = 'White'
        switch ($Entries[$i].result) {
            'failure' { $colour = 'Red' }
            'preview' { $colour = 'Yellow' }
            'success' { $colour = 'Green' }
        }
        Write-Host ("  " + $row) -ForegroundColor $colour
    }
}

function Export-AuditEntriesCsv {
    param([Parameter(Mandatory)][array]$Entries, [Parameter(Mandatory)][string]$Path)
    $rows = $Entries | ForEach-Object {
        [PSCustomObject]@{
            Timestamp    = if ($_.ts) { $_.ts.ToString('o') } else { '' }
            EntryId      = $_.entryId
            Mode         = $_.mode
            Event        = $_.event
            Result       = $_.result
            ActionType   = $_.actionType
            Description  = $_.description
            Error        = $_.error
            Tenant       = $_.tenant
            TargetJson   = if ($_.target) { ($_.target | ConvertTo-Json -Compress -Depth 5) } else { '' }
            ReverseType  = if ($_.reverse) { $_.reverse.type } else { '' }
            NoUndoReason = $_.noUndoReason
            Source       = $_.source
        }
    }
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Force
}

function Export-AuditEntriesHtml {
    param([Parameter(Mandatory)][array]$Entries, [Parameter(Mandatory)][string]$Path, [hashtable]$Filter)
    $filterParts = @()
    if ($Filter) {
        foreach ($k in $Filter.Keys) {
            $v = $Filter[$k]
            if ($null -ne $v -and "$v") { $filterParts += "<b>$k</b>=$v" }
        }
    }
    $filterText = if ($filterParts.Count -gt 0) { $filterParts -join ' &middot; ' } else { '(no filter)' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<title>M365 Manager audit export</title>')
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine('body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;background:#f5f7fa;color:#222;margin:24px}')
    [void]$sb.AppendLine('h1{font-size:18px;margin:0 0 4px 0}')
    [void]$sb.AppendLine('.meta{color:#666;font-size:13px;margin-bottom:16px}')
    [void]$sb.AppendLine('table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.06)}')
    [void]$sb.AppendLine('th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #e6e8eb;vertical-align:top}')
    [void]$sb.AppendLine('th{background:#eef2f5;font-weight:600;position:sticky;top:0}')
    [void]$sb.AppendLine('tr:hover{background:#fafbfc}')
    [void]$sb.AppendLine('td.desc{max-width:520px;word-break:break-word}')
    [void]$sb.AppendLine('.r-success{color:#0a7e2d}.r-failure{color:#b00020}.r-preview{color:#a76600}')
    [void]$sb.AppendLine('.m-PREVIEW{background:#fff8e1}.m-LIVE{background:#fff}')
    [void]$sb.AppendLine('</style></head><body>')
    [void]$sb.AppendLine('<h1>M365 Manager audit export</h1>')
    [void]$sb.AppendLine("<div class=meta>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') &middot; $($Entries.Count) entr$(if($Entries.Count -eq 1){'y'}else{'ies'}) &middot; filter: $filterText</div>")
    [void]$sb.AppendLine('<table><thead><tr>')
    [void]$sb.AppendLine('<th>Timestamp (UTC)</th><th>Mode</th><th>Event</th><th>Result</th><th>Action</th><th>Description</th><th>Error</th>')
    [void]$sb.AppendLine('</tr></thead><tbody>')
    foreach ($e in $Entries) {
        $ts = if ($e.ts) { $e.ts.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        $modeClass = if ($e.mode) { "m-$($e.mode)" } else { '' }
        $resClass  = if ($e.result) { "r-$($e.result)" } else { '' }
        $desc = [System.Net.WebUtility]::HtmlEncode([string]$e.description)
        $err  = [System.Net.WebUtility]::HtmlEncode([string]$e.error)
        [void]$sb.AppendLine("<tr class='$modeClass'><td>$ts</td><td>$($e.mode)</td><td>$($e.event)</td><td class='$resClass'>$($e.result)</td><td>$($e.actionType)</td><td class=desc>$desc</td><td class='r-failure'>$err</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table></body></html>')
    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
}

function Read-AuditFilterFromOperator {
    <#
        Prompt the operator for filter values. Blank inputs leave a
        filter unset. Returns a hashtable consumable by
        Filter-AuditEntries.
    #>
    $f = @{}
    $u = Read-UserInput "User UPN (blank for all)"
    if ($u) { $f.User = $u.Trim() }
    $d = Read-UserInput "Date range (e.g. '7d', '24h', '2026-05-01 / 2026-05-12'; blank = all time)"
    if ($d) {
        if ($d -match '^(?<n>\d+)(?<u>[hd])$') {
            $n = [int]$Matches['n']
            $f.From = if ($Matches['u'] -eq 'h') { (Get-Date).ToUniversalTime().AddHours(-$n) } else { (Get-Date).ToUniversalTime().AddDays(-$n) }
            $f.To = (Get-Date).ToUniversalTime()
        } elseif ($d -match '^(?<a>\S+)\s*[/-]\s*(?<b>\S+)$') {
            $a = $null; $b = $null
            [DateTime]::TryParse($Matches['a'], [ref]$a) | Out-Null
            [DateTime]::TryParse($Matches['b'], [ref]$b) | Out-Null
            if ($a) { $f.From = $a.ToUniversalTime() }
            if ($b) { $f.To = $b.ToUniversalTime() }
        } else {
            Write-Warn "Unrecognized date range '$d' -- ignored."
        }
    }
    $at = Read-UserInput "Action type (e.g. AssignLicense, AddToGroup; blank = any)"
    if ($at) { $f.ActionType = $at.Trim() }
    $ev = Read-UserInput "Event type (EXEC, PREVIEW, SESSION_START, etc; blank = any)"
    if ($ev) { $f.EventType = $ev.Trim().ToUpper() }
    $m = Read-UserInput "Mode (LIVE / PREVIEW / blank for both)"
    if ($m) { $f.Mode = $m.Trim().ToUpper() }
    $r = Read-UserInput "Result (success / failure / preview / blank for any)"
    if ($r) { $f.Result = $r.Trim().ToLower() }
    $t = Read-UserInput "Target text substring (group / mailbox / SKU; blank = any)"
    if ($t) { $f.Target = $t.Trim() }
    return $f
}

function Show-AuditLogViewer {
    <#
        Interactive audit log viewer. Loads every session-*.log
        (and optionally mark-*.log), lets the operator filter and
        paginate, exports CSV / HTML on demand.
    #>
    [CmdletBinding()]
    param([switch]$IncludeAiLog)

    Write-SectionHeader "Audit Log Viewer"
    Write-InfoMsg "Loading audit logs from $(Get-AuditLogDirectory)..."
    $all = Read-AuditEntries -IncludeAiLog:$IncludeAiLog
    Write-InfoMsg "$($all.Count) total entr$(if($all.Count -eq 1){'y'}else{'ies'}) loaded."
    if ($all.Count -eq 0) {
        Write-Warn "No audit data yet. Run something in LIVE or PREVIEW mode first."
        Pause-ForUser; return
    }

    $filter = @{}
    $filtered = $all
    $page = 0
    $pageSize = 20

    while ($true) {
        Write-Host ""
        Write-Host "  Filter: " -ForegroundColor DarkGray -NoNewline
        if ($filter.Count -eq 0) { Write-Host "(none)" -ForegroundColor DarkGray }
        else {
            $parts = @()
            foreach ($k in $filter.Keys) { $parts += "$k=$($filter[$k])" }
            Write-Host ($parts -join '  ') -ForegroundColor White
        }
        Show-AuditPage -Entries $filtered -Page $page -PageSize $pageSize
        Write-Host ""
        Write-Host "  [N]ext  [P]rev  [F]ilter  [C]lear-filter  [D]etail  [E]xport-CSV  [H]TML-export  [Q]uit" -ForegroundColor $script:Colors.Highlight -NoNewline
        Write-Host ": " -NoNewline
        $cmd = (Read-Host).Trim().ToLower()
        switch -Regex ($cmd) {
            '^n' { if (($page + 1) * $pageSize -lt $filtered.Count) { $page++ } else { Write-InfoMsg "At last page." } }
            '^p' { if ($page -gt 0) { $page-- } else { Write-InfoMsg "At first page." } }
            '^f' { $filter = Read-AuditFilterFromOperator; $filtered = Filter-AuditEntries -Entries $all -Filter $filter; $page = 0; Write-InfoMsg "$($filtered.Count) entr$(if($filtered.Count -eq 1){'y'}else{'ies'}) match." }
            '^c' { $filter = @{}; $filtered = $all; $page = 0; Write-InfoMsg "Filter cleared." }
            '^d' {
                $idxText = Read-UserInput "Row number on this page (1-$pageSize)"
                $idx = 0
                if ([int]::TryParse($idxText, [ref]$idx)) {
                    $global = $page * $pageSize + ($idx - 1)
                    if ($global -ge 0 -and $global -lt $filtered.Count) {
                        Show-AuditEntryDetail -Entry $filtered[$global]
                    } else { Write-Warn "Out of range." }
                }
            }
            '^e' {
                $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                $out = Join-Path (Get-AuditLogDirectory) "audit-export-$stamp.csv"
                Export-AuditEntriesCsv -Entries $filtered -Path $out
                Write-Success "CSV: $out"
            }
            '^h' {
                $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                $out = Join-Path (Get-AuditLogDirectory) "audit-export-$stamp.html"
                Export-AuditEntriesHtml -Entries $filtered -Path $out -Filter $filter
                Write-Success "HTML: $out"
            }
            '^q' { return }
            default { Write-InfoMsg "Unknown command." }
        }
    }
}

function Show-AuditEntryDetail {
    param([Parameter(Mandatory)]$Entry)
    $orDash = { param($v) if ($v) { [string]$v } else { '-' } }
    Write-Host ""
    Write-Host "  ---- entry detail ----" -ForegroundColor $script:Colors.Title
    Write-StatusLine "Timestamp"    $(if ($Entry.ts) { $Entry.ts.ToString('o') } else { '(unknown)' }) "White"
    Write-StatusLine "EntryId"      (& $orDash $Entry.entryId)     "Gray"
    Write-StatusLine "Mode"         (& $orDash $Entry.mode)        "White"
    Write-StatusLine "Event"        (& $orDash $Entry.event)       "White"
    Write-StatusLine "Result"       (& $orDash $Entry.result)      $(if ($Entry.result -eq 'failure') {'Red'} elseif ($Entry.result -eq 'preview') {'Yellow'} else {'Green'})
    Write-StatusLine "ActionType"   (& $orDash $Entry.actionType)  "Cyan"
    Write-StatusLine "Tenant"       (& $orDash $Entry.tenant)      "DarkGray"
    Write-StatusLine "Description"  (& $orDash $Entry.description) "White"
    if ($Entry.error)        { Write-StatusLine "Error"        $Entry.error 'Red' }
    if ($Entry.noUndoReason) { Write-StatusLine "NoUndoReason" $Entry.noUndoReason 'Yellow' }
    if ($Entry.target) {
        Write-Host "    target:" -ForegroundColor DarkGray
        ($Entry.target | ConvertTo-Json -Depth 5 -Compress) -split "(.{0,90})" | Where-Object { $_ } | ForEach-Object {
            Write-Host "      $_" -ForegroundColor White
        }
    }
    if ($Entry.reverse) {
        Write-Host "    reverse:" -ForegroundColor DarkGray
        Write-Host ("      type        : " + $Entry.reverse.type) -ForegroundColor White
        Write-Host ("      description : " + $Entry.reverse.description) -ForegroundColor White
    }
    Write-Host ""
}

function Start-AuditReportingMenu {
    <#
        Top-level "Audit & Reporting" entry point. Phase 2 commits
        add Undo, SignInLookup, UnifiedAuditLog -- those wire into
        this menu in their respective commits.
    #>
    while ($true) {
        $opts = @(
            "Audit log viewer (filter / page / export)",
            "Undo recent operation...",
            "Sign-in lookup (search Graph signIns)",
            "Sign-in lookup (recent activity for one user)",
            "Unified audit log (Search-UnifiedAuditLog)",
            "Guest user reports...",
            "Reporting (existing reports menu)"
        )
        $sel = Show-Menu -Title "Audit & Reporting" -Options $opts -BackLabel "Back to Main Menu"
        switch ($sel) {
            0  { Show-AuditLogViewer }
            1  { if (Get-Command Start-UndoMenu -ErrorAction SilentlyContinue) { Start-UndoMenu } else { Write-Warn "Undo menu unavailable."; Pause-ForUser } }
            2  { if (Get-Command Start-SignInSearch -ErrorAction SilentlyContinue) { Start-SignInSearch } else { Write-Warn "Sign-in lookup unavailable."; Pause-ForUser } }
            3  { if (Get-Command Show-UserRecentActivity -ErrorAction SilentlyContinue) { Show-UserRecentActivity } else { Write-Warn "Sign-in lookup unavailable."; Pause-ForUser } }
            4  { if (Get-Command Start-UnifiedAuditSearch -ErrorAction SilentlyContinue) { Start-UnifiedAuditSearch } else { Write-Warn "Unified audit log unavailable."; Pause-ForUser } }
            5  { if (Get-Command Start-GuestUsersMenu -ErrorAction SilentlyContinue) { Start-GuestUsersMenu } else { Write-Warn "Guest user module unavailable."; Pause-ForUser } }
            6  { if (Get-Command Start-ReportingMenu -ErrorAction SilentlyContinue) { Start-ReportingMenu } else { Write-Warn "Reporting menu unavailable."; Pause-ForUser } }
            -1 { return }
        }
    }
}
