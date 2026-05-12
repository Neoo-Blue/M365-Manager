# ============================================================
#  SignInLookup.ps1 — Microsoft Graph sign-in log lookup
#
#  Wraps GET /v1.0/auditLogs/signIns. Requires the connected
#  Graph context to hold AuditLog.Read.All and Directory.Read.All
#  scopes -- the module warns and returns early if either is
#  missing.
#
#  Two surface modes:
#    Start-SignInSearch        : full filter wizard
#    Show-UserRecentActivity   : UPN one-shot, last 30 days
#
#  Pagination: walks @odata.nextLink until exhausted or until
#  -MaxResults is hit. Graph caps a page at 1000 rows.
#
#  Exports: table + CSV + HTML, written next to the audit dir.
# ============================================================

$script:SignInRequiredScopes = @('AuditLog.Read.All','Directory.Read.All')

function Assert-SignInPermissions {
    <#
        Read the active Graph context's scopes; warn and return
        $false if either required scope is missing.
    #>
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) {
        Write-Warn "Microsoft.Graph.Authentication module not loaded. Connect via Connect-ForTask first."
        return $false
    }
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Warn "Not connected to Microsoft Graph. Connect first via the main menu (or this flow will trigger Connect-ForTask)."
        return $true   # connection will happen on Invoke-MgGraphRequest
    }
    $missing = @()
    foreach ($s in $script:SignInRequiredScopes) {
        if ($ctx.Scopes -notcontains $s) { $missing += $s }
    }
    if ($missing.Count -gt 0) {
        Write-Warn "The connected Graph app is missing scope(s): $($missing -join ', ')"
        Write-InfoMsg "Sign-in lookup will likely return 403. Reconnect after granting the scope(s) in Entra (or via Reconnect-GraphWithConsent)."
        return $false
    }
    return $true
}

function Format-SignInLocation {
    param($Location)
    if (-not $Location) { return '' }
    $city    = if ($Location.city)    { [string]$Location.city }    else { '' }
    $country = if ($Location.countryOrRegion) { [string]$Location.countryOrRegion } else { '' }
    if ($city -and $country) { return "$city, $country" }
    if ($country) { return $country }
    if ($city)    { return $city }
    return ''
}

function ConvertTo-SignInRecord {
    <#
        Normalize one Graph sign-in event into a flat
        PSCustomObject suitable for table / CSV / HTML.
    #>
    param($Raw)
    $loc = Format-SignInLocation -Location $Raw.location
    $errCode = $null
    $errMsg  = $null
    if ($Raw.status) {
        $errCode = $Raw.status.errorCode
        $errMsg  = $Raw.status.failureReason
    }
    return [PSCustomObject]@{
        TimeUtc            = if ($Raw.createdDateTime) { ([DateTime]$Raw.createdDateTime).ToUniversalTime() } else { $null }
        UserPrincipalName  = $Raw.userPrincipalName
        UserDisplayName    = $Raw.userDisplayName
        UserId             = $Raw.userId
        App                = $Raw.appDisplayName
        Client             = $Raw.clientAppUsed
        IpAddress          = $Raw.ipAddress
        Location           = $loc
        Country            = if ($Raw.location.countryOrRegion) { $Raw.location.countryOrRegion } else { '' }
        Device             = if ($Raw.deviceDetail.operatingSystem) { "$($Raw.deviceDetail.operatingSystem) ($($Raw.deviceDetail.browser))" } else { '' }
        Risk               = $Raw.riskLevelDuringSignIn
        RiskState          = $Raw.riskState
        CAStatus           = $Raw.conditionalAccessStatus
        Outcome            = if ($errCode -eq 0 -or $errCode -eq '0' -or -not $errCode) { 'success' } else { "failure ($errCode)" }
        ErrorMessage       = $errMsg
        CorrelationId      = $Raw.correlationId
    }
}

function Search-SignIns {
    <#
        Query Graph for sign-in events. Returns an array of
        normalized records sorted by TimeUtc descending.

        -UPN          : userPrincipalName filter (exact match)
        -From / -To   : DateTime, defaults to last 7 days
        -AppName      : appDisplayName exact match
        -IP           : ipAddress exact match (string)
        -Country      : 2-letter ISO code (e.g. 'US','GB')
        -RiskLevel    : low|medium|high|hidden|none|unknownFutureValue
        -CAStatus     : success|failure|notApplied
        -OnlyFailures : convenience -- only sign-ins where errorCode != 0
        -MaxResults   : stop after this many rows (default 1000)
    #>
    param(
        [string]$UPN,
        [DateTime]$From = (Get-Date).AddDays(-7),
        [DateTime]$To   = (Get-Date),
        [string]$AppName,
        [string]$IP,
        [string]$Country,
        [string]$RiskLevel,
        [string]$CAStatus,
        [switch]$OnlyFailures,
        [int]$MaxResults = 1000
    )

    if (-not (Assert-SignInPermissions)) {
        if (-not (Confirm-Action "Run anyway (likely to 403)?")) { return @() }
    }

    $filter = @()
    $filter += "createdDateTime ge $($From.ToUniversalTime().ToString('o'))"
    $filter += "createdDateTime le $($To.ToUniversalTime().ToString('o'))"
    if ($UPN)       { $filter += "userPrincipalName eq '$($UPN -replace "'", "''")'" }
    if ($AppName)   { $filter += "appDisplayName eq '$($AppName -replace "'", "''")'" }
    if ($IP)        { $filter += "ipAddress eq '$($IP -replace "'", "''")'" }
    if ($Country)   { $filter += "location/countryOrRegion eq '$($Country -replace "'", "''")'" }
    if ($RiskLevel) { $filter += "riskLevelDuringSignIn eq '$($RiskLevel -replace "'", "''")'" }
    if ($CAStatus)  { $filter += "conditionalAccessStatus eq '$($CAStatus -replace "'", "''")'" }
    $filterText = ($filter -join ' and ')

    $page = 1000
    $query = "`$filter=$([uri]::EscapeDataString($filterText))&`$top=$page"
    $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?$query"

    $all = New-Object System.Collections.ArrayList
    do {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        } catch {
            $msg = $_.Exception.Message
            Write-ErrorMsg "Graph query failed: $msg"
            if ($msg -match '403|Forbidden|Authorization') {
                Write-InfoMsg "Most likely a permission issue. Required: $($script:SignInRequiredScopes -join ', ')"
            }
            break
        }
        foreach ($entry in $resp.value) {
            $rec = ConvertTo-SignInRecord -Raw $entry
            if ($OnlyFailures -and $rec.Outcome -like 'success*') { continue }
            [void]$all.Add($rec)
            if ($all.Count -ge $MaxResults) { break }
        }
        $uri = $resp.'@odata.nextLink'
    } while ($uri -and $all.Count -lt $MaxResults)

    return @($all | Sort-Object TimeUtc -Descending)
}

function Show-SignInTable {
    param([array]$Records)
    if (-not $Records -or $Records.Count -eq 0) {
        Write-InfoMsg "(no sign-in events match the current filter)"
        return
    }
    Write-Host ""
    Write-Host ("  TIME (UTC)           UPN                              APP               LOCATION              OUTCOME") -ForegroundColor DarkGray
    Write-Host ("  " + ('-' * 130)) -ForegroundColor DarkGray
    foreach ($r in $Records) {
        $t = if ($r.TimeUtc) { $r.TimeUtc.ToString('yyyy-MM-dd HH:mm:ss') } else { '???' }
        $u = if ($r.UserPrincipalName) { $r.UserPrincipalName } else { '<unknown>' }
        $a = if ($r.App) { $r.App } else { '' }
        $l = if ($r.Location) { $r.Location } else { '' }
        $o = $r.Outcome
        $colour = if ($o -like 'success*') { 'Green' } else { 'Red' }
        $row = ("{0} {1,-32} {2,-17} {3,-21} {4}" -f $t, ($u.PadRight(32).Substring(0,32)), ($a.PadRight(17).Substring(0,17)), ($l.PadRight(21).Substring(0,21)), $o)
        Write-Host ("  " + $row) -ForegroundColor $colour
    }
    Write-Host ""
    Write-InfoMsg "$($Records.Count) sign-in event(s)"
}

function Export-SignInsCsv {
    param([array]$Records, [string]$Path)
    $Records | Export-Csv -LiteralPath $Path -NoTypeInformation -Force
}

function Export-SignInsHtml {
    param([array]$Records, [string]$Path, [hashtable]$Filter)
    $filterParts = @()
    if ($Filter) { foreach ($k in $Filter.Keys) { if ("$($Filter[$k])") { $filterParts += "<b>$k</b>=$($Filter[$k])" } } }
    $filterText = if ($filterParts.Count -gt 0) { $filterParts -join ' &middot; ' } else { '(no filter)' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset=utf-8>')
    [void]$sb.AppendLine('<title>Sign-in lookup</title>')
    [void]$sb.AppendLine('<style>body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;background:#f5f7fa;color:#222;margin:24px}')
    [void]$sb.AppendLine('h1{font-size:18px;margin:0 0 4px 0}.meta{color:#666;font-size:13px;margin-bottom:16px}')
    [void]$sb.AppendLine('table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.06)}')
    [void]$sb.AppendLine('th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #e6e8eb;vertical-align:top}')
    [void]$sb.AppendLine('th{background:#eef2f5;font-weight:600;position:sticky;top:0}')
    [void]$sb.AppendLine('.ok{color:#0a7e2d}.fail{color:#b00020}.risk-high{background:#fde7e7}.risk-medium{background:#fff4d6}</style></head><body>')
    [void]$sb.AppendLine('<h1>Sign-in lookup</h1>')
    [void]$sb.AppendLine("<div class=meta>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') &middot; $($Records.Count) row(s) &middot; filter: $filterText</div>")
    [void]$sb.AppendLine('<table><thead><tr><th>Time (UTC)</th><th>UPN</th><th>App</th><th>Client</th><th>IP</th><th>Location</th><th>Risk</th><th>CA</th><th>Outcome</th><th>Error</th></tr></thead><tbody>')
    foreach ($r in $Records) {
        $t = if ($r.TimeUtc) { $r.TimeUtc.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        $outClass = if ($r.Outcome -like 'success*') { 'ok' } else { 'fail' }
        $riskClass = if ($r.Risk -eq 'high') { 'risk-high' } elseif ($r.Risk -eq 'medium') { 'risk-medium' } else { '' }
        $upn = [System.Net.WebUtility]::HtmlEncode([string]$r.UserPrincipalName)
        $app = [System.Net.WebUtility]::HtmlEncode([string]$r.App)
        $cli = [System.Net.WebUtility]::HtmlEncode([string]$r.Client)
        $ip  = [System.Net.WebUtility]::HtmlEncode([string]$r.IpAddress)
        $loc = [System.Net.WebUtility]::HtmlEncode([string]$r.Location)
        $err = [System.Net.WebUtility]::HtmlEncode([string]$r.ErrorMessage)
        [void]$sb.AppendLine("<tr class='$riskClass'><td>$t</td><td>$upn</td><td>$app</td><td>$cli</td><td>$ip</td><td>$loc</td><td>$($r.Risk)</td><td>$($r.CAStatus)</td><td class='$outClass'>$($r.Outcome)</td><td class=fail>$err</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table></body></html>')
    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
}

function Read-SignInFilterFromOperator {
    $f = @{}
    $u = Read-UserInput "UPN (blank = all users)"
    if ($u) { $f.UPN = $u.Trim() }
    $r = Read-UserInput "Date range (e.g. '7d', '24h', '2026-05-01 / 2026-05-12'; blank = last 7 days)"
    $from = (Get-Date).AddDays(-7); $to = Get-Date
    if ($r) {
        if ($r -match '^(?<n>\d+)(?<u>[hd])$') {
            $n = [int]$Matches['n']
            $from = if ($Matches['u'] -eq 'h') { (Get-Date).AddHours(-$n) } else { (Get-Date).AddDays(-$n) }
            $to = Get-Date
        } elseif ($r -match '^(?<a>\S+)\s*[/-]\s*(?<b>\S+)$') {
            $a=$null;$b=$null
            [DateTime]::TryParse($Matches['a'],[ref]$a)|Out-Null
            [DateTime]::TryParse($Matches['b'],[ref]$b)|Out-Null
            if ($a) { $from = $a }; if ($b) { $to = $b }
        } else { Write-Warn "Unrecognized date range -- using last 7 days." }
    }
    $f.From = $from; $f.To = $to
    $a = Read-UserInput "App display name (blank = any)";    if ($a) { $f.AppName  = $a.Trim() }
    $ip = Read-UserInput "IP address (blank = any)";          if ($ip) { $f.IP      = $ip.Trim() }
    $c = Read-UserInput "Country (2-letter ISO, blank = any)"; if ($c) { $f.Country = $c.Trim().ToUpper() }
    $rl = Read-UserInput "Risk level (low|medium|high|none; blank = any)"
    if ($rl) { $f.RiskLevel = $rl.Trim().ToLower() }
    $ca = Read-UserInput "CA status (success|failure|notApplied; blank = any)"
    if ($ca) { $f.CAStatus = $ca.Trim() }
    $only = Read-UserInput "Only failures? (y/N)"
    if ($only -match '^[Yy]') { $f.OnlyFailures = $true }
    return $f
}

function Show-UserRecentActivity {
    Write-SectionHeader "Recent Sign-In Activity"
    if (-not (Connect-ForTask 'Report')) { return }  # Report task connects Graph
    $upn = Read-UserInput "User UPN"
    if ([string]::IsNullOrWhiteSpace($upn)) { return }
    Write-InfoMsg "Querying last 30 days of sign-ins for $upn..."
    $records = Search-SignIns -UPN $upn.Trim() -From ((Get-Date).AddDays(-30)) -To (Get-Date) -MaxResults 1000
    Show-SignInTable -Records $records
    if ($records.Count -gt 0 -and (Confirm-Action "Export to CSV + HTML?")) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $base = Join-Path (Get-AuditLogDirectory) "signins-$($upn.Split('@')[0])-$stamp"
        Export-SignInsCsv  -Records $records -Path "$base.csv"
        Export-SignInsHtml -Records $records -Path "$base.html" -Filter @{ UPN=$upn; Range='30d' }
        Write-Success "CSV : $base.csv"
        Write-Success "HTML: $base.html"
    }
    Pause-ForUser
}

function Start-SignInSearch {
    Write-SectionHeader "Sign-In Search"
    if (-not (Connect-ForTask 'Report')) { return }
    $filter = Read-SignInFilterFromOperator
    $params = @{}
    foreach ($k in $filter.Keys) { $params[$k] = $filter[$k] }
    Write-InfoMsg "Querying Graph..."
    $records = Search-SignIns @params
    Show-SignInTable -Records $records
    if ($records.Count -gt 0 -and (Confirm-Action "Export to CSV + HTML?")) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $base = Join-Path (Get-AuditLogDirectory) "signins-$stamp"
        Export-SignInsCsv  -Records $records -Path "$base.csv"
        Export-SignInsHtml -Records $records -Path "$base.html" -Filter $filter
        Write-Success "CSV : $base.csv"
        Write-Success "HTML: $base.html"
    }
    Pause-ForUser
}
