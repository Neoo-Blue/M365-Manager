# ============================================================
#  UnifiedAuditLog.ps1 — Search-UnifiedAuditLog wrapper
#
#  Different surface than SignInLookup.ps1:
#    - Sign-ins are on Microsoft Graph (Invoke-MgGraphRequest)
#    - Unified audit log is on Exchange Online / Purview
#      (Search-UnifiedAuditLog cmdlet, requires Connect-EXO)
#
#  Health check at entry: Get-AdminAuditLogConfig must report
#  UnifiedAuditLogIngestionEnabled = $true. If not, this module
#  prints the exact enable command and returns without searching.
#
#  Retention varies by license (90 / 180 / 365 days). We surface
#  a best-effort hint based on the tenant's subscribed SKUs so
#  the operator knows the floor of the lookup window.
# ============================================================

$script:UALOperationCatalog = @(
    @{ Group='Mail';        Ops=@('Send','MailboxLogin','AddFolderPermissions','RemoveFolderPermissions','UpdateInboxRules','Set-Mailbox','New-InboxRule','Set-MailboxAutoReplyConfiguration') },
    @{ Group='File access'; Ops=@('FileAccessed','FileDownloaded','FilePreviewed','FileUploaded','FileDeleted','FileShared','FileSyncDownloadedFull','FileSyncUploadedFull') },
    @{ Group='Identity';    Ops=@('Add user.','Delete user.','Reset user password.','Change user password.','Add member to role.','Remove member from role.') },
    @{ Group='Groups';      Ops=@('Add member to group.','Remove member from group.','Add group.','Delete group.','Update group.') },
    @{ Group='Compliance';  Ops=@('New-eDiscoveryCase','Set-eDiscoveryCase','New-ComplianceSearch','Start-ComplianceSearch','New-ComplianceSearchAction','New-CaseHoldPolicy') },
    @{ Group='Sharing';     Ops=@('AnonymousLinkCreated','SecureLinkCreated','SharingPolicyChanged','CompanyLinkCreated') }
)

function Assert-UnifiedAuditLogReady {
    <#
        Verify EXO is reachable AND unified audit log ingestion is
        actually turned on. Returns $true only if both pass.
    #>
    if (-not (Get-Command Get-AdminAuditLogConfig -ErrorAction SilentlyContinue)) {
        Write-Warn "ExchangeOnlineManagement cmdlets not loaded. Connect EXO from the main menu first."
        return $false
    }
    try {
        $cfg = Get-AdminAuditLogConfig -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Get-AdminAuditLogConfig failed: $($_.Exception.Message)"
        return $false
    }
    if (-not $cfg.UnifiedAuditLogIngestionEnabled) {
        Write-Host ""
        Write-ErrorMsg "Unified audit log ingestion is DISABLED in this tenant."
        Write-InfoMsg "Enable it (one-time, requires Global Admin or Compliance Admin) with:"
        Write-Host  "    Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true" -ForegroundColor Yellow
        Write-Warn "Audit ingestion has a 24-hour warm-up; new events may not appear until tomorrow."
        return $false
    }
    Write-Success "UnifiedAuditLogIngestionEnabled = True."
    return $true
}

function Get-UnifiedAuditRetentionHint {
    <#
        Best-effort: derive likely retention from tenant SKUs.
        Microsoft's docs:
            E5 / G5         -> 365 days default, configurable up to 10y
            E3 / G3 / Basic -> 180 days
            Standalone Audit add-on / 365 audit logging -> 1 year+
        Returns a human-readable string. Falls back to the generic
        "90/180/365 depending on license" when we can't tell.
    #>
    try {
        $skus = @(Get-MgSubscribedSku -ErrorAction Stop)
    } catch { return '90 / 180 / 365 days (depends on license; could not enumerate SKUs)' }

    $names = @($skus | ForEach-Object { $_.SkuPartNumber })
    $hasE5 = $names -match 'SPE_E5|ENTERPRISEPREMIUM|M365_E5|G5'
    $hasE3 = $names -match 'SPE_E3|ENTERPRISEPACK|M365_E3|G3'

    if ($hasE5) { return 'likely 365 days (E5 / G5 detected; tenant policy may extend up to 10 years)' }
    if ($hasE3) { return 'likely 180 days (E3 / G3 detected)' }
    return '90 / 180 / 365 days (depends on license)'
}

function Search-UAL {
    <#
        Run Search-UnifiedAuditLog with sensible defaults. Returns
        normalized records (Time, UserId, Operation, RecordType,
        IpAddress, Workload, ObjectId, AuditDataObj).
    #>
    param(
        [DateTime]$From = (Get-Date).AddDays(-7),
        [DateTime]$To   = (Get-Date),
        [string]$UserId,
        [string[]]$Operations,
        [string]$RecordType,
        [string]$IP,
        [int]$ResultSize = 1000
    )

    $params = @{
        StartDate  = $From
        EndDate    = $To
        ResultSize = $ResultSize
        ErrorAction= 'Stop'
    }
    if ($UserId)     { $params.UserIds     = $UserId }
    if ($Operations) { $params.Operations  = $Operations }
    if ($RecordType) { $params.RecordType  = $RecordType }
    if ($IP)         { $params.IPAddresses = $IP }

    $raw = @()
    try { $raw = @(Search-UnifiedAuditLog @params) }
    catch {
        Write-ErrorMsg "Search-UnifiedAuditLog failed: $($_.Exception.Message)"
        return @()
    }

    $out = New-Object System.Collections.ArrayList
    foreach ($r in $raw) {
        $parsed = $null
        if ($r.AuditData) {
            try { $parsed = $r.AuditData | ConvertFrom-Json -ErrorAction Stop } catch {}
        }
        [void]$out.Add([PSCustomObject]@{
            TimeUtc       = if ($r.CreationDate) { ([DateTime]$r.CreationDate).ToUniversalTime() } else { $null }
            UserId        = $r.UserIds
            Operation     = $r.Operations
            RecordType    = $r.RecordType
            ResultStatus  = if ($parsed) { $parsed.ResultStatus } else { '' }
            ObjectId      = if ($parsed -and $parsed.ObjectId) { $parsed.ObjectId } else { '' }
            ClientIP      = if ($parsed -and $parsed.ClientIP) { $parsed.ClientIP } else { '' }
            Workload      = if ($parsed -and $parsed.Workload) { $parsed.Workload } else { '' }
            UserAgent     = if ($parsed -and $parsed.UserAgent) { $parsed.UserAgent } else { '' }
            ApplicationId = if ($parsed -and $parsed.ApplicationId) { $parsed.ApplicationId } else { '' }
            AuditDataJson = $r.AuditData
        })
    }
    return @($out | Sort-Object TimeUtc -Descending)
}

function Show-UALTable {
    param([array]$Records)
    if (-not $Records -or $Records.Count -eq 0) {
        Write-InfoMsg "(no unified audit log entries match)"; return
    }
    Write-Host ""
    Write-Host ("  TIME (UTC)           USER                                  OPERATION                        WORKLOAD     IP") -ForegroundColor DarkGray
    Write-Host ("  " + ('-' * 130)) -ForegroundColor DarkGray
    foreach ($r in $Records) {
        $t = if ($r.TimeUtc) { $r.TimeUtc.ToString('yyyy-MM-dd HH:mm:ss') } else { '???' }
        $u = if ($r.UserId) { ([string]$r.UserId).PadRight(38).Substring(0,38) } else { ''.PadRight(38) }
        $o = if ($r.Operation) { ([string]$r.Operation).PadRight(32).Substring(0,32) } else { ''.PadRight(32) }
        $w = if ($r.Workload) { ([string]$r.Workload).PadRight(12).Substring(0,12) } else { ''.PadRight(12) }
        $ip = $r.ClientIP
        Write-Host ("  {0} {1} {2} {3} {4}" -f $t, $u, $o, $w, $ip) -ForegroundColor White
    }
    Write-Host ""
    Write-InfoMsg "$($Records.Count) row(s)"
}

function Export-UALCsv {
    param([array]$Records, [string]$Path)
    $Records | Select-Object TimeUtc,UserId,Operation,RecordType,ResultStatus,Workload,ObjectId,ClientIP,UserAgent,ApplicationId,AuditDataJson |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Force
}

function Export-UALHtml {
    param([array]$Records, [string]$Path, [hashtable]$Filter)
    $filterParts = @()
    if ($Filter) { foreach ($k in $Filter.Keys) { if ("$($Filter[$k])") { $filterParts += "<b>$k</b>=$($Filter[$k])" } } }
    $filterText = if ($filterParts.Count -gt 0) { $filterParts -join ' &middot; ' } else { '(no filter)' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset=utf-8><title>Unified audit log</title>')
    [void]$sb.AppendLine('<style>body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;margin:24px;background:#f5f7fa;color:#222}')
    [void]$sb.AppendLine('h1{font-size:18px;margin:0 0 4px 0}.meta{color:#666;font-size:13px;margin-bottom:16px}')
    [void]$sb.AppendLine('table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.06)}')
    [void]$sb.AppendLine('th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #e6e8eb;vertical-align:top}')
    [void]$sb.AppendLine('th{background:#eef2f5;font-weight:600;position:sticky;top:0}')
    [void]$sb.AppendLine('td.json{font-family:Menlo,Consolas,monospace;font-size:11px;color:#666;max-width:520px;word-break:break-word}</style></head><body>')
    [void]$sb.AppendLine('<h1>Unified audit log</h1>')
    [void]$sb.AppendLine("<div class=meta>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz') &middot; $($Records.Count) row(s) &middot; filter: $filterText</div>")
    [void]$sb.AppendLine('<table><thead><tr><th>Time (UTC)</th><th>User</th><th>Operation</th><th>Workload</th><th>IP</th><th>ObjectId</th><th>Result</th><th>AuditData</th></tr></thead><tbody>')
    foreach ($r in $Records) {
        $t = if ($r.TimeUtc) { $r.TimeUtc.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        $u = [System.Net.WebUtility]::HtmlEncode([string]$r.UserId)
        $o = [System.Net.WebUtility]::HtmlEncode([string]$r.Operation)
        $w = [System.Net.WebUtility]::HtmlEncode([string]$r.Workload)
        $ip = [System.Net.WebUtility]::HtmlEncode([string]$r.ClientIP)
        $obj = [System.Net.WebUtility]::HtmlEncode([string]$r.ObjectId)
        $res = [System.Net.WebUtility]::HtmlEncode([string]$r.ResultStatus)
        $j = [System.Net.WebUtility]::HtmlEncode([string]$r.AuditDataJson)
        [void]$sb.AppendLine("<tr><td>$t</td><td>$u</td><td>$o</td><td>$w</td><td>$ip</td><td>$obj</td><td>$res</td><td class=json>$j</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table></body></html>')
    Set-Content -LiteralPath $Path -Value $sb.ToString() -Encoding UTF8
}

function Read-UALFilterFromOperator {
    $f = @{}
    $u = Read-UserInput "User UPN (blank = all users)"
    if ($u) { $f.UserId = $u.Trim() }
    $r = Read-UserInput "Date range (e.g. '7d', '24h', 'YYYY-MM-DD / YYYY-MM-DD'; blank = last 7 days)"
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
        }
    }
    $f.From = $from; $f.To = $to

    # Pick an operation group
    $groupLabels = $script:UALOperationCatalog | ForEach-Object { "{0}  ({1})" -f $_.Group, ($_.Ops -join ', ') }
    $sel = Show-Menu -Title "Operation group (or skip)" -Options $groupLabels -BackLabel "Skip / any operation"
    if ($sel -ge 0) {
        $picked = $script:UALOperationCatalog[$sel].Ops
        $opLabels = $picked | ForEach-Object { $_ }
        $picks = Show-MultiSelect -Title "Specific operations (Enter for ALL in group)" -Options $opLabels
        if ($picks -and $picks.Count -gt 0) {
            $f.Operations = @($picks | ForEach-Object { $picked[$_] })
        } else {
            $f.Operations = $picked
        }
    }
    $rt = Read-UserInput "RecordType (blank = any; e.g. ExchangeItem, SharePointFileOperation, AzureActiveDirectoryStsLogon)"
    if ($rt) { $f.RecordType = $rt.Trim() }
    $ip = Read-UserInput "Client IP (blank = any)"
    if ($ip) { $f.IP = $ip.Trim() }
    return $f
}

function Start-UnifiedAuditSearch {
    Write-SectionHeader "Unified Audit Log"
    if (-not (Connect-ForTask 'Report')) { return }
    if (-not (Assert-UnifiedAuditLogReady)) { Pause-ForUser; return }

    Write-InfoMsg ("Retention (best-effort estimate): " + (Get-UnifiedAuditRetentionHint))
    Write-Host ""

    $filter = Read-UALFilterFromOperator
    $params = @{}
    foreach ($k in $filter.Keys) { $params[$k] = $filter[$k] }
    Write-InfoMsg "Querying unified audit log..."
    $records = Search-UAL @params
    Show-UALTable -Records $records

    if ($records.Count -gt 0 -and (Confirm-Action "Export to CSV + HTML?")) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $base = Join-Path (Get-AuditLogDirectory) "ual-$stamp"
        Export-UALCsv  -Records $records -Path "$base.csv"
        Export-UALHtml -Records $records -Path "$base.html" -Filter $filter
        Write-Success "CSV : $base.csv"
        Write-Success "HTML: $base.html"
    }
    Pause-ForUser
}
