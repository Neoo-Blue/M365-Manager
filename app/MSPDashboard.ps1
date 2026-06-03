# ============================================================
#  MSPDashboard.ps1 -- single-page HTML portfolio dashboard
#  (Phase 6 commit D)
#
#  Aggregates the cross-tenant reports from MSPReports.ps1 into
#  one self-contained HTML file. No external assets (everything
#  inline) so the file is safe to email or open from a USB stick.
#
#  Refresh by calling Update-MSPDashboard. The result lands in
#  <stateDir>/msp-dashboard/msp-dashboard-<ts>.html with a
#  symlink/copy at msp-dashboard-latest.html for easy
#  bookmarking.
# ============================================================

function Get-MSPDashboardDir {
    $base = if (Get-Command Get-StateDirectory -ErrorAction SilentlyContinue) { Get-StateDirectory } else { (Get-Location).Path }
    $d = Join-Path $base 'msp-dashboard'
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    return $d
}

function ConvertTo-HtmlSafe { param([string]$Text); if ($null -eq $Text) { return '' }; return ([System.Web.HttpUtility]::HtmlEncode([string]$Text)) }

function Get-PostureDotColor {
    param([string]$Status)
    switch ($Status) {
        'ok'    { return '#22c55e' }    # green
        'warn'  { return '#eab308' }    # yellow
        'fail'  { return '#ef4444' }    # red
        default { return '#9ca3af' }    # gray (unknown)
    }
}

function Build-MSPDashboardCard {
    <#
        Render one per-tenant card. Pure string assembly so the
        caller can join cards in any order. Pulls all metrics
        from the pre-aggregated $TenantSummary hashtable -- this
        function does NOT call any per-tenant report, it only
        formats what's already been gathered.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$TenantSummary
    )
    $t           = $TenantSummary
    $userCount   = if ($null -ne $t.UserCount)        { [int]$t.UserCount }        else { 0 }
    $monthlyUsd  = if ($null -ne $t.MonthlySpendUsd)  { [double]$t.MonthlySpendUsd } else { 0.0 }
    $mfaPct      = if ($null -ne $t.MfaCompliancePct){ [double]$t.MfaCompliancePct } else { 0.0 }
    $staleGuests = if ($null -ne $t.StaleGuests)      { [int]$t.StaleGuests }       else { 0 }
    $orphanTeams = if ($null -ne $t.OrphanTeams)      { [int]$t.OrphanTeams }       else { 0 }
    $posture     = if ($t.BreakGlassPosture)          { [string]$t.BreakGlassPosture } else { 'unknown' }
    $lastSync    = if ($t.LastSync)                   { [string]$t.LastSync }            else { '(never)' }
    $tenantId    = if ($t.TenantId)                   { [string]$t.TenantId }            else { '' }
    $name        = [string]$t.Name
    $nameSafe    = ConvertTo-HtmlSafe $name
    $tidSafe     = ConvertTo-HtmlSafe $tenantId
    $dotColor    = Get-PostureDotColor -Status $posture
    $mfaColor    = if ($mfaPct -ge 95) { '#22c55e' } elseif ($mfaPct -ge 80) { '#eab308' } else { '#ef4444' }
    $usdFormatted= '${0:N2}' -f $monthlyUsd

    @"
<div class="card">
  <div class="card-head">
    <span class="dot" style="background:$dotColor"></span>
    <h2>$nameSafe</h2>
    <span class="tid">$tidSafe</span>
  </div>
  <div class="metrics">
    <div class="metric"><span class="m-num">$userCount</span><span class="m-lbl">users</span></div>
    <div class="metric"><span class="m-num">$usdFormatted</span><span class="m-lbl">/ month</span></div>
    <div class="metric"><span class="m-num" style="color:$mfaColor">$([math]::Round($mfaPct,1))%</span><span class="m-lbl">MFA compliant</span></div>
    <div class="metric"><span class="m-num">$staleGuests</span><span class="m-lbl">stale guests</span></div>
    <div class="metric"><span class="m-num">$orphanTeams</span><span class="m-lbl">orphan teams</span></div>
    <div class="metric"><span class="m-num" style="color:$dotColor">$posture</span><span class="m-lbl">break-glass</span></div>
  </div>
  <div class="footer">last sync: $lastSync</div>
</div>
"@
}

function Build-MSPDashboardHtml {
    <#
        Build the full HTML page. Caller passes the per-tenant
        summary array (already aggregated by Update-MSPDashboard).
        Returns one big UTF-8 string ready to write to disk.
    #>
    param(
        [Parameter(Mandatory)][array]$TenantSummaries,
        [Parameter(Mandatory)][hashtable]$PortfolioTotals
    )
    $cardsHtml = (@($TenantSummaries | ForEach-Object { Build-MSPDashboardCard -TenantSummary $_ }) -join "`n")
    $ts        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    $tot       = $PortfolioTotals
    $totSpend  = '${0:N2}' -f [double]$tot.TotalMonthlyUsd
    $totUsers  = [int]$tot.TotalUsers
    $weightedMfa = [math]::Round([double]$tot.WeightedMfaPct, 1)
    $tenantCount = $TenantSummaries.Count

    @"
<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><title>MSP portfolio dashboard</title>
<style>
  body { font-family: -apple-system, Segoe UI, Roboto, sans-serif; background:#0f172a; color:#e2e8f0; margin:0; padding:24px; }
  h1 { margin:0 0 12px; font-size:24px; }
  .ts { color:#94a3b8; font-size:13px; margin-bottom:24px; }
  .totals { display:flex; gap:24px; margin-bottom:32px; }
  .totals .tot { background:#1e293b; padding:16px 20px; border-radius:8px; }
  .totals .tot-num { font-size:24px; font-weight:600; }
  .totals .tot-lbl { color:#94a3b8; font-size:13px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fill,minmax(340px,1fr)); gap:16px; }
  .card { background:#1e293b; border-radius:10px; padding:16px; border:1px solid #334155; }
  .card-head { display:flex; align-items:center; gap:8px; margin-bottom:12px; }
  .card-head h2 { margin:0; font-size:18px; flex:1; }
  .card-head .tid { font-family:Menlo,monospace; font-size:11px; color:#64748b; }
  .dot { width:12px; height:12px; border-radius:50%; display:inline-block; }
  .metrics { display:grid; grid-template-columns:1fr 1fr 1fr; gap:8px; }
  .metric { background:#0f172a; padding:8px 10px; border-radius:6px; text-align:center; }
  .m-num { display:block; font-size:18px; font-weight:600; }
  .m-lbl { display:block; font-size:11px; color:#94a3b8; margin-top:2px; }
  .footer { font-size:11px; color:#64748b; margin-top:12px; }
</style></head>
<body>
<h1>MSP portfolio dashboard</h1>
<div class="ts">Generated $ts &middot; $tenantCount tenant(s)</div>
<div class="totals">
  <div class="tot"><div class="tot-num">$totUsers</div><div class="tot-lbl">total users</div></div>
  <div class="tot"><div class="tot-num">$totSpend</div><div class="tot-lbl">monthly spend</div></div>
  <div class="tot"><div class="tot-num">$weightedMfa%</div><div class="tot-lbl">weighted MFA</div></div>
</div>
<div class="cards">
$cardsHtml
</div>
</body></html>
"@
}

function Get-MSPTenantSummary {
    <#
        Build the per-tenant summary hashtable that
        Build-MSPDashboardCard consumes. Called from inside
        Switch-Tenant context for ONE tenant -- caller
        (Update-MSPDashboard) handles the switching.
    #>
    param([Parameter(Mandatory)][hashtable]$Tenant)
    $summary = @{
        Name             = $Tenant.name
        TenantId         = $Tenant.tenantId
        LastSync         = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        UserCount        = 0
        MonthlySpendUsd  = 0.0
        MfaCompliancePct = 0.0
        StaleGuests      = 0
        OrphanTeams      = 0
        BreakGlassPosture= 'unknown'
    }
    # User count
    try {
        if (Get-Command Get-MgUser -ErrorAction SilentlyContinue) {
            $summary.UserCount = @(Get-MgUser -All -Property Id -ErrorAction Stop).Count
        }
    } catch {}
    # License spend
    try {
        if (Get-Command Get-LicenseUtilizationReport -ErrorAction SilentlyContinue) {
            $rep = Get-LicenseUtilizationReport
            $total = 0.0
            foreach ($s in @($rep)) { if ($s.MonthlyCostUsd) { $total += [double]$s.MonthlyCostUsd } }
            $summary.MonthlySpendUsd = $total
        }
    } catch {}
    # MFA compliance
    try {
        if (Get-Command Invoke-MfaComplianceScan -ErrorAction SilentlyContinue) {
            $scan = Invoke-MfaComplianceScan
            $total = @($scan).Count
            $noMfa = @(Get-UsersWithNoMfa -Scan $scan).Count
            if ($total -gt 0) { $summary.MfaCompliancePct = [math]::Round((1 - ($noMfa / $total)) * 100, 2) }
        }
    } catch {}
    # Stale guests + orphan teams
    try { if (Get-Command Get-StaleGuests   -ErrorAction SilentlyContinue) { $summary.StaleGuests = @(Get-StaleGuests).Count } } catch {}
    try { if (Get-Command Get-OrphanedTeams -ErrorAction SilentlyContinue) { $summary.OrphanTeams = @(Get-OrphanedTeams).Count } } catch {}
    # Break-glass posture
    try {
        if (Get-Command Test-BreakGlassPosture -ErrorAction SilentlyContinue) {
            $p = Test-BreakGlassPosture
            if ($p -and $p.Status) { $summary.BreakGlassPosture = [string]$p.Status }
        }
    } catch {}
    return $summary
}

function Update-MSPDashboard {
    <#
        Refresh the HTML dashboard. Iterates Invoke-AcrossTenants
        across every (or named) tenant, builds a per-tenant
        summary, renders the HTML, writes it to
        <stateDir>/msp-dashboard/msp-dashboard-<ts>.html plus a
        msp-dashboard-latest.html copy.

        Returns the path of the latest file.
    #>
    [CmdletBinding()]
    param([string[]]$Tenants)

    if (-not (Get-Command Invoke-AcrossTenants -ErrorAction SilentlyContinue)) {
        Write-ErrorMsg "MSPReports.ps1 not loaded; cannot refresh dashboard."
        return $null
    }

    Write-InfoMsg "Refreshing MSP dashboard..."
    $perTenant = Invoke-AcrossTenants -Tenants $Tenants -Script {
        Get-MSPTenantSummary -Tenant $args[0]
    }

    $summaries = @()
    foreach ($r in $perTenant) {
        if ($r.Success -and $r.Result) {
            $summaries += ,$r.Result
        } else {
            $summaries += ,@{
                Name = $r.Tenant; TenantId = $r.TenantId
                LastSync = 'error'; UserCount = 0; MonthlySpendUsd = 0.0
                MfaCompliancePct = 0.0; StaleGuests = 0; OrphanTeams = 0
                BreakGlassPosture = 'fail'
            }
        }
    }
    # Portfolio totals
    $totalUsers = 0; $totalSpend = 0.0; $weightedNumerator = 0.0
    foreach ($s in $summaries) {
        $totalUsers += [int]$s.UserCount
        $totalSpend += [double]$s.MonthlySpendUsd
        $weightedNumerator += ([double]$s.MfaCompliancePct * [int]$s.UserCount)
    }
    $weightedMfa = if ($totalUsers -gt 0) { [math]::Round($weightedNumerator / $totalUsers, 2) } else { 0.0 }
    $totals = @{ TotalUsers = $totalUsers; TotalMonthlyUsd = $totalSpend; WeightedMfaPct = $weightedMfa }

    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    $html = Build-MSPDashboardHtml -TenantSummaries $summaries -PortfolioTotals $totals

    $dir   = Get-MSPDashboardDir
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $path  = Join-Path $dir ("msp-dashboard-{0}.html" -f $stamp)
    Set-Content -LiteralPath $path -Value $html -Encoding UTF8 -Force
    $latest = Join-Path $dir 'msp-dashboard-latest.html'
    Copy-Item -LiteralPath $path -Destination $latest -Force
    Write-Success "Dashboard written: $path"
    Write-InfoMsg ("Latest: {0}" -f $latest)
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'MSPDashboardRefresh' -Detail ("Refreshed across {0} tenant(s)" -f $summaries.Count) -ActionType 'MSPDashboardRefresh' -Target @{ tenantCount = $summaries.Count; totalSpendUsd = $totalSpend } -Result 'ok' | Out-Null
    }
    return $path
}
