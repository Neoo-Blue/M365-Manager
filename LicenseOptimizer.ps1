# ============================================================
#  LicenseOptimizer.ps1 -- license usage + savings dashboard
#
#  Required Graph scope: Reports.Read.All (added to MgScopes in
#  Phase 4 Commit A). Tenant reports come back de-identified by
#  default; the operator must turn off concealment in
#  Microsoft 365 admin center -> Settings -> Org settings ->
#  Reports -> "Display concealed user, group, and site names in
#  all reports". We detect the concealed-hash UPN shape on first
#  fetch and surface the exact fix.
#
#  Cost estimates use templates/license-prices.json (placeholder
#  monthly USD/user list prices). Customers override the file in-
#  place; the dashboard labels the savings column clearly.
# ============================================================

$script:LicenseRecommendations = @()   # populated by Get-LicenseUtilizationReport

$script:LicenseFamilies = @{
    'M365_E_FAMILY'        = @('SPE_E3','SPE_E5','ENTERPRISEPACK','ENTERPRISEPREMIUM','STANDARDPACK','OFFICESUBSCRIPTION')
    'POWER_BI'             = @('POWER_BI_PRO','POWER_BI_PREMIUM_PER_USER')
    'MOBILITY_AND_SECURITY'= @('EMS','EMSPREMIUM')
    'PROJECT'              = @('PROJECT_P1','PROJECT_P3')
    'FRONTLINE'            = @('M365_F1','M365_F3')
}

function Get-LicensePrices {
    $base = $null
    if ($PSScriptRoot) { $base = $PSScriptRoot }
    elseif ($env:M365ADMIN_ROOT) { $base = $env:M365ADMIN_ROOT }
    else { $base = (Get-Location).Path }
    $path = Join-Path $base 'templates/license-prices.json'
    if (-not (Test-Path -LiteralPath $path)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) {
            if ($prop.Name -like '_comment*') { continue }
            $val = [double]0
            if ([double]::TryParse([string]$prop.Value, [ref]$val)) { $h[$prop.Name] = $val }
        }
        return $h
    } catch { return @{} }
}

function Get-LicensePrice {
    param([Parameter(Mandatory)][string]$SkuPartNumber)
    $prices = Get-LicensePrices
    if ($prices.ContainsKey($SkuPartNumber)) { return [double]$prices[$SkuPartNumber] }
    return $null
}

function Get-GraphReportCsv {
    <#
        Fetch a Graph report endpoint (returns CSV) and parse into
        PSCustomObjects. Detects de-identified UPNs and surfaces
        the exact admin-center fix when triggered.
    #>
    param(
        [Parameter(Mandatory)][string]$ReportName,   # e.g. getOffice365ActiveUserDetail
        [string]$Period = 'D30'
    )
    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        Write-Warn "Microsoft.Graph.Authentication not loaded."; return @()
    }
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $uri = "https://graph.microsoft.com/v1.0/reports/${ReportName}(period='$Period')"
        try { Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $tmp -ErrorAction Stop }
        catch {
            $msg = $_.Exception.Message
            Write-ErrorMsg "Report fetch failed ($ReportName): $msg"
            if ($msg -match '403|Forbidden|Authorization') {
                Write-InfoMsg "Required scope: Reports.Read.All. Reconnect after consent."
            }
            return @()
        }
        $rows = @()
        try { $rows = @(Import-Csv -LiteralPath $tmp) } catch { Write-Warn "CSV parse failed: $_"; return @() }

        # De-identification detection: first non-empty UPN column value
        $upnCol = $null
        foreach ($c in 'User Principal Name','UserPrincipalName','User principal name') {
            if ($rows.Count -gt 0 -and ($rows[0].PSObject.Properties.Name -contains $c)) { $upnCol = $c; break }
        }
        if ($upnCol) {
            $firstUpn = ($rows | Where-Object { $_.$upnCol } | Select-Object -First 1).$upnCol
            if ($firstUpn -and $firstUpn -match '^[A-F0-9]{50,}$') {
                Write-Warn "Tenant Graph reports are returning DE-IDENTIFIED user names (the UPN column looks like a hash)."
                Write-InfoMsg "To enable real UPNs: Microsoft 365 admin center -> Settings -> Org settings -> Services tab -> Reports -> tick 'Display concealed user, group, and site names in all reports' -> Save. Re-run after the change propagates (5-10 minutes)."
                return @()
            }
        }
        return $rows
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } }
}

function Get-ActivityIndex {
    <#
        Build a UPN -> last-activity hashtable by unioning the four
        activity reports. Each row contributes its 'Last Activity
        Date' (or equivalent column). Returns @{ upn -> DateTime }.
    #>
    param([string]$Period = 'D30')
    $index = @{}
    foreach ($r in (Get-GraphReportCsv -ReportName 'getOffice365ActiveUserDetail' -Period $Period)) {
        $upn = $r.'User Principal Name'; if (-not $upn) { continue }
        $d = $null
        if ($r.'Last Activity Date') { [DateTime]::TryParse($r.'Last Activity Date', [ref]$d) | Out-Null }
        if ($d -and (-not $index.ContainsKey($upn) -or $index[$upn] -lt $d)) { $index[$upn] = $d }
    }
    foreach ($report in 'getMailboxUsageDetail','getTeamsUserActivityUserDetail','getOneDriveUsageAccountDetail') {
        foreach ($r in (Get-GraphReportCsv -ReportName $report -Period $Period)) {
            $upn = $r.'User Principal Name'; if (-not $upn) { continue }
            $col = $null
            foreach ($c in 'Last Activity Date','Last Active Date','Last Activity Date (UTC)') {
                if ($r.PSObject.Properties.Name -contains $c) { $col = $c; break }
            }
            if (-not $col) { continue }
            $d = $null; [DateTime]::TryParse([string]$r.$col, [ref]$d) | Out-Null
            if ($d -and (-not $index.ContainsKey($upn) -or $index[$upn] -lt $d)) { $index[$upn] = $d }
        }
    }
    return $index
}

function Get-LicensedUserMap {
    <#
        UPN -> @(SkuPartNumber,...) from Get-MgUser
        +Get-MgUserLicenseDetail. Filter on accountEnabled to
        avoid flagging soft-deleted users.
    #>
    $map = @{}
    try {
        $users = @(Get-MgUser -All -Property 'Id,UserPrincipalName,AccountEnabled,AssignedLicenses' -ErrorAction Stop)
    } catch { Write-ErrorMsg "Could not enumerate users: $_"; return @{} }
    $skuLookup = @{}
    try {
        foreach ($s in (Get-MgSubscribedSku -ErrorAction Stop)) { $skuLookup[[string]$s.SkuId] = [string]$s.SkuPartNumber }
    } catch { Write-Warn "Could not enumerate SKUs: $_" }
    foreach ($u in $users) {
        if (-not $u.AccountEnabled) { continue }
        if (-not $u.AssignedLicenses -or $u.AssignedLicenses.Count -eq 0) { continue }
        $parts = @()
        foreach ($lic in $u.AssignedLicenses) {
            $sid = [string]$lic.SkuId
            if ($skuLookup.ContainsKey($sid)) { $parts += $skuLookup[$sid] }
        }
        if ($parts.Count -gt 0) { $map[[string]$u.UserPrincipalName] = $parts }
    }
    return $map
}

# ============================================================
#  Findings
# ============================================================

function Get-InactiveLicensedUsers {
    param([int]$DaysInactive = 60)
    if (-not (Connect-ForTask 'Report')) { return @() }
    $cutoff = (Get-Date).AddDays(-$DaysInactive)
    $period = if ($DaysInactive -le 7) { 'D7' } elseif ($DaysInactive -le 30) { 'D30' } elseif ($DaysInactive -le 90) { 'D90' } else { 'D180' }
    $licensed = Get-LicensedUserMap
    $activity = Get-ActivityIndex -Period $period
    $hits = New-Object System.Collections.ArrayList
    foreach ($upn in $licensed.Keys) {
        $last = $activity[$upn]
        $isInactive = (-not $last) -or ($last -lt $cutoff)
        if ($isInactive) {
            $skus = $licensed[$upn]
            $cost = 0.0
            foreach ($s in $skus) { $p = Get-LicensePrice $s; if ($p) { $cost += $p } }
            [void]$hits.Add([PSCustomObject]@{
                UPN              = $upn
                LastActivityUtc  = $last
                DaysSinceActive  = if ($last) { [int]((Get-Date) - $last).TotalDays } else { 9999 }
                Skus             = ($skus -join ', ')
                EstMonthlyCost   = [Math]::Round($cost, 2)
            })
        }
    }
    return @($hits | Sort-Object DaysSinceActive -Descending)
}

function Get-UnusedLicenses {
    param([int]$Period = 30)
    return @(Get-InactiveLicensedUsers -DaysInactive $Period | Where-Object { $_.DaysSinceActive -ge $Period })
}

function Get-LicenseHoarders {
    <#
        Users on 2+ SKUs from the same family (e.g. SPE_E3 AND
        ENTERPRISEPACK is redundant: SPE_E3 already includes O365 E3).
    #>
    if (-not (Connect-ForTask 'Report')) { return @() }
    $licensed = Get-LicensedUserMap
    $hits = New-Object System.Collections.ArrayList
    foreach ($upn in $licensed.Keys) {
        $skus = @($licensed[$upn])
        $redundant = @()
        foreach ($family in $script:LicenseFamilies.GetEnumerator()) {
            $inFamily = @($skus | Where-Object { $family.Value -contains $_ })
            if ($inFamily.Count -ge 2) {
                $redundant += ("{0}:[{1}]" -f $family.Key, ($inFamily -join '+'))
            }
        }
        if ($redundant.Count -gt 0) {
            $cost = 0.0
            foreach ($s in $skus) { $p = Get-LicensePrice $s; if ($p) { $cost += $p } }
            [void]$hits.Add([PSCustomObject]@{
                UPN            = $upn
                Skus           = ($skus -join ', ')
                Redundant      = ($redundant -join ' ; ')
                EstMonthlyCost = [Math]::Round($cost, 2)
            })
        }
    }
    return @($hits | Sort-Object EstMonthlyCost -Descending)
}

function Get-PaidUnassignedLicenses {
    if (-not (Connect-ForTask 'Report')) { return @() }
    try {
        $skus = @(Get-MgSubscribedSku -ErrorAction Stop)
    } catch { Write-ErrorMsg "Could not enumerate SKUs: $_"; return @() }
    $hits = New-Object System.Collections.ArrayList
    foreach ($s in $skus) {
        $total = [int]$s.PrepaidUnits.Enabled
        $used  = [int]$s.ConsumedUnits
        $free  = $total - $used
        if ($free -le 0) { continue }
        $price = Get-LicensePrice $s.SkuPartNumber
        $monthlyWaste = if ($price) { $free * $price } else { $null }
        [void]$hits.Add([PSCustomObject]@{
            SkuPartNumber       = $s.SkuPartNumber
            Total               = $total
            Used                = $used
            Free                = $free
            EstMonthlyWasteUsd  = if ($monthlyWaste) { [Math]::Round($monthlyWaste, 2) } else { '' }
        })
    }
    return @($hits | Sort-Object Free -Descending)
}

function Get-DowngradeCandidates {
    <#
        Conservative: a user on a premium SKU (E5 family) whose
        last activity index entry is older than 60 days. We don't
        try to detect per-feature usage at this scope; that would
        require service-specific reports (Defender ATP, Power BI,
        etc.). Operators using this for serious cost cutting should
        cross-check the per-service usage reports manually before
        downgrading.
    #>
    param([int]$DaysSinceUsed = 60)
    if (-not (Connect-ForTask 'Report')) { return @() }
    $premium = @('SPE_E5','ENTERPRISEPREMIUM','SPB')
    $licensed = Get-LicensedUserMap
    $activity = Get-ActivityIndex -Period 'D90'
    $hits = New-Object System.Collections.ArrayList
    $cutoff = (Get-Date).AddDays(-$DaysSinceUsed)
    foreach ($upn in $licensed.Keys) {
        $skus = $licensed[$upn]
        $pSkus = @($skus | Where-Object { $premium -contains $_ })
        if ($pSkus.Count -eq 0) { continue }
        $last = $activity[$upn]
        if ($last -and $last -ge $cutoff) { continue }
        $costPremium = 0.0; $costE3 = 0.0
        foreach ($s in $pSkus) {
            $p = Get-LicensePrice $s; if ($p) { $costPremium += $p }
        }
        $costE3 = (Get-LicensePrice 'SPE_E3'); if (-not $costE3) { $costE3 = 36.0 }
        $savings = [Math]::Round(($costPremium - ($pSkus.Count * $costE3)), 2)
        [void]$hits.Add([PSCustomObject]@{
            UPN                       = $upn
            PremiumSkus               = ($pSkus -join ', ')
            LastActivityUtc           = $last
            DaysSinceActive           = if ($last) { [int]((Get-Date) - $last).TotalDays } else { 9999 }
            EstMonthlySavingsIfDowngraded = $savings
        })
    }
    return @($hits | Sort-Object EstMonthlySavingsIfDowngraded -Descending)
}

# ============================================================
#  Dashboard
# ============================================================

function Format-LicenseDashboardHtml {
    param(
        [array]$Inactive,
        [array]$Hoarders,
        [array]$Unassigned,
        [array]$Downgrade,
        [hashtable]$Prices
    )
    $inactiveSavings = ($Inactive    | Measure-Object -Property EstMonthlyCost -Sum).Sum
    $hoarderSavings  = ($Hoarders    | Measure-Object -Property EstMonthlyCost -Sum).Sum
    $unassignedWaste = ($Unassigned  | Where-Object EstMonthlyWasteUsd -is [double] | Measure-Object -Property EstMonthlyWasteUsd -Sum).Sum
    $downgradeSavings= ($Downgrade   | Measure-Object -Property EstMonthlySavingsIfDowngraded -Sum).Sum
    $totalSavings    = [Math]::Round((($inactiveSavings + $hoarderSavings + $unassignedWaste + $downgradeSavings) | ForEach-Object { if ($_) { $_ } else { 0 } } | Measure-Object -Sum).Sum, 2)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset=utf-8><title>License utilization</title>')
    [void]$sb.AppendLine('<style>body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;margin:24px;background:#f5f7fa;color:#222}')
    [void]$sb.AppendLine('h1{margin:0 0 4px 0}.meta{color:#666;font-size:13px;margin-bottom:20px}')
    [void]$sb.AppendLine('.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-bottom:20px}')
    [void]$sb.AppendLine('.tile{background:#fff;border:1px solid #e6e8eb;border-radius:6px;padding:14px}.tile h3{margin:0;font-size:13px;color:#666;font-weight:600;text-transform:uppercase;letter-spacing:.04em}.tile .v{font-size:26px;font-weight:600;margin-top:4px}.tile .note{font-size:12px;color:#888;margin-top:4px}')
    [void]$sb.AppendLine('table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.06);margin-bottom:18px}th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #e6e8eb;vertical-align:top}th{background:#eef2f5;font-weight:600}h2{font-size:15px;margin:22px 0 8px}</style></head><body>')
    [void]$sb.AppendLine('<h1>License utilization &amp; cost dashboard</h1>')
    [void]$sb.AppendLine("<div class=meta>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'). Estimated savings use templates/license-prices.json -- override that file with real EA pricing for accurate figures.</div>")
    [void]$sb.AppendLine('<div class=tiles>')
    [void]$sb.AppendLine("<div class=tile><h3>Total est. monthly savings opportunity</h3><div class=v>`$$totalSavings</div><div class=note>list-price placeholders</div></div>")
    [void]$sb.AppendLine("<div class=tile><h3>Inactive licensed users</h3><div class=v>$($Inactive.Count)</div><div class=note>no activity in window</div></div>")
    [void]$sb.AppendLine("<div class=tile><h3>License hoarders</h3><div class=v>$($Hoarders.Count)</div><div class=note>overlapping SKUs</div></div>")
    [void]$sb.AppendLine("<div class=tile><h3>Paid &amp; unassigned</h3><div class=v>$(($Unassigned | Measure-Object Free -Sum).Sum)</div><div class=note>seats across all SKUs</div></div>")
    [void]$sb.AppendLine("<div class=tile><h3>Downgrade candidates</h3><div class=v>$($Downgrade.Count)</div><div class=note>premium SKU + no use</div></div>")
    [void]$sb.AppendLine('</div>')

    function _renderTable {
        param([string]$title, [array]$rows, [string[]]$cols)
        $h = "<h2>$title</h2><table><thead><tr>"
        foreach ($c in $cols) { $h += "<th>$c</th>" }
        $h += "</tr></thead><tbody>"
        foreach ($r in ($rows | Select-Object -First 25)) {
            $h += "<tr>"
            foreach ($c in $cols) { $v = [System.Net.WebUtility]::HtmlEncode([string]$r.$c); $h += "<td>$v</td>" }
            $h += "</tr>"
        }
        $h += "</tbody></table>"
        return $h
    }
    [void]$sb.AppendLine((_renderTable "Inactive licensed users (top 25)" $Inactive @('UPN','DaysSinceActive','Skus','EstMonthlyCost')))
    [void]$sb.AppendLine((_renderTable "License hoarders (top 25 by cost)" $Hoarders  @('UPN','Skus','Redundant','EstMonthlyCost')))
    [void]$sb.AppendLine((_renderTable "Paid + unassigned seats"           $Unassigned @('SkuPartNumber','Total','Used','Free','EstMonthlyWasteUsd')))
    [void]$sb.AppendLine((_renderTable "Downgrade candidates (top 25)"     $Downgrade @('UPN','PremiumSkus','DaysSinceActive','EstMonthlySavingsIfDowngraded')))
    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

function Get-LicenseUtilizationReport {
    [CmdletBinding()]
    param([int]$DaysInactive = 60)
    if (-not (Connect-ForTask 'Report')) { return $null }
    Write-InfoMsg "Fetching activity reports (D30 + D90)..."
    $inactive   = Get-InactiveLicensedUsers -DaysInactive $DaysInactive
    $hoarders   = Get-LicenseHoarders
    $unassigned = Get-PaidUnassignedLicenses
    $downgrade  = Get-DowngradeCandidates -DaysSinceUsed $DaysInactive
    $prices     = Get-LicensePrices

    $script:LicenseRecommendations = @(
        $inactive  | ForEach-Object { [PSCustomObject]@{ Kind='Inactive';  UPN=$_.UPN; Sku=($_.Skus -split ',')[0].Trim(); Action='Remove'; Note="DaysSinceActive=$($_.DaysSinceActive)"; EstMonthlySavings=$_.EstMonthlyCost } }
        $hoarders  | ForEach-Object { [PSCustomObject]@{ Kind='Hoarder';   UPN=$_.UPN; Sku=($_.Skus -split ',')[0].Trim(); Action='Remove'; Note=$_.Redundant; EstMonthlySavings=$_.EstMonthlyCost } }
        $downgrade | ForEach-Object { [PSCustomObject]@{ Kind='Downgrade'; UPN=$_.UPN; Sku=($_.PremiumSkus -split ',')[0].Trim(); Action='Downgrade'; Note="DaysSinceActive=$($_.DaysSinceActive)"; EstMonthlySavings=$_.EstMonthlySavingsIfDowngraded } }
    )
    $html = Format-LicenseDashboardHtml -Inactive $inactive -Hoarders $hoarders -Unassigned $unassigned -Downgrade $downgrade -Prices $prices
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $out = Join-Path (Get-AuditLogDirectory) "license-utilization-$stamp.html"
    Set-Content -LiteralPath $out -Value $html -Encoding UTF8
    Write-Success "Dashboard: $out"
    return @{
        DashboardPath   = $out
        Inactive        = $inactive
        Hoarders        = $hoarders
        Unassigned      = $unassigned
        Downgrade       = $downgrade
        Recommendations = $script:LicenseRecommendations
    }
}

# ============================================================
#  Remediation
# ============================================================

function Invoke-LicenseRemediation {
    <#
        CSV columns: UPN, SKU, Action (Remove|Downgrade), TargetSKU (only for Downgrade)
        Validate-then-execute. Each Remove routes through Invoke-
        Action with ActionType=RemoveLicense + ReverseType=AssignLicense
        (same recipe Phase 2 set up). Downgrade is a paired
        Remove+Assign sequence; reverse recipe pairs them via the
        usual handler dispatch.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$WhatIf)
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { Write-Warn "Empty CSV."; return }
    if (-not (Connect-ForTask 'License')) { return }

    $skuMap = @{}
    try { foreach ($s in (Get-MgSubscribedSku -ErrorAction Stop)) { $skuMap[$s.SkuPartNumber] = [string]$s.SkuId } } catch {}

    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    try {
        $results = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $r = $rows[$i]
            $upn = [string]$r.UPN; $sku = [string]$r.SKU; $act = [string]$r.Action; $target = [string]$r.TargetSKU
            Write-Progress -Activity "License remediation" -Status "$upn $act $sku" -PercentComplete (($i / $rows.Count) * 100)
            $entry = [ordered]@{ UPN = $upn; SKU = $sku; Action = $act; TargetSKU = $target; Status = ''; Reason = '' }
            if (-not $skuMap.ContainsKey($sku)) { $entry.Status='Failed'; $entry.Reason="Unknown SKU '$sku'"; [void]$results.Add([PSCustomObject]$entry); continue }
            try {
                $skuId = $skuMap[$sku]
                $userId = (Get-MgUser -UserId $upn -ErrorAction Stop).Id
                switch ($act) {
                    'Remove' {
                        $ok = Invoke-Action `
                            -Description ("Remove license '{0}' from {1}" -f $sku, $upn) `
                            -ActionType 'RemoveLicense' `
                            -Target @{ userId = $userId; userUpn = $upn; skuId = $skuId; skuPart = $sku } `
                            -ReverseType 'AssignLicense' `
                            -ReverseDescription ("Re-assign '{0}' to {1}" -f $sku, $upn) `
                            -Action { Set-MgUserLicense -UserId $userId -AddLicenses @() -RemoveLicenses @($skuId) -ErrorAction Stop; $true }
                        $entry.Status = if ($ok) { if (Get-PreviewMode) {'Preview'} else {'Success'} } else { 'Failed' }
                    }
                    'Downgrade' {
                        if (-not $target -or -not $skuMap.ContainsKey($target)) { $entry.Status='Failed'; $entry.Reason="Unknown TargetSKU '$target'"; break }
                        $targetId = $skuMap[$target]
                        Invoke-Action `
                            -Description ("Assign target license '{0}' to {1}" -f $target, $upn) `
                            -ActionType 'AssignLicense' `
                            -Target @{ userId = $userId; userUpn = $upn; skuId = $targetId; skuPart = $target } `
                            -ReverseType 'RemoveLicense' `
                            -ReverseDescription ("Remove '{0}' from {1}" -f $target, $upn) `
                            -Action { Set-MgUserLicense -UserId $userId -AddLicenses @(@{SkuId=$targetId}) -RemoveLicenses @() -ErrorAction Stop; $true } | Out-Null
                        Invoke-Action `
                            -Description ("Remove premium license '{0}' from {1}" -f $sku, $upn) `
                            -ActionType 'RemoveLicense' `
                            -Target @{ userId = $userId; userUpn = $upn; skuId = $skuId; skuPart = $sku } `
                            -ReverseType 'AssignLicense' `
                            -ReverseDescription ("Re-assign '{0}' to {1}" -f $sku, $upn) `
                            -Action { Set-MgUserLicense -UserId $userId -AddLicenses @() -RemoveLicenses @($skuId) -ErrorAction Stop; $true } | Out-Null
                        $entry.Status = if (Get-PreviewMode) {'Preview'} else {'Success'}
                    }
                    default { $entry.Status='Failed'; $entry.Reason="Unknown Action '$act'" }
                }
            } catch { $entry.Status='Failed'; $entry.Reason=$_.Exception.Message }
            [void]$results.Add([PSCustomObject]$entry)
        }
        Write-Progress -Activity "License remediation" -Completed
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("license-remediation-$stamp.csv")
        $results | Export-Csv -LiteralPath $out -NoTypeInformation -Force
        Write-Success "Result CSV: $out"
    } finally { Set-PreviewMode -Enabled $previousMode }
}

# ============================================================
#  Menu
# ============================================================

function Start-LicenseOptimizerMenu {
    while ($true) {
        $sel = Show-Menu -Title "License & Cost" -Options @(
            "Run full utilization dashboard (HTML)",
            "Inactive licensed users (>= 60 days)...",
            "License hoarders (overlapping SKUs)",
            "Paid + unassigned seats",
            "Downgrade candidates (premium SKU + no use)",
            "Remediate from CSV..."
        ) -BackLabel "Back"
        switch ($sel) {
            0 { $r = Get-LicenseUtilizationReport -DaysInactive 60; if ($r) { Write-StatusLine "Inactive"   "$($r.Inactive.Count)" 'White'; Write-StatusLine "Hoarders" "$($r.Hoarders.Count)" 'White'; Write-StatusLine "Free seats" "$(($r.Unassigned | Measure-Object Free -Sum).Sum)" 'White'; Write-StatusLine "Downgrade" "$($r.Downgrade.Count)" 'White' }; Pause-ForUser }
            1 { $dt = Read-UserInput "Days threshold (default 60)"; $d = 60; [int]::TryParse($dt,[ref]$d) | Out-Null; Get-InactiveLicensedUsers -DaysInactive $d | Format-Table -AutoSize; Pause-ForUser }
            2 { Get-LicenseHoarders         | Format-Table -AutoSize; Pause-ForUser }
            3 { Get-PaidUnassignedLicenses  | Format-Table -AutoSize; Pause-ForUser }
            4 { Get-DowngradeCandidates     | Format-Table -AutoSize; Pause-ForUser }
            5 {
                $p = Read-UserInput "Path to remediation CSV (UPN, SKU, Action, TargetSKU?)"
                if (-not $p) { continue }
                $dry = Confirm-Action "Run as DRY-RUN first?"
                Invoke-LicenseRemediation -Path $p.Trim('"').Trim("'") -WhatIf:$dry
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
