# ============================================================
#  IncidentBulk.ps1 -- bulk + tabletop incident response
#
#  Two entry points:
#    Invoke-BulkIncidentResponse  -- "the phishing campaign hit
#                                     12 accounts" scenario.
#                                     CSV-driven, validate-first,
#                                     per-row execution, aggregated
#                                     report linking sub-incidents.
#    Invoke-IncidentTabletop      -- IR-team exercise mode. Loads
#                                     a scenario from
#                                     templates/tabletop-scenarios/,
#                                     runs the full playbook in
#                                     PREVIEW against a sandbox
#                                     user, produces a graded
#                                     report ("here is what your
#                                     team would have done in
#                                     this scenario").
# ============================================================

# ============================================================
#  Bulk incident response
# ============================================================

function Invoke-BulkIncidentResponse {
    <#
        Run Invoke-CompromisedAccountResponse for every row in a
        CSV. CSV columns: UPN, Severity, Reason, QuarantineSentMail.

        Validate-first pattern from Phase 1:
          1. Parse + validate the CSV.
          2. Print summary + ask for confirmation.
          3. Per-row execute (does NOT halt on per-row failure).
          4. Write a result CSV next to the input.
          5. Aggregate-report HTML at <stateDir>\<tenant>\incidents\
             bulk-<timestamp>\index.html linking each sub-incident.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$WhatIf,
        [switch]$NonInteractive
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-ErrorMsg "CSV not found: $Path"; return $null
    }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) {
        Write-Warn "CSV is empty: $Path"; return $null
    }

    # ---- Validate ----
    $errors = New-Object System.Collections.ArrayList
    $normalized = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]; $row = $i + 2   # +2 = 1-based + header
        $upn = [string]$r.UPN
        $sev = if ($r.Severity) { [string]$r.Severity } else { 'High' }
        $reason = [string]$r.Reason
        $q = $false
        if ($r.PSObject.Properties.Name -contains 'QuarantineSentMail') {
            $q = ([string]$r.QuarantineSentMail) -match '^(?i:true|yes|1)$'
        }
        if ([string]::IsNullOrWhiteSpace($upn))       { [void]$errors.Add("Row ${row}: missing UPN") }
        elseif ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { [void]$errors.Add("Row ${row}: invalid UPN '$upn'") }
        if ($sev -notin 'Low','Medium','High','Critical') { [void]$errors.Add("Row ${row}: invalid Severity '$sev' (Low|Medium|High|Critical)") }
        [void]$normalized.Add(@{ UPN=$upn; Severity=$sev; Reason=$reason; QuarantineSentMail=$q; RowNum=$row })
    }
    if ($errors.Count -gt 0) {
        Write-ErrorMsg ("Bulk validation failed ({0} error(s)):" -f $errors.Count)
        foreach ($e in $errors) { Write-Warn "  $e" }
        return $null
    }

    # ---- Summarize + confirm ----
    Write-SectionHeader "Bulk incident response"
    Write-StatusLine "Input"    $Path 'White'
    Write-StatusLine "Rows"     ("{0}" -f $normalized.Count) 'White'
    $bySev = $normalized | Group-Object -Property { $_.Severity }
    foreach ($g in $bySev) { Write-StatusLine ("  " + $g.Name) ("{0} row(s)" -f $g.Count) 'White' }
    if (-not $NonInteractive) {
        if (-not (Confirm-Action ("Run compromised-account response on all {0} rows?" -f $normalized.Count))) {
            Write-InfoMsg "Cancelled."; return $null
        }
    }

    # ---- Execute ----
    $bulkId = ("bulk-{0}" -f (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $bulkDir = Join-Path (Get-IncidentsDirectory) $bulkId
    New-Item -ItemType Directory -Path $bulkDir -Force | Out-Null
    $resultRows = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($n in $normalized) {
        $i++
        Write-Host ""
        Write-Host ("===== [{0}/{1}] {2} (severity {3}) =====" -f $i, $normalized.Count, $n.UPN, $n.Severity) -ForegroundColor Cyan
        $incidentId = $null
        $status = 'failed'
        $reason = ''
        try {
            $incidentId = Invoke-CompromisedAccountResponse `
                -UPN $n.UPN `
                -Severity $n.Severity `
                -Reason $n.Reason `
                -QuarantineSentMail:$n.QuarantineSentMail `
                -WhatIf:$WhatIf `
                -NonInteractive:$NonInteractive
            $status = if ($incidentId) { if ($WhatIf -or (Get-PreviewMode)) { 'Preview' } else { 'Success' } } else { 'Failed' }
            $reason = if ($incidentId) { '' } else { 'response returned $null' }
        } catch {
            $reason = "exception: $($_.Exception.Message)"
            Write-ErrorMsg "Row $($n.RowNum) failed: $reason"
        }
        [void]$resultRows.Add([PSCustomObject]@{
            Row         = $n.RowNum
            UPN         = $n.UPN
            Severity    = $n.Severity
            IncidentId  = $incidentId
            Status      = $status
            Reason      = $reason
        })
    }

    # ---- Result CSV ----
    $resultCsv = $Path -replace '\.csv$','' + ('-result-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.csv')
    $resultRows | Export-Csv -LiteralPath $resultCsv -NoTypeInformation -Encoding UTF8
    Write-Success "Result CSV: $resultCsv"

    # ---- Aggregate report ----
    $aggregate = New-BulkIncidentReport -BulkId $bulkId -BulkDir $bulkDir -ResultRows @($resultRows) -InputPath $Path
    Write-Success "Aggregate report: $aggregate"

    return @{ BulkId = $bulkId; ResultCsv = $resultCsv; AggregateReport = $aggregate; Rows = @($resultRows) }
}

function New-BulkIncidentReport {
    <#
        Single-page HTML linking each sub-incident's individual
        report.html. Lives at <bulkDir>\index.html so the operator
        can click through into per-incident detail.
    #>
    param(
        [Parameter(Mandatory)][string]$BulkId,
        [Parameter(Mandatory)][string]$BulkDir,
        [Parameter(Mandatory)][array]$ResultRows,
        [Parameter(Mandatory)][string]$InputPath
    )
    $rows = New-Object System.Collections.ArrayList
    foreach ($r in $ResultRows) {
        $link = if ($r.IncidentId) {
            $reportPath = Join-Path (Join-Path (Get-IncidentsDirectory) $r.IncidentId) 'report.html'
            if (Test-Path -LiteralPath $reportPath) {
                '<a href="' + ($reportPath -replace '\\','/') + '">' + [System.Net.WebUtility]::HtmlEncode($r.IncidentId) + '</a>'
            } else {
                [System.Net.WebUtility]::HtmlEncode($r.IncidentId)
            }
        } else { '-' }
        $color = switch ($r.Status) { 'Success' { 'green' } 'Preview' { '#888' } default { '#a00' } }
        [void]$rows.Add(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td style='color:{4}'>{5}</td><td>{6}</td></tr>" -f `
            $r.Row,
            [System.Net.WebUtility]::HtmlEncode([string]$r.UPN),
            [System.Net.WebUtility]::HtmlEncode([string]$r.Severity),
            $link,
            $color,
            [System.Net.WebUtility]::HtmlEncode([string]$r.Status),
            [System.Net.WebUtility]::HtmlEncode([string]$r.Reason)))
    }
    $rowsHtml = ($rows -join "`n")
    $tenant = if ($script:SessionState -and $script:SessionState.TenantName) { $script:SessionState.TenantName } else { 'unknown' }
    $modeLabel = if (Get-PreviewMode) { 'PREVIEW' } else { 'LIVE' }
    $html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'><title>Bulk incident response -- $BulkId</title>
<style>
  body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222;max-width:1000px;margin:1em auto;padding:0 1em}
  h1{color:#a00;border-bottom:2px solid #a00;padding-bottom:4px}
  table{border-collapse:collapse;font-size:13px;width:100%}
  th,td{border:1px solid #ddd;padding:6px 10px;text-align:left}
  th{background:#f4f4f4}
  a{color:#06c;text-decoration:none}
  a:hover{text-decoration:underline}
</style></head><body>
<h1>Bulk incident response -- $BulkId</h1>
<table>
  <tr><th>Tenant</th><td>$([System.Net.WebUtility]::HtmlEncode($tenant))</td></tr>
  <tr><th>Mode</th><td>$modeLabel</td></tr>
  <tr><th>Started (UTC)</th><td>$((Get-Date).ToUniversalTime().ToString('o'))</td></tr>
  <tr><th>Input CSV</th><td><code>$([System.Net.WebUtility]::HtmlEncode($InputPath))</code></td></tr>
  <tr><th>Sub-incidents</th><td>$($ResultRows.Count)</td></tr>
</table>

<h2>Per-row results</h2>
<table>
  <tr><th>Row</th><th>UPN</th><th>Severity</th><th>Incident</th><th>Status</th><th>Reason</th></tr>
$rowsHtml
</table>

<p style='color:#888;font-size:12px;margin-top:2em'>Generated by M365 Manager bulk incident response.</p>
</body></html>
"@
    $path = Join-Path $BulkDir 'index.html'
    Set-Content -LiteralPath $path -Value $html -Encoding UTF8 -Force
    return $path
}

# ============================================================
#  Tabletop mode
# ============================================================

function Get-TabletopScenariosDir {
    $root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    return Join-Path $root 'templates/tabletop-scenarios'
}

function Get-TabletopScenarios {
    <#
        Enumerate available tabletop scenarios under
        templates/tabletop-scenarios/. Returns hashtables with
        Name, Description, Severity, ExpectedActions.
    #>
    $dir = Get-TabletopScenariosDir
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $out = New-Object System.Collections.ArrayList
    foreach ($f in Get-ChildItem -LiteralPath $dir -Filter '*.json' -File) {
        try {
            $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json -ErrorAction Stop
            [void]$out.Add(@{ Name = [string]$j.name; Path = $f.FullName; Description = [string]$j.description; Severity = [string]$j.severity; Raw = $j })
        } catch { Write-Warn "Skipped malformed scenario $($f.Name): $($_.Exception.Message)" }
    }
    return @($out)
}

function Invoke-IncidentTabletop {
    <#
        Run a scenario from templates/tabletop-scenarios/ in
        PREVIEW mode against the configured tabletop user. Produces
        a graded report covering: which playbook steps fired, in
        what order, how long it took wall-clock, what the operator
        would have done (per the scenario's expectedActions list)
        vs what the playbook actually did. Useful for compliance
        audits demonstrating IR readiness.

        Does NOT touch any real user. The scenario JSON declares
        a sandbox UPN; we run the playbook with -WhatIf so the
        only artifacts written are forensic snapshots + audit
        entries.

        Configurable tabletop user via IncidentResponse.TabletopUPN
        (default sets to a placeholder; operator should set this
        to a known sandbox account in their tenant).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScenarioName,
        [string]$TabletopUPN
    )
    $scenarios = Get-TabletopScenarios
    $scenario = $scenarios | Where-Object { $_.Name -eq $ScenarioName } | Select-Object -First 1
    if (-not $scenario) {
        Write-ErrorMsg "Unknown scenario '$ScenarioName'. Available:"
        foreach ($s in $scenarios) { Write-Warn "  - $($s.Name) : $($s.Description)" }
        return $null
    }

    # Resolve sandbox UPN
    if (-not $TabletopUPN) {
        if (Get-Command Get-EffectiveConfig -ErrorAction SilentlyContinue) {
            $TabletopUPN = [string](Get-EffectiveConfig -Key 'IncidentResponse.TabletopUPN')
        }
        if (-not $TabletopUPN) {
            Write-ErrorMsg "No -TabletopUPN passed and IncidentResponse.TabletopUPN not configured. Pass -TabletopUPN <sandbox account> or set the config key."
            return $null
        }
    }

    Write-SectionHeader "Incident tabletop -- $ScenarioName"
    Write-StatusLine "Scenario"  $scenario.Description 'White'
    Write-StatusLine "Severity"  $scenario.Severity    'White'
    Write-StatusLine "Sandbox"   $TabletopUPN          'White'
    Write-Warn "Tabletop runs the playbook in PREVIEW mode. No tenant state will change."

    $start = Get-Date
    $incidentId = Invoke-CompromisedAccountResponse `
        -UPN $TabletopUPN `
        -Severity $scenario.Severity `
        -Reason ("Tabletop exercise: " + $scenario.Description) `
        -WhatIf `
        -NonInteractive
    $elapsed = (Get-Date) - $start

    # Grade against expectedActions
    $expected = @($scenario.Raw.expectedActions)
    $actualSteps = @()
    if ($incidentId -and (Get-Command Read-AuditEntries -ErrorAction SilentlyContinue)) {
        $actualSteps = @((Read-AuditEntries | Where-Object { $_.target -and $_.target.incidentId -eq $incidentId -and $_.actionType -like 'Incident:*' }) | ForEach-Object { [string]$_.actionType })
    }
    $matched = 0
    $missing = New-Object System.Collections.ArrayList
    foreach ($a in $expected) {
        $needle = "Incident:" + [string]$a
        if ($actualSteps -contains $needle) { $matched++ } else { [void]$missing.Add([string]$a) }
    }

    $grade = if ($expected.Count -eq 0) { 'N/A' } else { ("{0}/{1}" -f $matched, $expected.Count) }
    Write-Host ""
    Write-Host "  TABLETOP GRADE: $grade" -ForegroundColor $(if ($missing.Count -eq 0) { 'Green' } else { 'Yellow' })
    Write-StatusLine "Wall-clock"     ("{0:N1}s" -f $elapsed.TotalSeconds) 'White'
    Write-StatusLine "Steps observed" ("{0}" -f @($actualSteps).Count) 'White'
    if ($missing.Count -gt 0) {
        Write-Warn ("Missing actions ({0}): {1}" -f $missing.Count, ($missing -join ', '))
    }

    # Write tabletop report alongside the incident dir
    if ($incidentId) {
        $reportHtml = New-TabletopReport -IncidentId $incidentId -ScenarioName $ScenarioName -Scenario $scenario -ActualSteps $actualSteps -Missing @($missing) -Grade $grade -Elapsed $elapsed
        Write-Success "Tabletop report: $reportHtml"
    }

    return @{
        ScenarioName = $ScenarioName
        IncidentId   = $incidentId
        Grade        = $grade
        WallClockSec = [math]::Round($elapsed.TotalSeconds, 1)
        Missing      = @($missing)
        ActualSteps  = $actualSteps
    }
}

function New-TabletopReport {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$ScenarioName,
        [Parameter(Mandatory)]$Scenario,
        [Parameter(Mandatory)][array]$ActualSteps,
        [Parameter(Mandatory)][array]$Missing,
        [Parameter(Mandatory)][string]$Grade,
        [Parameter(Mandatory)][TimeSpan]$Elapsed
    )
    $dir = Get-IncidentDirectory -IncidentId $IncidentId
    if (-not $dir) { return $null }

    $missingHtml = if ($Missing.Count -eq 0) { '<li><em>(none -- all expected actions observed)</em></li>' } else {
        ($Missing | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }) -join "`n"
    }
    $stepsHtml = ($ActualSteps | ForEach-Object { "<li><code>$([System.Net.WebUtility]::HtmlEncode($_))</code></li>" }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'><title>Tabletop: $ScenarioName -- $IncidentId</title>
<style>body{font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222;max-width:980px;margin:1em auto;padding:0 1em} h1{color:#06c;border-bottom:2px solid #06c;padding-bottom:4px} table{border-collapse:collapse;font-size:13px;width:100%} th,td{border:1px solid #ddd;padding:6px 10px;text-align:left} th{background:#f4f4f4} code{background:#f4f4f4;padding:1px 5px;border-radius:3px;font-size:12px}</style>
</head><body>
<h1>Incident response tabletop: $([System.Net.WebUtility]::HtmlEncode($ScenarioName))</h1>
<table>
  <tr><th>Scenario</th><td>$([System.Net.WebUtility]::HtmlEncode([string]$Scenario.Description))</td></tr>
  <tr><th>Linked incident</th><td><code>$IncidentId</code> (PREVIEW)</td></tr>
  <tr><th>Severity</th><td>$([System.Net.WebUtility]::HtmlEncode([string]$Scenario.Severity))</td></tr>
  <tr><th>Grade</th><td><b>$Grade</b></td></tr>
  <tr><th>Wall-clock</th><td>$([math]::Round($Elapsed.TotalSeconds, 1)) seconds</td></tr>
</table>
<h2>Expected actions missed</h2>
<ul>$missingHtml</ul>
<h2>Playbook steps observed</h2>
<ul>$stepsHtml</ul>
<h2>Compliance note</h2>
<p>This tabletop exercise ran the compromised-account response playbook in PREVIEW mode. No tenant state was changed. The incident snapshot dir at <code>$dir</code> contains the forensic baseline the playbook captured plus the per-step audit entries -- suitable for an auditor reviewing IR readiness.</p>
</body></html>
"@
    $p = Join-Path $dir 'tabletop-report.html'
    Set-Content -LiteralPath $p -Value $html -Encoding UTF8 -Force
    return $p
}
