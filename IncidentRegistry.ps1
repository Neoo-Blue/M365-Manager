# ============================================================
#  IncidentRegistry.ps1 -- view / close / undo / export helpers
#
#  Companion to IncidentResponse.ps1. The playbook writes one
#  jsonl line per phase to <tenant>/incidents.jsonl (running,
#  completed, closed, false-positive) -- this module reads
#  those records back, exposes them through Get-Incident +
#  Get-Incidents, and provides the operator-facing close /
#  undo / export operations.
#
#  Tenant-scoped: every function reads the current tenant's
#  registry. Switching tenant via Switch-Tenant changes the
#  visible incident set.
# ============================================================

# ============================================================
#  Read side
# ============================================================

function Read-IncidentRegistry {
    <#
        Return every JSONL record from the current tenant's
        registry, newest-first. Each `id` may appear multiple
        times -- a running record + a completed record + an
        optional closed record. Callers that want a single
        current view use Get-Incident which folds them down.
    #>
    $path = Get-IncidentRegistryPath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return @() }
    $records = New-Object System.Collections.ArrayList
    foreach ($line in Get-Content -LiteralPath $path) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $obj = $line | ConvertFrom-Json -ErrorAction Stop
            [void]$records.Add($obj)
        } catch { Write-Warn "Skipped malformed registry line: $($_.Exception.Message)" }
    }
    return @($records)
}

function Get-Incident {
    <#
        Retrieve one incident -- folds together every JSONL
        record sharing the id (running / completed / closed)
        into a single PSCustomObject + lists the on-disk
        artifact filenames.

        Returns $null on miss.
    #>
    param([Parameter(Mandatory)][string]$Id)
    $all = Read-IncidentRegistry
    $matches = @($all | Where-Object { $_.id -eq $Id })
    if ($matches.Count -eq 0) { return $null }

    # Newest record wins on conflicting fields (status, completedUtc, etc.)
    $base = $matches[0]
    foreach ($prop in $base.PSObject.Properties) { }
    # Use most-recent record (Add-Content appends in order, so the last
    # matching record in the file is the latest).
    $merged = @{}
    foreach ($rec in $matches) {
        foreach ($prop in $rec.PSObject.Properties) {
            if ($null -ne $prop.Value) { $merged[$prop.Name] = $prop.Value }
        }
    }

    # Inventory the per-incident dir
    $dir = Get-IncidentDirectory -IncidentId $Id
    $artifacts = @()
    if ($dir -and (Test-Path -LiteralPath $dir)) {
        $artifacts = @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
    }
    $merged.directory = $dir
    $merged.artifacts = $artifacts

    return [PSCustomObject]$merged
}

function Get-Incidents {
    <#
        List incidents in the current tenant. Filters:
          -Status   : Open | Closed | All (default Open)
          -Days     : only incidents started within N days
          -Severity : Low | Medium | High | Critical
    #>
    param(
        [ValidateSet('Open','Closed','All')][string]$Status = 'Open',
        [int]$Days,
        [ValidateSet('Low','Medium','High','Critical')][string]$Severity
    )
    $all = Read-IncidentRegistry
    # Collapse by id so each incident appears once with the latest record's fields.
    $byId = @{}
    foreach ($rec in $all) {
        $byId[[string]$rec.id] = $rec
    }
    $merged = @()
    foreach ($id in $byId.Keys) { $merged += (Get-Incident -Id $id) }
    $filtered = $merged

    if ($Status -ne 'All') {
        $filtered = @($filtered | Where-Object {
            $isClosed = $_.status -in 'closed','false-positive'
            if ($Status -eq 'Closed') { $isClosed } else { -not $isClosed }
        })
    }
    if ($Days) {
        $cutoff = (Get-Date).AddDays(-1 * $Days).ToUniversalTime()
        $filtered = @($filtered | Where-Object {
            $ts = [DateTime]::MinValue
            [DateTime]::TryParse([string]$_.startedUtc, [ref]$ts) | Out-Null
            $ts -ge $cutoff
        })
    }
    if ($Severity) {
        $filtered = @($filtered | Where-Object { [string]$_.severity -eq $Severity })
    }
    return @($filtered | Sort-Object { [string]$_.startedUtc } -Descending)
}

function Show-Incidents {
    <#
        Compact table -- used by the Incident menu and by the AI
        Get-IncidentTimeline tool. Re-runs Get-Incidents and
        renders id / severity / status / upn / startedUtc.
    #>
    param(
        [ValidateSet('Open','Closed','All')][string]$Status = 'Open',
        [int]$Days = 30,
        [ValidateSet('Low','Medium','High','Critical')][string]$Severity
    )
    $args = @{ Status = $Status; Days = $Days }
    if ($Severity) { $args.Severity = $Severity }
    $list = Get-Incidents @args
    if (-not $list -or $list.Count -eq 0) {
        Write-Host "  (no incidents matched)" -ForegroundColor DarkGray
        return
    }
    Write-Host ""
    Write-Host ("  {0,-22} {1,-9} {2,-13} {3,-32} {4}" -f 'INCIDENT','SEVERITY','STATUS','UPN','STARTED (UTC)') -ForegroundColor White
    Write-Host ("  " + ('-' * 100)) -ForegroundColor DarkGray
    foreach ($i in $list) {
        $sev = [string]$i.severity
        $col = switch ($sev) { 'Critical' { 'Red' } 'High' { 'Red' } 'Medium' { 'Yellow' } default { 'Gray' } }
        Write-Host ("  {0,-22} {1,-9} {2,-13} {3,-32} {4}" -f `
            [string]$i.id, $sev, [string]$i.status, [string]$i.upn, [string]$i.startedUtc) -ForegroundColor $col
    }
    Write-Host ""
}

function Show-IncidentReport {
    <#
        Open the report.html in the operator's default browser.
        Falls back to Start-Process / xdg-open / open depending
        on platform. Returns the path opened.
    #>
    param([Parameter(Mandatory)][string]$Id)
    $i = Get-Incident -Id $Id
    if (-not $i) { Write-Warn "No incident '$Id'."; return $null }
    $dir = $i.directory
    if (-not $dir) { Write-Warn "Incident '$Id' has no on-disk artifacts."; return $null }
    $report = Join-Path $dir 'report.html'
    if (-not (Test-Path -LiteralPath $report)) { Write-Warn "No report.html in $dir."; return $null }
    try {
        if ($env:LOCALAPPDATA) {
            Start-Process $report -ErrorAction Stop
        } elseif (Get-Command open -ErrorAction SilentlyContinue) {
            & open $report
        } elseif (Get-Command xdg-open -ErrorAction SilentlyContinue) {
            & xdg-open $report
        } else {
            Write-InfoMsg "No browser launcher found; report path: $report"
        }
    } catch { Write-Warn "Could not open report: $($_.Exception.Message). Path: $report" }
    return $report
}

# ============================================================
#  Close / undo / export
# ============================================================

function Close-Incident {
    <#
        Mark an incident closed. Appends a new JSONL record with
        status='closed' (or 'false-positive') and the resolution
        notes. Idempotent -- closing an already-closed incident
        is a warning, not an error.

        When -FalsePositive is passed AND no -SkipUndo is set,
        invokes Undo-Incident -- the operator confirms each
        reversal individually.
    #>
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Resolution,
        [switch]$FalsePositive,
        [switch]$SkipUndo
    )
    $i = Get-Incident -Id $Id
    if (-not $i) { Write-ErrorMsg "No incident '$Id'."; return $false }
    if ($i.status -in 'closed','false-positive') {
        Write-Warn "Incident '$Id' is already $($i.status). Re-closing with new resolution."
    }

    $status = if ($FalsePositive) { 'false-positive' } else { 'closed' }
    $closedUtc = (Get-Date).ToUniversalTime().ToString('o')
    $closer = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'unknown' }

    Write-IncidentRegistryRecord -Record ([ordered]@{
        id            = $Id
        upn           = [string]$i.upn
        severity      = [string]$i.severity
        startedUtc    = [string]$i.startedUtc
        completedUtc  = [string]$i.completedUtc
        closedUtc     = $closedUtc
        closedBy      = $closer
        status        = $status
        resolution    = $Resolution
        falsePositive = [bool]$FalsePositive
        reportPath    = [string]$i.reportPath
    })
    Write-IncidentAuditEntry -IncidentId $Id -EventType 'INCIDENT_CLOSE' -Detail ("Incident closed: {0}" -f $Resolution) -ActionType 'Incident:Close' -Target @{ userUpn = [string]$i.upn; status = $status; resolution = $Resolution } -Result 'info'
    Write-Success "Incident $Id marked $status."

    if ($FalsePositive -and -not $SkipUndo) {
        Write-Host ""
        Write-Host "  False-positive: walking reversible steps for confirmation..." -ForegroundColor Yellow
        Undo-Incident -Id $Id | Out-Null
    }
    return $true
}

function Undo-Incident {
    <#
        Walk this incident's audit entries in reverse and route
        each reversible one through Invoke-Undo. Operator confirms
        each step individually -- there is no -Confirm:$false
        bypass on this path by design.

        Returns the number of reversals attempted.
    #>
    param([Parameter(Mandatory)][string]$Id)
    if (-not (Get-Command Invoke-Undo -ErrorAction SilentlyContinue)) {
        Write-ErrorMsg "Invoke-Undo not loaded (Undo.ps1 missing?)."
        return 0
    }
    if (-not (Get-Command Read-AuditEntries -ErrorAction SilentlyContinue)) {
        Write-ErrorMsg "Read-AuditEntries not loaded (AuditViewer.ps1 missing?)."
        return 0
    }
    $entries = @(Read-AuditEntries)
    # Pull only entries tied to this incident with a non-null reverse recipe.
    $reversible = @($entries | Where-Object {
        $_.target -and $_.target.incidentId -eq $Id -and $_.reverse -and $_.reverse.type
    })
    if ($reversible.Count -eq 0) {
        Write-InfoMsg "No reversible audit entries found for $Id."
        return 0
    }
    # Latest-first
    $reversible = @($reversible | Sort-Object { $_.ts } -Descending)
    Write-Host ""
    Write-Host ("  $Id : {0} reversible step(s)" -f $reversible.Count) -ForegroundColor Cyan
    $attempted = 0
    foreach ($e in $reversible) {
        Write-Host ""
        Write-Host ("    Reverse: {0}" -f [string]$e.reverse.description) -ForegroundColor White
        Write-Host ("      original entryId: $($e.entryId)") -ForegroundColor DarkGray
        $ans = Read-Host "    [Y]es / [N]o / [Q]uit"
        if ($ans -match '^[Qq]') { break }
        if ($ans -notmatch '^[Yy]') { continue }
        try {
            Invoke-Undo -EntryId $e.entryId | Out-Null
            $attempted++
        } catch { Write-Warn "Undo failed: $($_.Exception.Message)" }
    }
    Write-Host ""
    Write-Success "Attempted $attempted reversal(s) on $Id."
    Write-IncidentAuditEntry -IncidentId $Id -EventType 'INCIDENT_UNDO' -Detail ("Undo walk over {0} reversible step(s), attempted {1}" -f $reversible.Count, $attempted) -ActionType 'Incident:Undo' -Target @{ attempted = $attempted } -Result 'info'
    return $attempted
}

function Export-Incident {
    <#
        Bundle the snapshot dir + audit-log entries + report into
        a single zip for compliance handoff. Path defaults to a
        timestamped file in the current directory.
    #>
    param(
        [Parameter(Mandatory)][string]$Id,
        [string]$Path
    )
    $i = Get-Incident -Id $Id
    if (-not $i) { Write-ErrorMsg "No incident '$Id'."; return $null }
    $dir = $i.directory
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { Write-ErrorMsg "Incident '$Id' has no artifact directory."; return $null }

    if (-not $Path) {
        $Path = Join-Path (Get-Location).Path ("incident-export-{0}-{1}.zip" -f $Id, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    }

    # Stage a temp dir so we can include both the artifact dir and a
    # filtered slice of the session audit log (incident-only entries).
    $stage = Join-Path ([IO.Path]::GetTempPath()) ("incident-export-" + [Guid]::NewGuid())
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        # Copy artifacts
        Copy-Item -Path (Join-Path $dir '*') -Destination $stage -Recurse -Force

        # Filter the live audit log down to this incident's lines + write inline
        if (Get-Command Read-AuditEntries -ErrorAction SilentlyContinue) {
            $entries = @(Read-AuditEntries | Where-Object {
                $_.target -and $_.target.incidentId -eq $Id
            } | Sort-Object { $_.ts })
            if ($entries.Count -gt 0) {
                $auditPath = Join-Path $stage 'audit-entries.jsonl'
                foreach ($e in $entries) { Add-Content -LiteralPath $auditPath -Value ($e | ConvertTo-Json -Depth 10 -Compress) }
            }
        }

        # Index file -- the registry record + a quick orientation note
        $index = [ordered]@{
            id           = $Id
            exportedUtc  = (Get-Date).ToUniversalTime().ToString('o')
            exportedBy   = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'unknown' }
            tenantName   = [string]$i.tenantName
            registryRecord = $i
            note         = 'Bundle includes snapshot.json, audit-24h.json, mail-sent-7d.json, shares-7d.json, inbox-rules.json, report.html, audit-entries.jsonl, and any temp-password.txt if it was written during the run. Hand off to compliance via secure channel; the bundle may contain sensitive forensic data.'
        }
        Set-Content -LiteralPath (Join-Path $stage 'INDEX.json') -Value ($index | ConvertTo-Json -Depth 12) -Encoding UTF8 -Force

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $Path -Force
        Write-Success ("Exported incident $Id to $Path")
        Write-IncidentAuditEntry -IncidentId $Id -EventType 'INCIDENT_EXPORT' -Detail "Exported to $Path" -ActionType 'Incident:Export' -Target @{ path = $Path } -Result 'info'
        return $Path
    } finally {
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
#  Menu (slot 22)
# ============================================================

function Start-IncidentResponseMenu {
    <#
        Operator menu -- routes to playbook trigger, list,
        view, close, export, undo.
    #>
    $running = $true
    while ($running) {
        Write-SectionHeader "Incident Response"
        $sel = Show-Menu -Title "Pick an action" -Options @(
            "Run compromised-account response (single UPN)",
            "List open incidents",
            "List incidents (all)",
            "View incident report",
            "Close incident",
            "Mark incident as false positive (with undo walk)",
            "Undo an incident's reversible steps",
            "Export incident for compliance handoff"
        ) -BackLabel "Back"
        switch ($sel) {
            0 {
                $upn = Read-UserInput "Compromised UPN"
                if (-not $upn) { continue }
                $sevSel = Show-Menu -Title "Severity" -Options @('Low (forensic only)','Medium (contain)','High (default)','Critical (full + quarantine prompt)') -BackLabel "Cancel"
                if ($sevSel -eq -1) { continue }
                $sev = @('Low','Medium','High','Critical')[$sevSel]
                $reason = Read-UserInput "Reason (optional)"
                $quarantine = $false
                if ($sev -eq 'Critical') {
                    if (Confirm-Action "Enable -QuarantineSentMail (purges 7d sent items, irreversible)?") { $quarantine = $true }
                }
                Invoke-CompromisedAccountResponse -UPN $upn -Severity $sev -Reason $reason -QuarantineSentMail:$quarantine | Out-Null
            }
            1 { Show-Incidents -Status Open }
            2 { Show-Incidents -Status All }
            3 {
                $id = Read-UserInput "Incident id"
                if ($id) { Show-IncidentReport -Id $id | Out-Null }
            }
            4 {
                $id = Read-UserInput "Incident id"
                if (-not $id) { continue }
                $res = Read-UserInput "Resolution notes"
                if ($res) { Close-Incident -Id $id -Resolution $res | Out-Null }
            }
            5 {
                $id = Read-UserInput "Incident id"
                if (-not $id) { continue }
                $res = Read-UserInput "False-positive reason"
                if ($res) { Close-Incident -Id $id -Resolution $res -FalsePositive | Out-Null }
            }
            6 {
                $id = Read-UserInput "Incident id"
                if ($id) { Undo-Incident -Id $id | Out-Null }
            }
            7 {
                $id = Read-UserInput "Incident id"
                if (-not $id) { continue }
                $p = Read-UserInput "Destination path (blank = current dir)"
                if ($p) { Export-Incident -Id $id -Path $p | Out-Null }
                else    { Export-Incident -Id $id | Out-Null }
            }
            -1 { $running = $false }
        }
    }
}
