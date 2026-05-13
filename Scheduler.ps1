# ============================================================
#  Scheduler.ps1 -- recurring health checks
#
#  Primary target: Windows Task Scheduler (Register-ScheduledTask).
#  Cross-platform path: shell out to `crontab` on macOS / Linux.
#
#  Scheduled scripts run with -NonInteractive so prompt paths
#  in UI.ps1 either return safe defaults or fail fast.
#
#  Credential model:
#    Scheduled tasks need auth without a human at the console. We
#    store an encrypted service-principal credential file at
#    <stateDir>/scheduler-cred.xml -- DPAPI-protected on Windows
#    via the same Protect-Secret pattern as the AI key. Operator
#    runs Register-SchedulerCredential ONCE; runs after that
#    auto-authenticate against Graph with the stored creds.
# ============================================================

$script:SchedulerStateName = 'scheduled-checks.json'

# ============================================================
#  Schedule parsing
# ============================================================

function ConvertTo-ScheduleSpec {
    <#
        Parse a friendly schedule string into a structured spec:
          'Daily 09:00'           -> @{ Frequency='Daily';  Time='09:00' }
          'Weekly Mon 09:00'      -> @{ Frequency='Weekly'; Day='Mon'; Time='09:00' }
          'Monthly 1 09:00'       -> @{ Frequency='Monthly';Day=1;     Time='09:00' }
          'Hourly'                -> @{ Frequency='Hourly' }
          'cron 0 9 * * *'        -> @{ Frequency='Cron';   Expression='0 9 * * *' }
        Returns $null if unrecognized.
    #>
    param([Parameter(Mandatory)][string]$Schedule)
    $s = $Schedule.Trim()
    if ($s -match '^(?i)cron\s+(.+)$')                          { return @{ Frequency='Cron';    Expression=$Matches[1] } }
    if ($s -match '^(?i)Hourly$')                                { return @{ Frequency='Hourly' } }
    if ($s -match '^(?i)Daily\s+(\d{1,2}:\d{2})$')              { return @{ Frequency='Daily';   Time=$Matches[1] } }
    if ($s -match '^(?i)Weekly\s+(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(\d{1,2}:\d{2})$') {
        return @{ Frequency='Weekly';  Day=$Matches[1]; Time=$Matches[2] }
    }
    if ($s -match '^(?i)Monthly\s+(\d{1,2})\s+(\d{1,2}:\d{2})$') {
        return @{ Frequency='Monthly'; Day=[int]$Matches[1]; Time=$Matches[2] }
    }
    return $null
}

function ConvertTo-CronExpression {
    <#
        Friendly schedule -> 5-field cron expression. Used on
        non-Windows platforms where we manage the crontab.
    #>
    param([Parameter(Mandatory)][hashtable]$Spec)
    switch ($Spec.Frequency) {
        'Cron'    { return $Spec.Expression }
        'Hourly'  { return '0 * * * *' }
        'Daily'   { $hm = $Spec.Time -split ':'; return "$([int]$hm[1]) $([int]$hm[0]) * * *" }
        'Weekly'  {
            $dow = @{ Sun=0;Mon=1;Tue=2;Wed=3;Thu=4;Fri=5;Sat=6 }[$Spec.Day]
            $hm = $Spec.Time -split ':'
            return "$([int]$hm[1]) $([int]$hm[0]) * * $dow"
        }
        'Monthly' {
            $hm = $Spec.Time -split ':'
            return "$([int]$hm[1]) $([int]$hm[0]) $([int]$Spec.Day) * *"
        }
    }
    return $null
}

# ============================================================
#  State file
# ============================================================

function Get-SchedulerStatePath {
    $dir = Get-StateDirectory
    if (-not $dir) { return $null }
    return Join-Path $dir $script:SchedulerStateName
}

function Read-SchedulerState {
    $p = Get-SchedulerStatePath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return @() }
    try { return @((Get-Content -LiteralPath $p -Raw | ConvertFrom-Json)) } catch { return @() }
}

function Write-SchedulerState {
    param([Parameter(Mandatory)][array]$Records)
    $p = Get-SchedulerStatePath
    if (-not $p) { return }
    try { ($Records | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $p -Encoding UTF8 -Force }
    catch { Write-Warn "Could not write scheduler state: $_" }
}

# ============================================================
#  Credential storage (DPAPI-protected via Phase 0.5 helpers)
# ============================================================

function Get-SchedulerCredentialPath {
    $dir = Get-StateDirectory
    if (-not $dir) { return $null }
    return Join-Path $dir 'scheduler-cred.xml'
}

function Register-SchedulerCredential {
    <#
        One-time setup: ask for tenant id + app id + client secret
        (or cert path) and save encrypted. Subsequent scheduled runs
        load via Get-SchedulerCredential. We deliberately don't
        store interactive-flow tokens because they expire and
        require interactive refresh.
    #>
    $tenantId = Read-UserInput "Tenant ID (Azure AD directory id GUID)"
    $appId    = Read-UserInput "Application (client) ID"
    $secret   = Read-UserInput "Client secret (will be DPAPI-encrypted)"
    if (-not $tenantId -or -not $appId -or -not $secret) {
        Write-Warn "All three values required."; return $false
    }
    $encrypted = Protect-ApiKey -PlainKey $secret
    $cred = @{
        tenantId       = $tenantId
        appId          = $appId
        encryptedSecret= $encrypted
        registeredAt   = (Get-Date).ToUniversalTime().ToString('o')
    }
    $p = Get-SchedulerCredentialPath
    if (-not $p) { Write-ErrorMsg "Cannot resolve state directory."; return $false }
    ($cred | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $p -Encoding UTF8 -Force
    Write-Success "Scheduler credential stored (encrypted) at $p"
    return $true
}

function Get-SchedulerCredential {
    <#
        Decrypt the stored cred and return @{ TenantId; AppId;
        Secret } -- caller passes Secret to Connect-MgGraph as
        the client secret.
    #>
    $p = Get-SchedulerCredentialPath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) {
        Write-Warn "No scheduler credential registered. Run Register-SchedulerCredential first."
        return $null
    }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        $secret = Unprotect-ApiKey -StoredKey $raw.encryptedSecret
        return @{ TenantId = $raw.tenantId; AppId = $raw.appId; Secret = $secret; RegisteredAt = $raw.registeredAt }
    } catch { Write-ErrorMsg "Could not load scheduler credential: $_"; return $null }
}

# ============================================================
#  Platform abstraction
# ============================================================

function Test-IsWindowsHost { return ($null -ne $env:LOCALAPPDATA -or [System.Environment]::OSVersion.Platform -eq 'Win32NT') }

function Register-WindowsScheduledHealthCheck {
    param([string]$Name, [string]$Command, [hashtable]$Spec)
    if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
        Write-ErrorMsg "ScheduledTasks module not available on this host."
        return $false
    }
    $action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $Command
    $trigger = switch ($Spec.Frequency) {
        'Daily'   { New-ScheduledTaskTrigger -Daily -At $Spec.Time }
        'Weekly'  { New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Spec.Day -At $Spec.Time }
        'Monthly' { New-ScheduledTaskTrigger -At $Spec.Time -Once }   # ScheduledTask cmdlets don't expose monthly day-of-month directly without extra XML; we approximate
        'Hourly'  { New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1) }
        default   { $null }
    }
    if (-not $trigger) { Write-ErrorMsg "Unsupported schedule for ScheduledTask"; return $false }
    Register-ScheduledTask -TaskName $Name -Action $action -Trigger $trigger -Description "M365 Manager scheduled health check" -Force | Out-Null
    return $true
}

function Register-CronHealthCheck {
    param([string]$Name, [string]$Command, [hashtable]$Spec)
    $cron = ConvertTo-CronExpression -Spec $Spec
    if (-not $cron) { Write-ErrorMsg "Could not derive cron expression"; return $false }
    # Marker comments so we can find/remove our entries cleanly later
    $marker  = "# m365mgr-$Name"
    $line    = "$cron $Command $marker"
    try {
        $current = (& crontab -l 2>$null) -split "`n"
        $filtered = @($current | Where-Object { $_ -and $_ -notmatch "# m365mgr-$([regex]::Escape($Name))$" })
        $filtered += $line
        ($filtered -join "`n") + "`n" | & crontab -
        return $true
    } catch { Write-ErrorMsg "crontab edit failed: $_"; return $false }
}

# ============================================================
#  Public API
# ============================================================

function New-ScheduledHealthCheck {
    <#
        Register a scheduled health-check. -Script is a repo-relative
        path (e.g. 'health-checks/health-license-usage.ps1').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Script,
        [Parameter(Mandatory)][string]$Schedule,
        [ValidateSet('file','email','teams','all')][string]$Output = 'file',
        [ValidateSet('always','failure','findings')][string]$NotifyOn = 'findings'
    )
    $spec = ConvertTo-ScheduleSpec -Schedule $Schedule
    if (-not $spec) { Write-ErrorMsg "Unrecognized schedule '$Schedule'. Try: 'Daily 09:00', 'Weekly Mon 09:00', 'Monthly 1 09:00', 'Hourly', or 'cron <expr>'."; return $false }

    $repoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $scriptAbs = Join-Path $repoRoot $Script
    if (-not (Test-Path -LiteralPath $scriptAbs)) { Write-ErrorMsg "Script not found: $scriptAbs"; return $false }

    $command = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptAbs`" -NonInteractive -Output $Output -NotifyOn $NotifyOn"

    $registered = $false
    if (Test-IsWindowsHost) {
        $registered = Register-WindowsScheduledHealthCheck -Name $Name -Command $command -Spec $spec
    } else {
        # On macOS / Linux invoke pwsh
        $cmd = "pwsh -NoProfile -File '$scriptAbs' -NonInteractive -Output $Output -NotifyOn $NotifyOn"
        $registered = Register-CronHealthCheck -Name $Name -Command $cmd -Spec $spec
    }
    if (-not $registered) { return $false }

    $state = Read-SchedulerState
    $state += [PSCustomObject]@{
        name      = $Name
        script    = $Script
        schedule  = $Schedule
        spec      = $spec
        output    = $Output
        notifyOn  = $NotifyOn
        addedAt   = (Get-Date).ToUniversalTime().ToString('o')
        platform  = if (Test-IsWindowsHost) { 'WindowsScheduledTask' } else { 'cron' }
    }
    Write-SchedulerState -Records $state
    Write-Success "Scheduled '$Name' ($Schedule) on $(if (Test-IsWindowsHost) {'Task Scheduler'} else {'cron'})"
    return $true
}

function Get-ScheduledHealthChecks {
    <#
        Reads the local state index. On Windows we also enrich
        with Get-ScheduledTask state if available.
    #>
    $state = Read-SchedulerState
    if (-not $state -or $state.Count -eq 0) { return @() }
    if (Test-IsWindowsHost -and (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue)) {
        foreach ($s in $state) {
            try {
                $t = Get-ScheduledTask -TaskName $s.name -ErrorAction Stop
                $info = Get-ScheduledTaskInfo -TaskName $s.name -ErrorAction SilentlyContinue
                $s | Add-Member -NotePropertyName LastRunTime  -NotePropertyValue $info.LastRunTime  -Force
                $s | Add-Member -NotePropertyName NextRunTime  -NotePropertyValue $info.NextRunTime  -Force
                $s | Add-Member -NotePropertyName LastResult   -NotePropertyValue $info.LastTaskResult -Force
            } catch {}
        }
    }
    return $state
}

function Remove-ScheduledHealthCheck {
    param([Parameter(Mandatory)][string]$Name)
    if (Test-IsWindowsHost -and (Get-Command Unregister-ScheduledTask -ErrorAction SilentlyContinue)) {
        try { Unregister-ScheduledTask -TaskName $Name -Confirm:$false -ErrorAction Stop } catch { Write-Warn "Task-Scheduler removal failed: $_" }
    } else {
        try {
            $current = (& crontab -l 2>$null) -split "`n"
            $filtered = @($current | Where-Object { $_ -and $_ -notmatch "# m365mgr-$([regex]::Escape($Name))$" })
            ($filtered -join "`n") + "`n" | & crontab -
        } catch { Write-Warn "crontab removal failed: $_" }
    }
    $state = Read-SchedulerState
    $remaining = @($state | Where-Object { $_.name -ne $Name })
    Write-SchedulerState -Records $remaining
    Write-Success "Removed '$Name'."
}

function Test-ScheduledHealthCheck {
    <#
        Run a registered check now in-process and return its
        structured result JSON (if the script emits one). The
        script is run with -NonInteractive so it never blocks.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $state = Read-SchedulerState
    $rec = $state | Where-Object { $_.name -eq $Name } | Select-Object -First 1
    if (-not $rec) { Write-ErrorMsg "No scheduled check named '$Name'."; return $null }
    $repoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $scriptAbs = Join-Path $repoRoot $rec.script
    if (-not (Test-Path -LiteralPath $scriptAbs)) { Write-ErrorMsg "Script missing: $scriptAbs"; return $null }
    Write-InfoMsg "Running '$Name' in-process (non-interactive)..."
    & $scriptAbs -NonInteractive -Output 'file' -NotifyOn 'findings'
}

# ============================================================
#  Result viewer
# ============================================================

function Get-HealthResults {
    <#
        List health-result-*.json files in the audit dir and
        parse the most recent per check. Returns one row per
        check with its latest status + findings count.
    #>
    $dir = Get-AuditLogDirectory
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $files = @(Get-ChildItem -LiteralPath $dir -Filter 'health-result-*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) { return @() }
    $byName = @{}
    foreach ($f in $files) {
        try { $j = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json } catch { continue }
        if (-not $j.checkName) { continue }
        if (-not $byName.ContainsKey($j.checkName)) {
            $byName[$j.checkName] = [PSCustomObject]@{
                CheckName       = $j.checkName
                LastRunUtc      = $j.completedAtUtc
                Status          = $j.status
                Findings        = $j.findingCount
                File            = $f.FullName
            }
        }
    }
    return @($byName.Values | Sort-Object LastRunUtc -Descending)
}

# ============================================================
#  Menu
# ============================================================

function Start-SchedulerMenu {
    while ($true) {
        $sel = Show-Menu -Title "Scheduled Health Checks" -Options @(
            "List registered checks",
            "Register a new check...",
            "Run a check once now (in-process)...",
            "Remove a check...",
            "View latest results",
            "One-time: register scheduler credential (cert / secret)"
        ) -BackLabel "Back"
        switch ($sel) {
            0 { Get-ScheduledHealthChecks | Format-Table -AutoSize; Pause-ForUser }
            1 {
                $n = Read-UserInput "Friendly name (no spaces)"; if (-not $n) { continue }
                $sc = Read-UserInput "Script path (repo-relative, e.g. health-checks/health-license-usage.ps1)"
                if (-not $sc) { continue }
                $sched = Read-UserInput "Schedule (e.g. 'Daily 09:00', 'Weekly Mon 09:00', 'cron 0 9 * * *')"
                $oSel = Show-Menu -Title "Output channel" -Options @("file","email","teams","all") -BackLabel "Cancel"
                if ($oSel -eq -1) { continue }
                $output = @('file','email','teams','all')[$oSel]
                $nSel = Show-Menu -Title "Notify on" -Options @("always","failure","findings") -BackLabel "Cancel"
                if ($nSel -eq -1) { continue }
                $notifyOn = @('always','failure','findings')[$nSel]
                New-ScheduledHealthCheck -Name $n -Script $sc -Schedule $sched -Output $output -NotifyOn $notifyOn | Out-Null
                Pause-ForUser
            }
            2 { $n = Read-UserInput "Check name"; if ($n) { Test-ScheduledHealthCheck -Name $n }; Pause-ForUser }
            3 { $n = Read-UserInput "Check name to remove"; if ($n -and (Confirm-Action "Remove '$n'?")) { Remove-ScheduledHealthCheck -Name $n }; Pause-ForUser }
            4 { Get-HealthResults | Format-Table -AutoSize; Pause-ForUser }
            5 { Register-SchedulerCredential | Out-Null; Pause-ForUser }
            -1 { return }
        }
    }
}
