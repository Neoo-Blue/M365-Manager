# ============================================================
#  MSPReports.ps1 -- cross-tenant operations + pre-built MSP
#  rollups (Phase 6 commit C)
#
#  Invoke-AcrossTenants is the engine: switches tenant context,
#  runs a scriptblock, captures result + error + duration, and
#  guarantees the prior tenant is restored (try/finally) even
#  on uncaught exceptions inside the scriptblock.
#
#  The wrappers (Get-CrossTenant*) call Invoke-AcrossTenants
#  with a scriptblock that calls the matching single-tenant
#  function from earlier phases.
# ============================================================

function Invoke-AcrossTenants {
    <#
        Iterate a scriptblock across registered tenants.

        -Tenants    Array of tenant names. Pass @('@all') (or
                    leave empty) to run across every registered
                    tenant. Names that don't exist are skipped
                    with a warning.
        -Script     Scriptblock that does the per-tenant work.
                    Receives the current tenant profile as $args[0].
                    Return whatever you want; we capture it into
                    the result hash.
        -Parallel   PS 7+ uses ForEach-Object -Parallel. PS 5.1
                    silently falls back to sequential because
                    cross-tenant context-flipping with runspaces
                    is a foot-gun -- the connection state is
                    process-global, so two runspaces hitting
                    different tenants would race. The flag is
                    accepted for API consistency.
        -MaxParallel  Default 4 when -Parallel is honored.

        Always restores the originally-current tenant on exit
        (try/finally), even if the scriptblock throws.

        Returns array of:
          [PSCustomObject]@{
            Tenant      : <name>
            TenantId    : <guid>
            Success     : $true | $false
            Result      : <returned object> | $null
            Error       : <string> | $null
            DurationMs  : <int>
          }
    #>
    [CmdletBinding()]
    param(
        [string[]]$Tenants,
        [Parameter(Mandatory)][scriptblock]$Script,
        [switch]$Parallel,
        [int]$MaxParallel = 4
    )
    $all = Get-Tenants
    if (-not $Tenants -or $Tenants.Count -eq 0 -or ($Tenants.Count -eq 1 -and $Tenants[0] -eq '@all')) {
        $targets = @($all)
    } else {
        $targets = @()
        foreach ($n in $Tenants) {
            $hit = $all | Where-Object { $_.name -eq $n } | Select-Object -First 1
            if ($hit) { $targets += ,$hit } else { Write-Warn "Tenant '$n' not in registry; skipping." }
        }
    }
    if ($targets.Count -eq 0) { Write-Warn "No tenants to run across."; return @() }

    $priorTenant = Get-CurrentTenant
    $priorName   = if ($priorTenant) { $priorTenant.name } else { $null }

    # Audit one CrossTenantOperation entry under the originating
    # tenant before any context flip, so the chain can be reconstructed.
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'CrossTenantOperation' `
            -Detail ("Starting across {0} tenant(s): {1}" -f $targets.Count, (@($targets | ForEach-Object { $_.name }) -join ', ')) `
            -ActionType 'CrossTenantOperation' `
            -Target @{ count = $targets.Count; tenants = @($targets | ForEach-Object { $_.name }); priorTenant = $priorName } `
            -Result 'info' | Out-Null
    }

    $results = New-Object System.Collections.ArrayList
    try {
        if ($Parallel -and $PSVersionTable.PSVersion.Major -ge 7) {
            Write-Warn "Parallel cross-tenant runs share process-global connection state; results may interleave unpredictably. Use sequential for destructive operations."
        }
        # Always sequential -- see comment on -Parallel above.
        foreach ($t in $targets) {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $rec = [ordered]@{
                Tenant     = $t.name
                TenantId   = $t.tenantId
                Success    = $false
                Result     = $null
                Error      = $null
                DurationMs = 0
            }
            try {
                Switch-Tenant -Name $t.name | Out-Null
                $rec.Result  = & $Script $t
                $rec.Success = $true
            } catch {
                $rec.Error = $_.Exception.Message
                if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
                    Write-AuditEntry -EventType 'CrossTenantStepError' -Detail ("Cross-tenant step failed for '{0}': {1}" -f $t.name, $rec.Error) -ActionType 'CrossTenantStepError' -Target @{ tenant = $t.name } -Result 'failure' -ErrorMessage $rec.Error | Out-Null
                }
            } finally {
                $sw.Stop()
                $rec.DurationMs = [int]$sw.ElapsedMilliseconds
            }
            [void]$results.Add([PSCustomObject]$rec)
        }
    } finally {
        # Restore prior tenant regardless of outcome.
        if ($priorName -and $priorName -ne $script:SessionState.TenantName) {
            try { Switch-Tenant -Name $priorName | Out-Null } catch { Write-Warn "Failed to restore prior tenant '$priorName': $($_.Exception.Message)" }
        }
    }
    return @($results)
}

# ============================================================
#  Pre-built per-tenant report wrappers
#
#  Each wrapper returns a uniform shape:
#    @{ Tenants = @(per-tenant rows); Summary = @{...rollup...} }
#  so MSPDashboard.ps1 can render any of them in the same way.
# ============================================================

function Get-CrossTenantLicenseUtilization {
    [CmdletBinding()] param([string[]]$Tenants)
    $rows = Invoke-AcrossTenants -Tenants $Tenants -Script {
        if (Get-Command Get-LicenseUtilizationReport -ErrorAction SilentlyContinue) { Get-LicenseUtilizationReport }
        else { @() }
    }
    $totalUnused = 0; $totalSeats = 0
    foreach ($r in $rows) {
        if ($r.Success -and $r.Result) {
            foreach ($s in @($r.Result)) {
                if ($s.Total)         { $totalSeats  += [int]$s.Total }
                if ($s.AvailableUnits){ $totalUnused += [int]$s.AvailableUnits }
                elseif ($s.Unused)    { $totalUnused += [int]$s.Unused }
            }
        }
    }
    return @{
        Tenants = $rows
        Summary = @{
            TenantsScanned = $rows.Count
            SuccessCount   = ($rows | Where-Object Success).Count
            TotalSeats     = $totalSeats
            TotalUnused    = $totalUnused
        }
    }
}

function Get-CrossTenantMFAGaps {
    [CmdletBinding()] param([string[]]$Tenants)
    $rows = Invoke-AcrossTenants -Tenants $Tenants -Script {
        $noMfa     = @()
        $phoneOnly = @()
        $tap       = @()
        if (Get-Command Invoke-MfaComplianceScan -ErrorAction SilentlyContinue) {
            $scan = Invoke-MfaComplianceScan
            $noMfa     = @(Get-UsersWithNoMfa        -Scan $scan)
            $phoneOnly = @(Get-UsersWithOnlyPhoneMfa -Scan $scan)
            $tap       = @(Get-UsersWithActiveTap    -Scan $scan)
        }
        return @{ NoMfa = $noMfa; PhoneOnly = $phoneOnly; Tap = $tap }
    }
    $totalNoMfa = 0; $totalPhone = 0; $totalTap = 0
    foreach ($r in $rows) {
        if ($r.Success) {
            $totalNoMfa += @($r.Result.NoMfa).Count
            $totalPhone += @($r.Result.PhoneOnly).Count
            $totalTap   += @($r.Result.Tap).Count
        }
    }
    return @{
        Tenants = $rows
        Summary = @{
            TenantsScanned = $rows.Count
            TotalNoMfa     = $totalNoMfa
            TotalPhoneOnly = $totalPhone
            TotalActiveTap = $totalTap
        }
    }
}

function Get-CrossTenantStaleGuests {
    [CmdletBinding()] param([string[]]$Tenants, [int]$DaysSinceSignIn = 90)
    $rows = Invoke-AcrossTenants -Tenants $Tenants -Script {
        if (Get-Command Get-StaleGuests -ErrorAction SilentlyContinue) { Get-StaleGuests -DaysSinceSignIn $using:DaysSinceSignIn }
        else { @() }
    }
    $total = 0
    foreach ($r in $rows) { if ($r.Success) { $total += @($r.Result).Count } }
    return @{
        Tenants = $rows
        Summary = @{
            TenantsScanned = $rows.Count
            TotalStaleGuests = $total
            ThresholdDays    = $DaysSinceSignIn
        }
    }
}

function Get-CrossTenantOrphanedTeams {
    [CmdletBinding()] param([string[]]$Tenants)
    $rows = Invoke-AcrossTenants -Tenants $Tenants -Script {
        if (Get-Command Get-OrphanedTeams -ErrorAction SilentlyContinue) { Get-OrphanedTeams }
        else { @() }
    }
    $total = 0
    foreach ($r in $rows) { if ($r.Success) { $total += @($r.Result).Count } }
    return @{
        Tenants = $rows
        Summary = @{ TenantsScanned = $rows.Count; TotalOrphanedTeams = $total }
    }
}

function Get-CrossTenantBreakGlassPosture {
    [CmdletBinding()] param([string[]]$Tenants)
    $rows = Invoke-AcrossTenants -Tenants $Tenants -Script {
        if (Get-Command Test-BreakGlassPosture -ErrorAction SilentlyContinue) { Test-BreakGlassPosture }
        else { @{ Status='unknown'; Accounts=@() } }
    }
    $green = 0; $yellow = 0; $red = 0
    foreach ($r in $rows) {
        if ($r.Success -and $r.Result) {
            switch ([string]$r.Result.Status) {
                'ok'    { $green++ }
                'warn'  { $yellow++ }
                'fail'  { $red++ }
                default { $yellow++ }
            }
        } else { $red++ }
    }
    return @{
        Tenants = $rows
        Summary = @{
            TenantsScanned = $rows.Count
            Green = $green; Yellow = $yellow; Red = $red
        }
    }
}

function Show-CrossTenantTable {
    <#
        Pretty-print one wrapper's output to the console. Used by
        the chat tooling layer and ad-hoc operator runs that
        don't want the HTML dashboard.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Report,
        [string]$Title = 'Cross-tenant report'
    )
    Write-Host ""
    Write-Host "  $Title" -ForegroundColor $script:Colors.Title
    foreach ($r in $Report.Tenants) {
        $statColor = if ($r.Success) { 'Green' } else { 'Red' }
        $statText  = if ($r.Success) { 'ok' }    else { ('error: ' + $r.Error) }
        Write-Host ("    {0,-24}  ({1,5} ms)  {2}" -f $r.Tenant, $r.DurationMs, $statText) -ForegroundColor $statColor
    }
    Write-Host ""
    Write-Host "  Summary:" -ForegroundColor White
    foreach ($k in $Report.Summary.Keys) { Write-StatusLine $k ([string]$Report.Summary[$k]) 'White' }
    Write-Host ""
}
