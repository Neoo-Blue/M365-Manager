# ============================================================
#  AICostTracker.ps1 -- token usage + USD cost tracking
#
#  Maintains:
#    - $script:AICostState : in-memory session running totals
#    - <stateDir>\ai-cost\events-YYYY-MM.jsonl  : per-call event log
#    - <stateDir>\ai-cost\monthly.json          : rollup keyed by tenant
#  Each Invoke-AIChat / Invoke-AIChatToolingTurn call lands one
#  event line and updates the session/day/month running totals.
#  Budget alerts fire once per threshold cross per month.
# ============================================================

if ($null -eq (Get-Variable -Name AICostState -Scope Script -ErrorAction SilentlyContinue)) {
    $script:AICostState = @{
        SessionStartedUtc = (Get-Date).ToUniversalTime().ToString("o")
        SessionUsd        = 0.0
        SessionInTokens   = 0
        SessionOutTokens  = 0
        SessionCalls      = 0
        DayUsd            = $null   # lazy-loaded
        MonthUsd          = $null
        AlertedThreshold  = @{}     # month -> set of pct levels already warned
    }
}
if ($null -eq (Get-Variable -Name AIPriceTable -Scope Script -ErrorAction SilentlyContinue)) {
    $script:AIPriceTable = $null
}

function Get-AICostDir {
    $base = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'M365Manager' } else { Join-Path $HOME '.m365manager' }
    $p = Join-Path $base 'ai-cost'
    if (-not (Test-Path $p)) {
        New-Item -Path $p -ItemType Directory -Force | Out-Null
        if ($IsLinux -or $IsMacOS) { try { & chmod 700 $p 2>$null } catch {} }
    }
    return $p
}

function Get-AIPriceTable {
    <#
        Load templates/ai-prices.json once and cache. Returns the
        full provider -> model -> { input; output } hashtable.
    #>
    [CmdletBinding()]
    param([switch]$Reload)
    if ($script:AIPriceTable -and -not $Reload) { return $script:AIPriceTable }
    $root = if ($PSScriptRoot) { $PSScriptRoot } elseif ($script:ScriptRoot) { $script:ScriptRoot } else { (Get-Location).Path }
    $p = Join-Path $root 'templates/ai-prices.json'
    if (-not (Test-Path $p)) {
        Write-Warn "ai-prices.json not found at $p -- cost tracking will record \$0 for every call."
        $script:AIPriceTable = @{}
        return $script:AIPriceTable
    }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        $t = @{}
        foreach ($prov in $raw.PSObject.Properties) {
            if ($prov.Name -like '_*') { continue }
            $entry = @{}
            foreach ($m in $prov.Value.PSObject.Properties) {
                $entry[[string]$m.Name] = @{
                    input  = [double]$m.Value.input
                    output = [double]$m.Value.output
                }
            }
            $t[[string]$prov.Name] = $entry
        }
        $script:AIPriceTable = $t
        return $script:AIPriceTable
    } catch {
        Write-Warn "Failed to parse ai-prices.json: $($_.Exception.Message)"
        $script:AIPriceTable = @{}
        return $script:AIPriceTable
    }
}

function Get-AIModelPrice {
    <#
        Return @{ Input; Output; Source = 'exact'|'family'|'unknown' }
        for the given provider+model. Family matches scan keys that
        end in '*' against the model name's prefix; 'claude-*' is the
        last-resort fallback.
    #>
    param(
        [Parameter(Mandatory)][string]$Provider,
        [Parameter(Mandatory)][string]$Model
    )
    $pt = Get-AIPriceTable
    if (-not $pt.ContainsKey($Provider)) { return @{ Input = 0.0; Output = 0.0; Source = 'unknown' } }
    $entry = $pt[$Provider]
    if ($entry.ContainsKey($Model)) {
        return @{ Input = [double]$entry[$Model].input; Output = [double]$entry[$Model].output; Source = 'exact' }
    }
    # Longest-prefix family match first
    $families = @($entry.Keys | Where-Object { $_ -like '*`**' } | Sort-Object { -$_.Length })
    foreach ($k in $families) {
        $prefix = $k.TrimEnd('*')
        if ($Model -like ($prefix + '*')) {
            return @{ Input = [double]$entry[$k].input; Output = [double]$entry[$k].output; Source = 'family' }
        }
    }
    return @{ Input = 0.0; Output = 0.0; Source = 'unknown' }
}

function Get-AICostMonthlyRollup {
    <#
        Return the on-disk monthly rollup hashtable (creates a blank
        one if the file is missing). Schema:
            { "YYYY-MM": { totalUsd; tenants: { "tenantId": usd, ... }; alerted: [pct,...] } }
    #>
    $p = Join-Path (Get-AICostDir) 'monthly.json'
    if (-not (Test-Path $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        $h = @{}
        foreach ($month in $raw.PSObject.Properties) {
            $m = @{ totalUsd = [double]$month.Value.totalUsd; tenants = @{}; alerted = @() }
            if ($month.Value.tenants) {
                foreach ($t in $month.Value.tenants.PSObject.Properties) { $m.tenants[[string]$t.Name] = [double]$t.Value }
            }
            if ($month.Value.alerted) { $m.alerted = @($month.Value.alerted) }
            $h[[string]$month.Name] = $m
        }
        return $h
    } catch { return @{} }
}

function Save-AICostMonthlyRollup {
    param([Parameter(Mandatory)][hashtable]$Rollup)
    $p = Join-Path (Get-AICostDir) 'monthly.json'
    $json = $Rollup | ConvertTo-Json -Depth 8
    Set-Content -LiteralPath $p -Value $json -Encoding UTF8 -Force
}

function Add-AICostEvent {
    <#
        Record one billable call. Returns @{
            Cost; InputTokens; OutputTokens; PriceSource;
            CumulativeSession; CumulativeMonth; AlertFired
        }. Side-effects:
          - Appends one JSONL line to events-YYYY-MM.jsonl
          - Updates monthly rollup
          - Updates $script:AICostState
          - Fires a budget alert when crossing AlertAtPct
            (default 80) of MonthlyBudgetUsd
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][hashtable]$Usage,
        [string]$Reason = 'chat'   # 'chat' | 'tooling-hop' | 'interpret'
    )
    $provider = [string]$Config.Provider
    $model    = [string]$Config.Model
    $inTok    = if ($Usage.InputTokens)  { [int]$Usage.InputTokens }  else { 0 }
    $outTok   = if ($Usage.OutputTokens) { [int]$Usage.OutputTokens } else { 0 }
    $price    = Get-AIModelPrice -Provider $provider -Model $model
    $cost     = (($inTok / 1000000.0) * $price.Input) + (($outTok / 1000000.0) * $price.Output)

    $now      = (Get-Date).ToUniversalTime()
    $month    = $now.ToString('yyyy-MM')
    $tenantId = 'unknown'
    if ($script:SessionState) {
        if ($script:SessionState.TenantDomain) { $tenantId = $script:SessionState.TenantDomain }
        elseif ($script:SessionState.TenantId) { $tenantId = $script:SessionState.TenantId }
        elseif ($script:SessionState.TenantName) { $tenantId = $script:SessionState.TenantName }
    }

    # Event log line
    $event = [ordered]@{
        ts          = $now.ToString("o")
        provider    = $provider
        model       = $model
        inputTokens = $inTok
        outputTokens= $outTok
        usd         = [math]::Round($cost, 6)
        priceSource = $price.Source
        tenant      = $tenantId
        reason      = $Reason
        session     = $PID
    }
    $evFile = Join-Path (Get-AICostDir) ("events-{0}.jsonl" -f $month)
    try { Add-Content -LiteralPath $evFile -Value ($event | ConvertTo-Json -Depth 4 -Compress) -ErrorAction Stop } catch {}

    # Session running total
    $script:AICostState.SessionUsd       = [double]$script:AICostState.SessionUsd + $cost
    $script:AICostState.SessionInTokens  = [int]$script:AICostState.SessionInTokens + $inTok
    $script:AICostState.SessionOutTokens = [int]$script:AICostState.SessionOutTokens + $outTok
    $script:AICostState.SessionCalls     = [int]$script:AICostState.SessionCalls + 1

    # Monthly rollup
    $roll = Get-AICostMonthlyRollup
    if (-not $roll.ContainsKey($month)) { $roll[$month] = @{ totalUsd = 0.0; tenants = @{}; alerted = @() } }
    $roll[$month].totalUsd  = [double]$roll[$month].totalUsd + $cost
    if (-not $roll[$month].tenants.ContainsKey($tenantId)) { $roll[$month].tenants[$tenantId] = 0.0 }
    $roll[$month].tenants[$tenantId] = [double]$roll[$month].tenants[$tenantId] + $cost

    # Budget alert check
    $alertFired = $null
    # Phase 6: budget + threshold are tenant-overridable. Tenant
    # override wins over the global Config value, env var beats both.
    if (Get-Command Get-EffectiveConfig -ErrorAction SilentlyContinue) {
        $budgetRaw = Get-EffectiveConfig -Key 'AI.MonthlyBudgetUsd' -GlobalConfig $Config
        $budget    = if ($budgetRaw) { [double]$budgetRaw } else { 0.0 }
        $alertPctRaw = Get-EffectiveConfig -Key 'AI.AlertAtPct' -GlobalConfig $Config
        $alertPct  = if ($alertPctRaw) { [int]$alertPctRaw } else { 80 }
    } else {
        $budget   = if ($Config.ContainsKey('MonthlyBudgetUsd')) { [double]$Config.MonthlyBudgetUsd } else { 0.0 }
        $alertPct = if ($Config.ContainsKey('AlertAtPct')) { [int]$Config.AlertAtPct } else { 80 }
    }
    if ($budget -gt 0) {
        $used = [double]$roll[$month].totalUsd
        $pctNow = ($used / $budget) * 100.0
        $crossings = @(50, $alertPct, 100, 150) | Sort-Object -Unique
        foreach ($lvl in $crossings) {
            if ($pctNow -ge $lvl -and -not ($roll[$month].alerted -contains $lvl)) {
                $roll[$month].alerted += $lvl
                $alertFired = @{ Pct = $lvl; UsedUsd = $used; BudgetUsd = $budget; Month = $month }
                if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
                    Write-AuditEntry -EventType 'AIBudgetAlert' -Detail ("Crossed {0}% of monthly AI budget ({1:N2} / {2:N2} USD)" -f $lvl, $used, $budget) -ActionType 'AIBudgetAlert' -Target @{ pct = $lvl; usd = $used; budgetUsd = $budget; month = $month } -Result 'warn' | Out-Null
                }
                break
            }
        }
    }

    Save-AICostMonthlyRollup -Rollup $roll
    $script:AICostState.MonthUsd = [double]$roll[$month].totalUsd

    return @{
        Cost              = $cost
        InputTokens       = $inTok
        OutputTokens      = $outTok
        PriceSource       = $price.Source
        CumulativeSession = [double]$script:AICostState.SessionUsd
        CumulativeMonth   = [double]$roll[$month].totalUsd
        AlertFired        = $alertFired
    }
}

function Get-AICostState { return $script:AICostState }

function Reset-AICostSession {
    $script:AICostState.SessionStartedUtc = (Get-Date).ToUniversalTime().ToString("o")
    $script:AICostState.SessionUsd        = 0.0
    $script:AICostState.SessionInTokens   = 0
    $script:AICostState.SessionOutTokens  = 0
    $script:AICostState.SessionCalls      = 0
}

function Show-AICostFooter {
    <#
        Render a one-line cost footer after a chat turn. Skips when
        the call recorded zero cost (Ollama / unpriced model) so the
        chat stays clean for local-only workflows.
    #>
    param([Parameter(Mandatory)][hashtable]$Result)
    if (-not $Result -or $Result.Cost -le 0.0) { return }
    $mark = if ($Result.PriceSource -eq 'unknown') { ' [price unknown -- 0 USD recorded]' } elseif ($Result.PriceSource -eq 'family') { ' [family-priced]' } else { '' }
    $line = "  cost: in={0} out={1} ${2:F4} | session ${3:F4} | month ${4:F4}{5}" -f `
        $Result.InputTokens, $Result.OutputTokens, $Result.Cost, $Result.CumulativeSession, $Result.CumulativeMonth, $mark
    Write-Host $line -ForegroundColor DarkGray

    if ($Result.AlertFired) {
        $a = $Result.AlertFired
        Write-Warn ("AI BUDGET ALERT: monthly spend {0:N2} USD ({1}% of {2:N2} budget) -- consider /noplan or a cheaper model." -f $a.UsedUsd, $a.Pct, $a.BudgetUsd)
    }
}

function Show-AICostSummary {
    <#
        /cost command output. Renders the current session totals
        plus the current-month rollup and per-tenant breakdown.
    #>
    $s = $script:AICostState
    Write-Host ""
    Write-Host "  AI COST -- CURRENT SESSION" -ForegroundColor $script:Colors.Title
    Write-StatusLine "Started"        $s.SessionStartedUtc 'White'
    Write-StatusLine "Calls"          ("{0}"      -f $s.SessionCalls)      'White'
    Write-StatusLine "Input tokens"   ("{0:N0}"  -f $s.SessionInTokens)   'White'
    Write-StatusLine "Output tokens"  ("{0:N0}"  -f $s.SessionOutTokens)  'White'
    Write-StatusLine "Session USD"    ("{0:N4}" -f $s.SessionUsd)         $(if ($s.SessionUsd -gt 1.0) { 'Yellow' } else { 'Green' })

    $now   = (Get-Date).ToUniversalTime().ToString('yyyy-MM')
    $roll  = Get-AICostMonthlyRollup
    if ($roll.ContainsKey($now)) {
        Write-Host ""
        Write-Host ("  AI COST -- MONTH {0}" -f $now) -ForegroundColor $script:Colors.Title
        Write-StatusLine "Total USD" ("{0:N4}" -f $roll[$now].totalUsd) $(if ($roll[$now].totalUsd -gt 10.0) { 'Yellow' } else { 'Green' })
        $byTenant = $roll[$now].tenants.GetEnumerator() | Sort-Object Value -Descending
        foreach ($t in $byTenant) {
            Write-StatusLine ("  " + $t.Key) ("{0:N4} USD" -f $t.Value) 'White'
        }
    }
    Write-Host ""
}

function Show-AICostHistory {
    <#
        /costs command output. Renders last-6-months totals (and
        last-7-day total scanned from the current month's event log
        if present).
    #>
    [CmdletBinding()]
    param([int]$Days = 7)
    Write-Host ""
    Write-Host "  AI COST -- LAST 6 MONTHS" -ForegroundColor $script:Colors.Title
    $roll = Get-AICostMonthlyRollup
    $months = @($roll.Keys | Sort-Object -Descending | Select-Object -First 6)
    if ($months.Count -eq 0) { Write-Host "  (no cost records yet)" -ForegroundColor DarkGray; Write-Host ""; return }
    foreach ($mk in $months) { Write-StatusLine $mk ("{0:N4} USD" -f $roll[$mk].totalUsd) 'White' }

    # Last-N-day total from current-month event log
    $now    = (Get-Date).ToUniversalTime()
    $cutoff = $now.AddDays(-1 * $Days)
    $evFile = Join-Path (Get-AICostDir) ("events-{0}.jsonl" -f $now.ToString('yyyy-MM'))
    if (Test-Path $evFile) {
        $sum = 0.0; $calls = 0
        foreach ($ln in (Get-Content -LiteralPath $evFile)) {
            try {
                $e = $ln | ConvertFrom-Json -ErrorAction Stop
                $ts = [datetime]::Parse($e.ts).ToUniversalTime()
                if ($ts -ge $cutoff) { $sum += [double]$e.usd; $calls++ }
            } catch {}
        }
        Write-Host ""
        Write-StatusLine ("Last $Days days") ("{0:N4} USD across {1} call(s)" -f $sum, $calls) 'White'
    }
    Write-Host ""
}
