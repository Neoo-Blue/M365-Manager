# ============================================================
#  AIPlanner.ps1 -- multi-step plan approval + execution
#
#  Auto-triggered when the AI's first-turn response contains
#  >= $script:AIAutoPlanThreshold tool_use blocks, OR when the
#  operator types /plan to force plan mode for the next prompt.
#  /noplan forces direct tool calls (skip the planner).
#
#  The AI submits a plan via a special meta-tool 'submit_plan'
#  (defined in ai-tools/_meta.json). The planner intercepts
#  that call, renders the plan, prompts the operator for
#  approval, then either:
#    - approves all (executes every step without per-step
#      prompts; still audits + still respects PREVIEW)
#    - step-by-step (per-step Y/A/skip/abort confirm)
#    - edit (drops the operator into $EDITOR with the plan
#      JSON; re-enters approval after save)
#    - reject (discards; AI gets a tool_result indicating
#      rejection and can re-plan).
#  Plan execution writes ActionType="AIPlan" at start, per-
#  step entries, and a PlanResult summary at the end.
# ============================================================

if ($null -eq (Get-Variable -Name AIAutoPlanThreshold -Scope Script -ErrorAction SilentlyContinue)) {
    $script:AIAutoPlanThreshold = 3
}
if ($null -eq (Get-Variable -Name AIPlanModeNext -Scope Script -ErrorAction SilentlyContinue)) {
    # 'auto' | 'force' | 'skip'  -- set by /plan / /noplan chat commands
    $script:AIPlanModeNext = 'auto'
}

function Set-AIPlanMode {
    param([ValidateSet('auto','force','skip')][string]$Mode)
    $script:AIPlanModeNext = $Mode
}

function Get-AIPlanMode { return $script:AIPlanModeNext }

# ============================================================
#  Plan parsing / validation
# ============================================================

function ConvertTo-PlanHashtable {
    <#
        Plans arrive as PSCustomObject from ConvertFrom-Json (the
        AI's submit_plan tool input). Normalize to a hashtable
        + array of step hashtables so the executor's iteration
        is consistent.
    #>
    param($PlanInput)
    if (-not $PlanInput) { return $null }
    $h = @{}
    foreach ($p in $PlanInput.PSObject.Properties) { $h[$p.Name] = $p.Value }
    $steps = @()
    foreach ($s in @($h.steps)) {
        $sh = @{}
        foreach ($p in $s.PSObject.Properties) { $sh[$p.Name] = $p.Value }
        if ($sh.params -and $sh.params -isnot [hashtable]) {
            $ph = @{}
            foreach ($p in $sh.params.PSObject.Properties) { $ph[$p.Name] = $p.Value }
            $sh.params = $ph
        }
        if (-not $sh.dependsOn) { $sh.dependsOn = @() }
        $steps += $sh
    }
    $h.steps = $steps
    return $h
}

function Test-PlanShape {
    <#
        Light validation: each step must have id/description/tool,
        the tool name must exist in the catalog, every dependsOn
        id must reference an earlier step, and the dependency
        graph must be acyclic. Returns @{ Valid; Errors }.
    #>
    param([Parameter(Mandatory)][hashtable]$Plan)
    $errs = New-Object System.Collections.ArrayList
    if (-not $Plan.steps -or @($Plan.steps).Count -eq 0) {
        [void]$errs.Add("Plan has no steps.")
        return @{ Valid = $false; Errors = @($errs) }
    }
    $idMap = @{}
    foreach ($s in @($Plan.steps)) {
        foreach ($f in 'id','description','tool') {
            if (-not $s.ContainsKey($f) -or [string]::IsNullOrEmpty([string]$s[$f])) {
                [void]$errs.Add("Step is missing required field '$f'."); continue
            }
        }
        if ($s.id) {
            if ($idMap.ContainsKey([int]$s.id)) { [void]$errs.Add("Duplicate step id $($s.id).") }
            else { $idMap[[int]$s.id] = $s }
        }
        if ($s.tool -and (Get-Command Get-AIToolByName -ErrorAction SilentlyContinue)) {
            $def = Get-AIToolByName -Name ([string]$s.tool)
            if (-not $def) { [void]$errs.Add("Step $($s.id) references unknown tool '$($s.tool)'.") }
            elseif ($def.isMeta -and [string]$s.tool -ne 'ask_operator') {
                [void]$errs.Add("Step $($s.id) uses meta tool '$($s.tool)' which is not allowed inside a plan.")
            }
        }
    }
    # Dependency graph check (acyclic + forward-only)
    foreach ($s in @($Plan.steps)) {
        foreach ($d in @($s.dependsOn)) {
            if (-not $idMap.ContainsKey([int]$d)) { [void]$errs.Add("Step $($s.id) depends on missing id $d.") }
            elseif ([int]$d -ge [int]$s.id)        { [void]$errs.Add("Step $($s.id) depends on $d which is not an earlier step.") }
        }
    }
    return @{ Valid = ($errs.Count -eq 0); Errors = @($errs) }
}

# ============================================================
#  Plan rendering / approval UX
# ============================================================

function Show-AIPlan {
    param([Parameter(Mandatory)][hashtable]$Plan)
    Write-Host ""
    $destCount = 0
    foreach ($s in @($Plan.steps)) { if ($s.destructive) { $destCount++ } }
    Write-Host ("  AI PLAN: " + [string]$Plan.goal) -ForegroundColor $script:Colors.Title
    if ($Plan.estimatedDurationSec) { Write-StatusLine "Estimated duration" ("{0} seconds" -f $Plan.estimatedDurationSec) 'DarkGray' }
    Write-StatusLine "Steps total"      ("{0}" -f @($Plan.steps).Count) 'White'
    Write-StatusLine "Destructive"      ("{0}" -f $destCount) $(if ($destCount -gt 0) { 'Red' } else { 'Green' })
    Write-Host ""
    foreach ($s in @($Plan.steps)) {
        $destMark = if ($s.destructive) { '[DESTRUCTIVE] ' } else { '              ' }
        $col      = if ($s.destructive) { 'Red' } else { 'White' }
        $dep      = if (@($s.dependsOn).Count -gt 0) { " (depends on: " + ((@($s.dependsOn) -join ', ')) + ")" } else { '' }
        Write-Host ("  [{0,2}] {1}{2}{3}" -f $s.id, $destMark, $s.tool, $dep) -ForegroundColor $col
        Write-Host ("       {0}" -f $s.description) -ForegroundColor DarkGray
        if ($s.params) {
            foreach ($k in @($s.params.Keys)) {
                Write-Host ("         {0}: {1}" -f $k, ($s.params[$k])) -ForegroundColor DarkGray
            }
        }
    }
    Write-Host ""
}

function Read-PlanApproval {
    <#
        Returns one of: 'approveAll' | 'stepByStep' | 'edit' |
        'reject'. Honors NonInteractive mode by reading
        $env:M365MGR_PLAN_APPROVAL (preset by an automation
        scenario) or defaulting to 'reject'.
    #>
    if (Get-NonInteractiveMode) {
        $forced = $env:M365MGR_PLAN_APPROVAL
        if ($forced -in 'approveAll','stepByStep','edit','reject') { return $forced }
        return 'reject'
    }
    while ($true) {
        Write-Host "  Plan action: [A]pprove all / [S]tep-by-step / [E]dit / [R]eject" -ForegroundColor $script:Colors.Highlight -NoNewline
        Write-Host ": " -NoNewline
        $a = Read-Host
        switch -Regex ($a) {
            '^[Aa]'  { return 'approveAll' }
            '^[Ss]'  { return 'stepByStep' }
            '^[Ee]'  { return 'edit' }
            '^[Rr]'  { return 'reject' }
            default  { Write-Warn "Unrecognized response. Try one of A / S / E / R." }
        }
    }
}

function Edit-AIPlan {
    <#
        Drops the plan JSON into a temp file, opens $EDITOR
        (or notepad on Windows / nano on POSIX as a fallback),
        re-parses on save. Returns the edited plan hashtable or
        $null if parsing failed.
    #>
    param([Parameter(Mandatory)][hashtable]$Plan)
    $tmp = [IO.Path]::GetTempFileName() + '.json'
    try {
        ($Plan | ConvertTo-Json -Depth 12) | Set-Content -LiteralPath $tmp -Encoding UTF8 -Force
        $editor = $env:EDITOR
        if (-not $editor) { $editor = if ($env:LOCALAPPDATA) { 'notepad' } else { 'nano' } }
        Write-InfoMsg "Opening $editor on $tmp ..."
        & $editor $tmp
        try {
            $raw = Get-Content -LiteralPath $tmp -Raw | ConvertFrom-Json -ErrorAction Stop
            return (ConvertTo-PlanHashtable -PlanInput $raw)
        } catch {
            Write-ErrorMsg "Edited plan won't parse as JSON: $($_.Exception.Message)"
            return $null
        }
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}

# ============================================================
#  Executor
# ============================================================

function Get-TopologicalStepOrder {
    <#
        Returns the steps in dependency-respecting order. Plans
        with cycles fail Test-PlanShape upstream so this can
        assume acyclicity.
    #>
    param([Parameter(Mandatory)][array]$Steps)
    $byId = @{}
    foreach ($s in $Steps) { $byId[[int]$s.id] = $s }
    $emitted = @{}
    $result  = New-Object System.Collections.ArrayList
    while ($emitted.Count -lt $byId.Count) {
        $progress = $false
        foreach ($s in ($Steps | Sort-Object id)) {
            if ($emitted.ContainsKey([int]$s.id)) { continue }
            $deps = @($s.dependsOn)
            $allDepsDone = $true
            foreach ($d in $deps) { if (-not $emitted.ContainsKey([int]$d)) { $allDepsDone = $false; break } }
            if ($allDepsDone) {
                [void]$result.Add($s)
                $emitted[[int]$s.id] = $true
                $progress = $true
            }
        }
        if (-not $progress) { break }   # shouldn't happen on acyclic input
    }
    return @($result)
}

function Invoke-AIPlan {
    <#
        Execute an approved plan. Returns @{
            Goal; StepCount; Executed; Succeeded; Failed;
            Skipped; StepResults
        } and writes:
          - one AIPlan audit entry at start
          - one AIToolCall entry per step (via Invoke-AIToolCall)
          - one PlanResult summary entry at the end
        Mode controls per-step prompting:
          approveAll -> no per-step prompt
          stepByStep -> Y/A/skip/abort per step
        FailureMode (from plan, default 'stop'): stop | ask
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Plan,
        [Parameter(Mandatory)][ValidateSet('approveAll','stepByStep')][string]$Mode,
        [hashtable]$Config
    )

    $entryId = New-AuditEntryId
    Write-AuditEntry -EventType 'AIPlan' -Detail ([string]$Plan.goal) -ActionType 'AIPlan' -Target @{ stepCount = @($Plan.steps).Count } -Result 'info' -EntryId $entryId | Out-Null

    $ordered = Get-TopologicalStepOrder -Steps @($Plan.steps)
    $stepResults = New-Object System.Collections.ArrayList
    $succeeded = 0; $failed = 0; $skipped = 0
    $runAll = ($Mode -eq 'approveAll')

    foreach ($s in $ordered) {
        $destMark = if ($s.destructive) { '[DESTRUCTIVE] ' } else { '' }
        Write-Host ""
        Write-Host ("  > [{0}/{1}] {2}{3} -- {4}" -f $s.id, @($ordered).Count, $destMark, $s.tool, $s.description) -ForegroundColor $(if ($s.destructive) { 'Red' } else { 'Cyan' })

        $action = if ($runAll) { 'y' } else {
            Write-Host "  [Y]es  [A]ll  [S]kip  [Q]uit" -ForegroundColor $script:Colors.Highlight -NoNewline
            Write-Host ": " -NoNewline
            Read-Host
        }
        if ($action -match '^[Aa]') { $runAll = $true; $action = 'y' }
        if ($action -match '^[Qq]') {
            [void]$stepResults.Add(@{ id = $s.id; tool = $s.tool; status = 'aborted' })
            break
        }
        if ($action -match '^[Ss]') {
            [void]$stepResults.Add(@{ id = $s.id; tool = $s.tool; status = 'skipped' })
            $skipped++; continue
        }
        if ($action -notmatch '^[Yy]') {
            [void]$stepResults.Add(@{ id = $s.id; tool = $s.tool; status = 'skipped' })
            $skipped++; continue
        }

        $paramsHt = @{}
        if ($s.params) { foreach ($k in @($s.params.Keys)) { $paramsHt[$k] = $s.params[$k] } }
        $out = Invoke-AIToolCall -ToolName ([string]$s.tool) -Params $paramsHt -Config $Config
        if ($out.ok) {
            $succeeded++
            [void]$stepResults.Add(@{ id = $s.id; tool = $s.tool; status = 'success'; result = $out.result })
            Write-Host "    ok" -ForegroundColor Green
        } else {
            $failed++
            [void]$stepResults.Add(@{ id = $s.id; tool = $s.tool; status = 'failed'; error = $out.error; details = $out.details })
            Write-Host ("    failed: {0} :: {1}" -f $out.error, $out.details) -ForegroundColor Red
            $failureMode = if ($Plan.failureMode) { [string]$Plan.failureMode } else { 'stop' }
            if ($failureMode -eq 'stop') {
                Write-Warn "failureMode=stop, aborting remaining steps."
                break
            }
            # failureMode=ask -> prompt to abort / continue / revise
            $next = if (Get-NonInteractiveMode) { 'continue' } else {
                Write-Host "  [C]ontinue / [A]bort / [R]evise" -ForegroundColor $script:Colors.Highlight -NoNewline; Write-Host ": " -NoNewline; Read-Host
            }
            if ($next -match '^[Aa]') { break }
            if ($next -match '^[Rr]') {
                Write-InfoMsg "Revise loop signalled -- caller will request a new plan."
                # Caller handles re-plan via the special return-state
                return @{
                    Goal         = $Plan.goal
                    StepCount    = @($ordered).Count
                    Executed     = ($succeeded + $failed + $skipped)
                    Succeeded    = $succeeded
                    Failed       = $failed
                    Skipped      = $skipped
                    StepResults  = @($stepResults)
                    ReviseRequested = $true
                }
            }
        }
    }

    $summary = @{
        Goal        = [string]$Plan.goal
        StepCount   = @($ordered).Count
        Executed    = ($succeeded + $failed + $skipped)
        Succeeded   = $succeeded
        Failed      = $failed
        Skipped     = $skipped
        StepResults = @($stepResults)
    }
    Write-AuditEntry -EventType 'AIPlanResult' -Detail ("Plan '{0}' :: {1} ok, {2} failed, {3} skipped" -f $Plan.goal, $succeeded, $failed, $skipped) -ActionType 'AIPlanResult' -Target $summary -Result 'info' | Out-Null

    Write-Host ""
    Write-Host "  Plan summary:" -ForegroundColor White
    Write-StatusLine "Succeeded" "$succeeded" 'Green'
    Write-StatusLine "Failed"    "$failed"    $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    Write-StatusLine "Skipped"   "$skipped"   'DarkGray'
    Write-Host ""

    return $summary
}

# ============================================================
#  Orchestration -- intercept submit_plan tool_use
# ============================================================

function Invoke-AIPlanApprovalFlow {
    <#
        Driven by Start-AIAssistant when the AI submits a plan
        (submit_plan meta tool). Reads operator approval,
        optionally edits, executes, then returns:
            @{
                Status    : 'approveAll' | 'stepByStep' | 'edit' | 'rejected' | 'revised'
                PlanFinal : the plan that was ultimately executed (or rejected)
                Result    : Invoke-AIPlan result hashtable (or $null on reject)
            }
        Caller stuffs the relevant tool_result content back to
        the model so it knows what happened.
    #>
    param(
        [Parameter(Mandatory)]$PlanInput,
        [hashtable]$Config
    )
    $plan = ConvertTo-PlanHashtable -PlanInput $PlanInput
    if (-not $plan) { return @{ Status = 'rejected'; Reason = 'parse_failed' } }

    while ($true) {
        $check = Test-PlanShape -Plan $plan
        if (-not $check.Valid) {
            Write-Warn "Plan validation issues:"
            foreach ($e in $check.Errors) { Write-Warn "  - $e" }
            return @{ Status = 'rejected'; Reason = 'validation_failed'; ValidationErrors = $check.Errors; PlanFinal = $plan }
        }
        Show-AIPlan -Plan $plan
        $decision = Read-PlanApproval
        switch ($decision) {
            'reject' {
                Write-Warn "Plan rejected."
                return @{ Status = 'rejected'; PlanFinal = $plan }
            }
            'edit' {
                $edited = Edit-AIPlan -Plan $plan
                if (-not $edited) { Write-Warn "Edit produced invalid JSON; keeping the original."; continue }
                $plan = $edited
                continue
            }
            default {
                $r = Invoke-AIPlan -Plan $plan -Mode $decision -Config $Config
                $status = if ($r.ReviseRequested) { 'revised' } else { $decision }
                return @{ Status = $status; PlanFinal = $plan; Result = $r }
            }
        }
    }
}
