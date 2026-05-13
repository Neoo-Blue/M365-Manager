# ============================================================
#  AIUx.ps1 -- spinner + about + dryrun helpers for the
#  AI chat REPL. Lives separately from AIAssistant.ps1 so the
#  long REPL doesn't grow further.
# ============================================================

$script:AISpinnerFrames = @('|','/','-','\')

function Start-AISpinner {
    <#
        Returns a hashtable token; pass it to Stop-AISpinner.
        Cooperative: caller redraws the cursor at end. Renders
        as "  Mark: <label> X" and rewrites the X frame ~5 Hz.
    #>
    param([string]$Label = 'thinking')
    Write-Host ("  Mark: {0} " -f $Label) -ForegroundColor "Cyan" -NoNewline
    $job = Start-Job -ScriptBlock {
        param($frames)
        $i = 0
        while ($true) {
            $f = $frames[$i % $frames.Count]
            Write-Output $f
            Start-Sleep -Milliseconds 180
            $i++
        }
    } -ArgumentList (,$script:AISpinnerFrames)
    return @{ Job = $job; Label = $Label; Frame = 0 }
}

function Stop-AISpinner {
    param([Parameter(Mandatory)]$Token)
    if ($Token.Job) {
        try { Stop-Job -Job $Token.Job -ErrorAction SilentlyContinue | Out-Null } catch {}
        try { Remove-Job -Job $Token.Job -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    # Wipe the spinner line so the cursor lands at the start of the next row.
    try {
        $y = $Host.UI.RawUI.CursorPosition.Y
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $y
        Write-Host (" " * 80) -NoNewline
        $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, $y
    } catch {}
}

function Invoke-OllamaStream {
    <#
        Real streaming for Ollama. Sends stream=true and reads
        line-by-line JSON chunks off the response stream, writing
        each .message.content delta to Console.Out as it arrives.
        Returns @{ Text; Usage } at the end.

        Anthropic/OpenAI SSE streaming would need an HttpClient
        keep-alive pattern that PS 5.1 doesn't expose cleanly --
        those still use Invoke-RestMethod with a spinner.
    #>
    param(
        [Parameter(Mandatory)][string]$Endpoint,
        [Parameter(Mandatory)][string]$Model,
        [Parameter(Mandatory)][string]$SystemPrompt,
        [Parameter(Mandatory)][array]$SafeMessages,
        [array]$Tools = $null
    )
    $payload = @{
        model    = $Model
        stream   = $true
        messages = @(@{ role='system'; content=$SystemPrompt }) + $SafeMessages
    }
    if ($Tools) { $payload.tools = $Tools }
    $bodyJson = $payload | ConvertTo-Json -Depth 12

    $req = [System.Net.WebRequest]::Create("$Endpoint/api/chat")
    $req.Method      = 'POST'
    $req.ContentType = 'application/json'
    $req.Timeout     = 180000
    $rawBody = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)
    $req.ContentLength = $rawBody.Length
    $rs = $req.GetRequestStream()
    $rs.Write($rawBody, 0, $rawBody.Length); $rs.Close()

    $buffer = New-Object System.Text.StringBuilder
    $usage  = @{ InputTokens = 0; OutputTokens = 0 }
    $toolCalls = @()
    try {
        $resp   = $req.GetResponse()
        $stream = $resp.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream)
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try { $j = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }
            if ($j.message.content) {
                $delta = [string]$j.message.content
                [void]$buffer.Append($delta)
                Write-Host $delta -NoNewline -ForegroundColor White
            }
            if ($j.message.tool_calls) {
                foreach ($tc in $j.message.tool_calls) { $toolCalls += $tc }
            }
            if ($j.done) {
                if ($j.prompt_eval_count) { $usage.InputTokens  = [int]$j.prompt_eval_count }
                if ($j.eval_count)        { $usage.OutputTokens = [int]$j.eval_count }
            }
        }
        $reader.Close(); $resp.Close()
    } catch {
        if ($buffer.Length -eq 0) { throw }
        Write-Host ""; Write-Warn ("Stream cut off: " + $_.Exception.Message)
    }
    Write-Host ""
    return @{ Text = $buffer.ToString(); Usage = $usage; ToolCalls = $toolCalls }
}

function Show-AIAbout {
    <#
        /about -- a compact diagnostic snapshot. Goes in /help banner.
        Renders provider/model, tool capability, planner state, cost
        running totals, current session, audit log path, and the
        ai-tools / templates / chat-sessions / ai-cost dirs so the
        operator can find their data.
    #>
    param([hashtable]$Config)
    Write-Host ""
    Write-Host "  M365 MANAGER -- AI ASSISTANT (Mark)" -ForegroundColor $script:Colors.Title
    if ($Config) {
        Write-StatusLine "Provider"  ("{0} / {1}" -f $Config.Provider, $Config.Model) 'White'
        $endpoint = if ($Config.Endpoint) { $Config.Endpoint } else { '(default)' }
        Write-StatusLine "Endpoint"  $endpoint 'White'
    }
    $toolSupport = $false
    if (Get-Command Test-ProviderToolSupport -ErrorAction SilentlyContinue) {
        $toolSupport = Test-ProviderToolSupport -Config $Config
    }
    Write-StatusLine "Native tools" $(if ($toolSupport) { 'YES' } else { 'NO (RUN: regex fallback)' }) $(if ($toolSupport) { 'Green' } else { 'Yellow' })
    if (Get-Command Get-AIToolCatalog -ErrorAction SilentlyContinue) {
        $cat = Get-AIToolCatalog
        Write-StatusLine "Tool catalog" ("{0} tools (incl. meta)" -f @($cat).Count) 'White'
    }
    if (Get-Command Get-AIPlanMode -ErrorAction SilentlyContinue) {
        Write-StatusLine "Plan mode (next prompt)" (Get-AIPlanMode) 'White'
        Write-StatusLine "Auto-plan threshold"     ([string]$script:AIAutoPlanThreshold) 'White'
    }
    Write-StatusLine "Preview mode" $(if ((Get-Command Get-PreviewMode -ErrorAction SilentlyContinue) -and (Get-PreviewMode)) { 'YES (no tenant changes)' } else { 'NO (changes applied)' }) $(if ((Get-Command Get-PreviewMode -ErrorAction SilentlyContinue) -and (Get-PreviewMode)) { 'Yellow' } else { 'Green' })
    if (Get-Command Get-AICostState -ErrorAction SilentlyContinue) {
        $cs = Get-AICostState
        Write-StatusLine "Session cost USD"  ("{0:N4}" -f [double]$cs.SessionUsd) 'White'
        Write-StatusLine "Session calls"      ([string]$cs.SessionCalls) 'White'
    }
    if (Get-Command Get-AISessionCurrent -ErrorAction SilentlyContinue) {
        $cur = Get-AISessionCurrent
        Write-StatusLine "Saved session" $(if ($cur.Id) { "$($cur.Title) ($($cur.Id))" } else { '(unsaved)' }) 'White'
        Write-StatusLine "Ephemeral"     $(if ($cur.Ephemeral) { 'YES' } else { 'NO' }) $(if ($cur.Ephemeral) { 'Yellow' } else { 'Green' })
    }
    if (Get-Command Get-AIAuditLogPath -ErrorAction SilentlyContinue) {
        Write-StatusLine "Audit log" (Get-AIAuditLogPath) 'DarkGray'
    }
    if (Get-Command Get-AISessionDir -ErrorAction SilentlyContinue) {
        Write-StatusLine "Sessions dir" (Get-AISessionDir) 'DarkGray'
    }
    if (Get-Command Get-AICostDir -ErrorAction SilentlyContinue) {
        Write-StatusLine "Cost dir"     (Get-AICostDir) 'DarkGray'
    }
    Write-Host ""
}

function Build-RejectionToolResult {
    <#
        Operator-declined tool calls deserve a structured reason so
        the AI can adapt rather than blindly retry. Returns a JSON
        string for use as a tool_result content payload.
    #>
    param(
        [Parameter(Mandatory)][string]$ToolName,
        [string]$Note,
        [string]$Reason = 'operator_declined'
    )
    $h = @{
        ok       = $false
        error    = $Reason
        tool     = $ToolName
        guidance = switch ($Reason) {
            'operator_declined' { 'The operator declined this tool call. Ask what they want changed before retrying. Do not propose the same call again.' }
            'preview_blocked'   { 'Preview mode is on. Describe what the tool would do, but do not call mutating tools until the operator switches to LIVE.' }
            default             { 'Adjust your plan in light of the operator note before retrying.' }
        }
    }
    if ($Note) { $h.note = $Note }
    return ($h | ConvertTo-Json -Compress)
}
