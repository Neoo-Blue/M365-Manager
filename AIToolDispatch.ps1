# ============================================================
#  AIToolDispatch.ps1 -- native tool calling for AI providers
#
#  Replaces the regex `RUN:` extraction path in AIAssistant.ps1
#  with provider-native tool calling (Anthropic, OpenAI / Azure
#  OpenAI, Ollama). The old regex path stays in AIAssistant.ps1
#  for old / small Ollama models that lack tool support; it
#  emits a deprecation warning when it fires.
#
#  The tool registry under ai-tools/ is the PRIMARY allow-list
#  -- the AI can only invoke names listed in those JSON files.
#  The AST allow-list from Phase 0.5 Commit 5 stays as defense
#  in depth for the text-mode fallback.
# ============================================================

$script:AIToolCatalog = $null
$script:AIToolingDeprecationWarned = $false

# ============================================================
#  Catalog loading
# ============================================================

function Get-AIToolsDirectory {
    if ($PSScriptRoot) { return Join-Path $PSScriptRoot 'ai-tools' }
    if ($env:M365ADMIN_ROOT) { return Join-Path $env:M365ADMIN_ROOT 'ai-tools' }
    return Join-Path (Get-Location).Path 'ai-tools'
}

function Get-AIToolCatalog {
    <#
        Returns the merged tool catalog (cached). Pass -Reload
        to re-scan the ai-tools/ directory.
    #>
    [CmdletBinding()]
    param([switch]$Reload)
    if ($script:AIToolCatalog -and -not $Reload) { return $script:AIToolCatalog }
    $dir = Get-AIToolsDirectory
    if (-not (Test-Path -LiteralPath $dir)) { $script:AIToolCatalog = @(); return @() }
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)
    $cat = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($f in $files) {
        try {
            $items = @(Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json)
            foreach ($it in $items) {
                if (-not $it.name) { Write-Warn "Tool entry in $($f.Name) missing 'name'; skipped."; continue }
                if ($seen.ContainsKey([string]$it.name)) {
                    Write-Warn "Duplicate tool '$($it.name)' in $($f.Name) -- first definition wins."
                    continue
                }
                $seen[[string]$it.name] = $true
                [void]$cat.Add($it)
            }
        } catch { Write-Warn "Could not parse $($f.Name): $($_.Exception.Message)" }
    }
    $script:AIToolCatalog = @($cat)
    return $script:AIToolCatalog
}

function Get-AIToolByName {
    param([Parameter(Mandatory)][string]$Name)
    $cat = Get-AIToolCatalog
    return @($cat | Where-Object { $_.name -eq $Name })[0]
}

function Show-AIToolCatalog {
    $cat = Get-AIToolCatalog
    if (-not $cat -or $cat.Count -eq 0) { Write-Warn "No tools loaded. Check ai-tools/ directory."; return }
    Write-Host ""
    Write-Host ("  {0,3} tool(s) loaded from ai-tools/" -f $cat.Count) -ForegroundColor DarkGray
    Write-Host ""
    $byCategory = @{}
    foreach ($t in $cat) {
        $cat2 = if ($t.isMeta) { '_meta' } else { 'tools' }
        if (-not $byCategory.ContainsKey($cat2)) { $byCategory[$cat2] = @() }
        $byCategory[$cat2] += $t
    }
    foreach ($k in $byCategory.Keys | Sort-Object) {
        foreach ($t in $byCategory[$k] | Sort-Object name) {
            $dest = if ($t.destructive) { '[DESTRUCTIVE] ' } else { '              ' }
            $col  = if ($t.destructive) { 'Red' } else { 'White' }
            Write-Host ("    {0}{1}" -f $dest, $t.name) -ForegroundColor $col
            if ($t.description) {
                Write-Host ("                {0}" -f $t.description) -ForegroundColor DarkGray
            }
        }
    }
    Write-Host ""
}

# ============================================================
#  Schema validation
# ============================================================

function Test-AIToolInput {
    <#
        Validate $InputHash against the tool's JSON-schema
        parameters object. Returns @{ Valid; Errors }. Checks
        required fields and basic type / enum constraints.
    #>
    param(
        [Parameter(Mandatory)]$ToolDef,
        [hashtable]$InputHash
    )
    $errs = New-Object System.Collections.ArrayList
    if (-not $ToolDef.parameters) { return @{ Valid = $true; Errors = @() } }

    $required = @()
    if ($ToolDef.parameters.required) { $required = @($ToolDef.parameters.required) }
    foreach ($r in $required) {
        if (-not $InputHash.ContainsKey($r) -or [string]::IsNullOrEmpty([string]$InputHash[$r])) {
            [void]$errs.Add("Required parameter '$r' is missing.")
        }
    }

    $props = $ToolDef.parameters.properties
    if ($props -and $InputHash) {
        $propNames = $props.PSObject.Properties.Name
        foreach ($k in $InputHash.Keys) {
            if ($propNames -notcontains $k) {
                [void]$errs.Add("Unknown parameter '$k'. Allowed: $($propNames -join ', ')")
                continue
            }
            $p = $props.$k
            $v = $InputHash[$k]
            if ($p.type -eq 'integer') {
                $n = 0
                if (-not [int]::TryParse([string]$v, [ref]$n)) { [void]$errs.Add("Parameter '$k' must be integer, got '$v'.") }
            }
            elseif ($p.type -eq 'boolean') {
                if ($v -isnot [bool] -and "$v" -notmatch '^(?i:true|false|1|0|yes|no)$') {
                    [void]$errs.Add("Parameter '$k' must be boolean.")
                }
            }
            elseif ($p.enum) {
                if ("$v" -notin $p.enum) { [void]$errs.Add("Parameter '$k' must be one of: $($p.enum -join ', ')") }
            }
        }
    }
    return @{ Valid = ($errs.Count -eq 0); Errors = @($errs) }
}

# ============================================================
#  Tool invocation -- map tool name to actual PowerShell call
# ============================================================

function Invoke-AIToolImpl {
    <#
        Given a tool name + validated params hashtable, invoke
        the corresponding PowerShell cmdlet / module function.
        Returns the cmdlet's output (truncated for AI context
        elsewhere). All Set-* / Remove-* / Add-* calls go via
        Invoke-Action so they audit + respect PREVIEW.
    #>
    param(
        [Parameter(Mandatory)][string]$ToolName,
        [Parameter(Mandatory)]$ToolDef,
        [hashtable]$Params = @{}
    )

    # Build a splat of all properties whose value isn't null
    $splat = @{}
    foreach ($k in $Params.Keys) {
        if ($null -ne $Params[$k] -and "$($Params[$k])" -ne '') { $splat[$k] = $Params[$k] }
    }

    switch -Regex ($ToolName) {

        # ---- License wrappers (unwrap to Set-MgUserLicense) ----
        '^Set-MgUserLicense-Add$' {
            return (Invoke-Action `
                -Description ("AI: Assign license {0} to {1}" -f $splat.SkuId, $splat.UserId) `
                -ActionType 'AssignLicense' `
                -Target @{ userId = $splat.UserId; skuId = $splat.SkuId } `
                -ReverseType 'RemoveLicense' `
                -ReverseDescription ("Remove license {0} from {1}" -f $splat.SkuId, $splat.UserId) `
                -Action {
                    Set-MgUserLicense -UserId $splat.UserId -AddLicenses @(@{ SkuId = $splat.SkuId }) -RemoveLicenses @() -ErrorAction Stop
                })
        }
        '^Set-MgUserLicense-Remove$' {
            return (Invoke-Action `
                -Description ("AI: Remove license {0} from {1}" -f $splat.SkuId, $splat.UserId) `
                -ActionType 'RemoveLicense' `
                -Target @{ userId = $splat.UserId; skuId = $splat.SkuId } `
                -ReverseType 'AssignLicense' `
                -ReverseDescription ("Re-assign {0} to {1}" -f $splat.SkuId, $splat.UserId) `
                -Action {
                    Set-MgUserLicense -UserId $splat.UserId -AddLicenses @() -RemoveLicenses @($splat.SkuId) -ErrorAction Stop
                })
        }

        # ---- Mailbox wrappers ----
        '^Set-Mailbox-Type$' {
            return (Invoke-Action `
                -Description ("AI: Convert mailbox {0} to {1}" -f $splat.Identity, $splat.Type) `
                -ActionType 'ConvertMailboxType' `
                -Target @{ identity = $splat.Identity; type = $splat.Type } `
                -NoUndoReason 'Mailbox type changes require manual reversal (re-license + Set-Mailbox).' `
                -Action { Set-Mailbox -Identity $splat.Identity -Type $splat.Type -ErrorAction Stop; $true })
        }
        '^Set-Mailbox-Forwarding$' {
            $fwd = [string]$splat.ForwardingSmtpAddress
            $deliver = if ($splat.ContainsKey('DeliverToMailboxAndForward')) { [bool]$splat.DeliverToMailboxAndForward } else { $true }
            $isClear = ([string]::IsNullOrWhiteSpace($fwd))
            $desc = if ($isClear) { "Clear forwarding on $($splat.Identity)" } else { "Set forwarding $($splat.Identity) -> $fwd (keep copy: $deliver)" }
            return (Invoke-Action `
                -Description ("AI: $desc") `
                -ActionType $(if ($isClear) { 'ClearForwarding' } else { 'SetForwarding' }) `
                -Target @{ identity = $splat.Identity; forwardTo = $fwd; keepCopy = $deliver } `
                -ReverseType $(if ($isClear) { 'SetForwarding' } else { 'ClearForwarding' }) `
                -ReverseDescription ("Inverse of $desc") `
                -Action {
                    if ($isClear) { Set-Mailbox -Identity $splat.Identity -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop }
                    else          { Set-Mailbox -Identity $splat.Identity -ForwardingSmtpAddress ("smtp:" + $fwd) -DeliverToMailboxAndForward $deliver -ErrorAction Stop }
                    $true
                })
        }

        # ---- SharePoint owner wrappers ----
        '^Set-SiteOwner-Add$'    { return (Set-SiteOwner -SiteUrl $splat.SiteUrl -UPN $splat.UPN -Direction Add) }
        '^Set-SiteOwner-Remove$' { return (Set-SiteOwner -SiteUrl $splat.SiteUrl -UPN $splat.UPN -Direction Remove) }

        # ---- Renamed Search-SignIns params (FromDate / ToDate -> -From / -To) ----
        '^Search-SignIns$' {
            $args = @{}
            if ($splat.UPN)          { $args.UPN = $splat.UPN }
            if ($splat.FromDate)     { $args.From = [DateTime]$splat.FromDate }
            if ($splat.ToDate)       { $args.To = [DateTime]$splat.ToDate }
            if ($splat.AppName)      { $args.AppName = $splat.AppName }
            if ($splat.IP)           { $args.IP = $splat.IP }
            if ($splat.Country)      { $args.Country = $splat.Country }
            if ($splat.RiskLevel)    { $args.RiskLevel = $splat.RiskLevel }
            if ($splat.OnlyFailures) { $args.OnlyFailures = [bool]$splat.OnlyFailures }
            if ($splat.MaxResults)   { $args.MaxResults = [int]$splat.MaxResults }
            return (Search-SignIns @args)
        }

        # ---- Phase 7 incident-response tools ----
        '^Invoke-CompromisedAccountResponse$' {
            # The function manages Invoke-Action internally per step,
            # so don't wrap again. Pass through splat verbatim.
            return (Invoke-CompromisedAccountResponse @splat)
        }
        '^Get-IncidentTimeline$' {
            return (Get-IncidentTimeline @splat)
        }
        '^Get-IncidentList$' {
            return (Get-IncidentList @splat)
        }
        '^Summarize-AuditEvents$' {
            return (Summarize-AuditEvents @splat)
        }

        # ---- Search-UAL date rename ----
        '^Search-UAL$' {
            $args = @{}
            if ($splat.FromDate)   { $args.From = [DateTime]$splat.FromDate }
            if ($splat.ToDate)     { $args.To   = [DateTime]$splat.ToDate }
            if ($splat.UserId)     { $args.UserId = $splat.UserId }
            if ($splat.Operations) { $args.Operations = @($splat.Operations) }
            if ($splat.IP)         { $args.IP = $splat.IP }
            return (Search-UAL @args)
        }

        # ---- Default: direct splat against same-named cmdlet / module function ----
        default {
            $cmd = Get-Command $ToolName -ErrorAction SilentlyContinue
            if (-not $cmd) { throw "No cmdlet named '$ToolName' is currently loaded." }
            # Pre-merge fix: when the catalog says wrapInInvokeAction=true,
            # wrap the default-branch call in Invoke-Action so destructive
            # SDK / Exchange cmdlets (Remove-MgGroupMember,
            # Revoke-MgUserSignInSession, Set-MailboxAutoReplyConfiguration,
            # Remove-MailboxPermission, Update-MgUser,
            # Remove-DistributionGroupMember, etc.) get a proper EXEC audit
            # line with actionType + noUndoReason rather than just an
            # AIToolCall line. Custom functions that already wrap internally
            # (Remove-Guest, Remove-AllAuthMethods, Revoke-OneDriveAccess,
            # Remove-UserFromTeam, Set-TeamOwnership, New-TemporaryAccessPass)
            # still wrap themselves -- double-wrapping is harmless because
            # Invoke-Action's PREVIEW short-circuit + audit entries
            # idempotently dedupe by entryId.
            if ($ToolDef.wrapInInvokeAction -and (Get-Command Invoke-Action -ErrorAction SilentlyContinue)) {
                $targetHash = @{}
                foreach ($k in $Params.Keys) { $targetHash[$k] = $Params[$k] }
                return (Invoke-Action `
                    -Description ("AI: {0} {1}" -f $ToolName, (($Params | ConvertTo-Json -Compress -Depth 4))) `
                    -ActionType ("AI:" + $ToolName) `
                    -Target $targetHash `
                    -NoUndoReason ("AI-invoked '{0}' has no curated reverse recipe; manual reversal required." -f $ToolName) `
                    -Action {
                        & $ToolName @splat
                        $true
                    })
            }
            return (& $ToolName @splat)
        }
    }
}

# ============================================================
#  Public dispatch entry point
# ============================================================

function Invoke-AIToolCall {
    <#
        Validate + execute a single AI tool_use. Returns a
        hashtable @{ ok=$bool; result=...; error=...; details=...; }
        suitable for stuffing into a tool_result block sent back
        to the model.
    #>
    param(
        [Parameter(Mandatory)][string]$ToolName,
        [hashtable]$Params = @{},
        [hashtable]$Config
    )
    $tool = Get-AIToolByName -Name $ToolName
    if (-not $tool) {
        return @{ ok = $false; error = 'unknown_tool'; details = "No tool '$ToolName' in the catalog. Use the /tools chat command to list available tools." }
    }
    if ($tool.isMeta) {
        return @{ ok = $false; error = 'meta_tool'; details = "Tool '$ToolName' is a meta tool and must be handled by the caller, not Invoke-AIToolCall." }
    }
    $check = Test-AIToolInput -ToolDef $tool -InputHash $Params
    if (-not $check.Valid) {
        # Returning a structured error -- the model receives this
        # as a tool_result and can self-correct on the next turn.
        return @{ ok = $false; error = 'validation_failed'; details = ($check.Errors -join ' ; ') }
    }

    Write-AIAuditEntry -EventType 'AIToolCall' -Detail ("{0} {1}" -f $ToolName, (($Params | ConvertTo-Json -Compress -Depth 5)))

    try {
        $result = Invoke-AIToolImpl -ToolName $ToolName -ToolDef $tool -Params $Params
        return @{ ok = $true; result = $result }
    } catch {
        return @{ ok = $false; error = 'execution_failed'; details = $_.Exception.Message }
    }
}

# ============================================================
#  Provider tool-payload builders
# ============================================================

function Build-AnthropicToolsPayload {
    param([array]$Catalog)
    $tools = New-Object System.Collections.ArrayList
    foreach ($t in $Catalog) {
        [void]$tools.Add(@{
            name         = $t.name
            description  = [string]$t.description
            input_schema = $t.parameters
        })
    }
    return @($tools)
}

function Build-OpenAIToolsPayload {
    param([array]$Catalog)
    $tools = New-Object System.Collections.ArrayList
    foreach ($t in $Catalog) {
        [void]$tools.Add(@{
            type     = 'function'
            function = @{
                name        = $t.name
                description = [string]$t.description
                parameters  = $t.parameters
            }
        })
    }
    return @($tools)
}

# ============================================================
#  Capability probe / cache
# ============================================================

$script:AIToolCapabilityCache = @{}

function Test-ProviderToolSupport {
    <#
        Returns $true if the configured provider+model supports
        native tool calling. Anthropic / OpenAI / Azure: always
        true. Ollama: cached per model, probe once. Custom:
        operator-controlled via the config's ToolSupport flag.
    #>
    param([hashtable]$Config)
    $p = ([string]$Config.Provider).ToLowerInvariant()
    $key = ("{0}::{1}" -f $p, $Config.Model)
    if ($script:AIToolCapabilityCache.ContainsKey($key)) { return $script:AIToolCapabilityCache[$key] }
    $result = switch ($p) {
        'anthropic'   { $true }
        'openai'      { $true }
        'azureopenai' { $true }
        'ollama'      {
            # Probe via small request. If the response shape has
            # tool_calls support indicated, mark as supported. To
            # avoid the cost we look for "tools" in the model
            # listing if available, else assume true for recent
            # models and surface failures at first call.
            $true
        }
        'custom'      { [bool]($Config.ToolSupport) }
        default       { $false }
    }
    $script:AIToolCapabilityCache[$key] = $result
    return $result
}

# ============================================================
#  Tool-aware chat turn -- one round trip to the provider.
#  The REPL drives the conversation loop.
# ============================================================

function Invoke-AIChatToolingTurn {
    <#
        Send one turn to the provider with the full tool catalog
        attached. Returns:
          @{
              Text       = '...'                  # any assistant text content
              ToolUses   = @(@{ id; name; input } ...)  # normalized across providers
              AssistantContent = $rawContentArray # exact assistant content array
                                                  # (Anthropic-shaped, used for follow-up
                                                  #  user message construction)
              StopReason = 'end_turn'|'tool_use'|'max_tokens'|...
              Usage      = @{ InputTokens; OutputTokens }
              Error      = $null|'...'
          }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$Messages,
        [array]$Catalog = $null
    )

    if (-not $Catalog) { $Catalog = Get-AIToolCatalog }
    $provider = [string]$Config.Provider
    $endpoint = [string]$Config.Endpoint
    $model    = [string]$Config.Model
    try { $key = Unprotect-ApiKey -StoredKey ([string]$Config.ApiKey) } catch { return @{ Error = "Decrypt: $_" } }

    # ---- Privacy redaction on outbound payload ----
    $privacy = if ($Config.ContainsKey('Privacy') -and $Config.Privacy -is [hashtable]) { $Config.Privacy } else { @{ ExternalRedaction='Enabled'; ExternalPayloadCapBytes=8192; TrustedProviders=@() } }
    $isExternal = Test-IsExternalProvider -Provider $provider -Endpoint $endpoint -TrustedProviders $privacy.TrustedProviders
    $fullRedact = $isExternal -and ($privacy.ExternalRedaction -eq 'Enabled')
    $secretsOnly = -not $fullRedact

    $safeMessages = @()
    $counts = @{ JWT=0; SECRET=0; THUMB=0; UPN=0; GUID=0; TENANT=0; NAME=0 }
    foreach ($m in $Messages) {
        $content = $m.content
        if ($content -is [string]) {
            $safeMessages += @{ role = $m.role; content = (Convert-ToSafePayload -Text $content -SecretsOnly:$secretsOnly -Counts $counts) }
        } else {
            # array of blocks (tool_use, tool_result, etc.) -- redact embedded strings
            $newBlocks = @()
            foreach ($b in @($content)) {
                if ($b -is [string]) { $newBlocks += (Convert-ToSafePayload -Text $b -SecretsOnly:$secretsOnly -Counts $counts) }
                elseif ($b.PSObject.Properties.Name -contains 'content' -and $b.content -is [string]) {
                    $clone = @{}
                    foreach ($p in $b.PSObject.Properties) { $clone[$p.Name] = $p.Value }
                    $clone.content = Convert-ToSafePayload -Text $b.content -SecretsOnly:$secretsOnly -Counts $counts
                    $newBlocks += $clone
                } else { $newBlocks += $b }
            }
            $safeMessages += @{ role = $m.role; content = $newBlocks }
        }
    }

    $systemPrompt = $script:AISystemPrompt
    if ($fullRedact) { $systemPrompt += $script:PrivacySystemPromptAddendum }
    if ($script:PlanningSystemPromptAddendum) { $systemPrompt += $script:PlanningSystemPromptAddendum }
    $planModeNow = if (Get-Command Get-AIPlanMode -ErrorAction SilentlyContinue) { Get-AIPlanMode } else { 'auto' }
    if     ($planModeNow -eq 'force' -and $script:PlanForceSystemPromptAddendum) { $systemPrompt += $script:PlanForceSystemPromptAddendum }
    elseif ($planModeNow -eq 'skip'  -and $script:NoPlanSystemPromptAddendum)    { $systemPrompt += $script:NoPlanSystemPromptAddendum }

    $body = $null; $headers = $null; $uri = $endpoint; $usage = $null
    $rawAssistantContent = @()
    $toolUses = New-Object System.Collections.ArrayList
    $stopReason = 'end_turn'
    $text = ''

    switch -Regex ($provider) {

        # ---------- Anthropic ----------
        '^Anthropic$' {
            $body = @{
                model      = $model
                max_tokens = 4096
                system     = $systemPrompt
                tools      = (Build-AnthropicToolsPayload -Catalog $Catalog)
                messages   = @($safeMessages | ForEach-Object {
                                  if ($_.content -is [string]) { @{ role = $_.role; content = $_.content } }
                                  else { @{ role = $_.role; content = $_.content } }
                              })
            } | ConvertTo-Json -Depth 12
            $headers = @{ 'x-api-key' = $key; 'anthropic-version' = '2023-06-01'; 'content-type' = 'application/json' }
            try {
                $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -ErrorAction Stop
            } catch { return @{ Error = $_.Exception.Message } }
            $rawAssistantContent = @($resp.content)
            $stopReason = [string]$resp.stop_reason
            foreach ($b in $resp.content) {
                if ($b.type -eq 'text') { $text += $b.text }
                elseif ($b.type -eq 'tool_use') {
                    $inputHash = @{}
                    if ($b.input) { foreach ($p in $b.input.PSObject.Properties) { $inputHash[$p.Name] = $p.Value } }
                    [void]$toolUses.Add(@{ id = $b.id; name = $b.name; input = $inputHash })
                }
            }
            if ($resp.usage) { $usage = @{ InputTokens = [int]$resp.usage.input_tokens; OutputTokens = [int]$resp.usage.output_tokens } }
        }

        # ---------- OpenAI / Azure OpenAI ----------
        '^(OpenAI|AzureOpenAI)$' {
            $body = @{
                model       = $model
                messages    = @(@{ role='system'; content=$systemPrompt }) + $safeMessages
                tools       = (Build-OpenAIToolsPayload -Catalog $Catalog)
                tool_choice = 'auto'
                max_tokens  = 4096
            } | ConvertTo-Json -Depth 12
            $headers = @{ 'Content-Type' = 'application/json' }
            if ($provider -eq 'AzureOpenAI') { $headers['api-key'] = $key } else { $headers['Authorization'] = "Bearer $key" }
            try {
                $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -ErrorAction Stop
            } catch { return @{ Error = $_.Exception.Message } }
            $msg = $resp.choices[0].message
            $rawAssistantContent = @($msg)
            $stopReason = [string]$resp.choices[0].finish_reason
            if ($msg.content) { $text += [string]$msg.content }
            if ($msg.tool_calls) {
                foreach ($tc in $msg.tool_calls) {
                    $argsRaw = [string]$tc.function.arguments
                    $inputHash = @{}
                    try {
                        $argsObj = $argsRaw | ConvertFrom-Json -ErrorAction Stop
                        foreach ($p in $argsObj.PSObject.Properties) { $inputHash[$p.Name] = $p.Value }
                    } catch { Write-Warn "OpenAI tool_call arguments not valid JSON: $argsRaw" }
                    [void]$toolUses.Add(@{ id = $tc.id; name = $tc.function.name; input = $inputHash })
                }
            }
            if ($resp.usage) { $usage = @{ InputTokens = [int]$resp.usage.prompt_tokens; OutputTokens = [int]$resp.usage.completion_tokens } }
        }

        # ---------- Ollama ----------
        '^Ollama$' {
            $body = @{
                model    = $model
                stream   = $false
                messages = @(@{ role='system'; content=$systemPrompt }) + $safeMessages
                tools    = (Build-OpenAIToolsPayload -Catalog $Catalog)
            } | ConvertTo-Json -Depth 12
            $headers = @{ 'Content-Type' = 'application/json' }
            $uri = "$endpoint/api/chat"
            try {
                $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -ErrorAction Stop
            } catch {
                # If Ollama returned an error about tools, the model probably doesn't support them
                $msg = $_.Exception.Message
                if ($msg -match 'tools|tool_calls|unsupported') {
                    $script:AIToolCapabilityCache[("ollama::" + $Config.Model)] = $false
                    return @{ Error = "ollama_no_tool_support" }
                }
                return @{ Error = $msg }
            }
            $rawAssistantContent = @($resp.message)
            $text = [string]$resp.message.content
            $stopReason = 'end_turn'
            if ($resp.message.tool_calls) {
                foreach ($tc in $resp.message.tool_calls) {
                    $inputHash = @{}
                    if ($tc.function.arguments) {
                        # Ollama returns arguments as a hashtable already (not a JSON string).
                        if ($tc.function.arguments -is [string]) {
                            try {
                                $argsObj = $tc.function.arguments | ConvertFrom-Json
                                foreach ($p in $argsObj.PSObject.Properties) { $inputHash[$p.Name] = $p.Value }
                            } catch {}
                        } else {
                            foreach ($p in $tc.function.arguments.PSObject.Properties) { $inputHash[$p.Name] = $p.Value }
                        }
                    }
                    [void]$toolUses.Add(@{ id = ("ollama-call-" + [guid]::NewGuid().ToString().Substring(0,8)); name = $tc.function.name; input = $inputHash })
                }
            }
            $usage = @{
                InputTokens  = if ($resp.prompt_eval_count) { [int]$resp.prompt_eval_count } else { 0 }
                OutputTokens = if ($resp.eval_count)        { [int]$resp.eval_count }        else { 0 }
            }
        }

        default {
            return @{ Error = "Tool calling not implemented for provider '$provider'." }
        }
    }

    # ---- Restore placeholders in assistant text + tool_use inputs before display ----
    if ($text) { $text = Restore-FromSafePayload -Text $text }
    foreach ($tu in $toolUses) {
        $restored = @{}
        foreach ($k in $tu.input.Keys) {
            $v = $tu.input[$k]
            if ($v -is [string]) { $restored[$k] = Restore-FromSafePayload -Text $v }
            else { $restored[$k] = $v }
        }
        $tu.input = $restored
    }

    return @{
        Text             = $text
        ToolUses         = @($toolUses)
        AssistantContent = $rawAssistantContent
        StopReason       = $stopReason
        Usage            = $usage
        Error            = $null
    }
}

# ============================================================
#  Tool-result builder -- shape the right block per provider
# ============================================================

function Build-ToolResultMessage {
    <#
        Given a provider name + array of tool-result records
        (id, name, content), return a properly-shaped message
        to append to the conversation history.
    #>
    param(
        [Parameter(Mandatory)][string]$Provider,
        [Parameter(Mandatory)][array]$ToolResults
    )
    switch ($Provider) {
        'Anthropic' {
            $blocks = @()
            foreach ($r in $ToolResults) {
                $blocks += @{
                    type        = 'tool_result'
                    tool_use_id = $r.id
                    content     = [string]$r.content
                    is_error    = [bool]$r.isError
                }
            }
            return @{ role = 'user'; content = $blocks }
        }
        default {
            # OpenAI / Azure / Ollama: one separate message per tool result
            $msgs = @()
            foreach ($r in $ToolResults) {
                $msgs += @{
                    role         = 'tool'
                    tool_call_id = $r.id
                    content      = [string]$r.content
                }
            }
            return ,$msgs   # comma to force array wrap
        }
    }
}

# ============================================================
#  Format a tool result for AI consumption (truncated, JSON)
# ============================================================

function Format-AIToolResultPayload {
    param($Result, [int]$MaxBytes = 4000)
    if (-not $Result) { return '(no output)' }
    try {
        $json = $Result | ConvertTo-Json -Depth 6 -Compress
    } catch { $json = [string]$Result }
    if ($json.Length -gt $MaxBytes) {
        $json = $json.Substring(0, $MaxBytes) + "...[truncated $($json.Length - $MaxBytes) chars]"
    }
    return $json
}
