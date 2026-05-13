# ============================================================
#  AIAssistant.ps1 - AI Assistant "Mark" (Hidden Option 99)
#  Code-level safeguards for small local models
# ============================================================

$script:AIConfigPath = $null
if ($PSScriptRoot) { $script:AIConfigPath = Join-Path $PSScriptRoot "ai_config.json" }
elseif ($env:M365ADMIN_ROOT) { $script:AIConfigPath = Join-Path $env:M365ADMIN_ROOT "ai_config.json" }
elseif ($MyInvocation.MyCommand.Path) { $script:AIConfigPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "ai_config.json" }
else { $script:AIConfigPath = Join-Path (Get-Location).Path "ai_config.json" }

# Keep the system prompt SHORT - small models lose track with long prompts
$script:AISystemPrompt = @"
You are Mark, an AI assistant in an M365 admin tool. You execute PowerShell commands.

To run a command, put it on its own line starting with RUN:
Example:
RUN: Get-MgUser -Search "displayName:Jackie" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department"

RULES:
1. NEVER say "I found" or "I have successfully" BEFORE a RUN: command. You do NOT have results yet.
2. ONLY say "Let me look up..." or "Let me search..." before a RUN: command.
3. Always search by name first. Never ask the user for email addresses.
4. Add -ErrorAction Stop to all commands. Use -Confirm:`$false on destructive commands.
5. When multiple users are found, ask which one before proceeding.
6. For onboarding, ask: email to create, job title, department, usage location, copy from existing user?
7. For offboarding, search user first, then ask: OOO message? Forward to? Who gets access?

CORRECT:
"Let me look up that user."
RUN: Get-MgUser -Search "displayName:name" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department"

WRONG (never do this):
"I found John (john@domain.com)." <-- you don't have results yet!
"I have successfully removed John." <-- the command hasn't run yet!

KEY COMMANDS:
Search user: Get-MgUser -Search "displayName:NAME" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,JobTitle,Department"
Search group: Get-MgGroup -Search "displayName:NAME" -ConsistencyLevel eventual -Property "Id,DisplayName,Mail"
Group members: Get-MgGroupMember -GroupId "ID" -All | Select-Object Id, @{N='Name';E={`$_.AdditionalProperties['displayName']}}, @{N='Email';E={`$_.AdditionalProperties['userPrincipalName']}}
User groups: Get-MgUserMemberOf -UserId "ID" -All | Select-Object @{N='Name';E={`$_.AdditionalProperties['displayName']}}, @{N='Type';E={`$_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.',''}}
Add to group: New-MgGroupMember -GroupId "GID" -DirectoryObjectId "UID" -ErrorAction Stop
Remove from group: Remove-MgGroupMemberByRef -GroupId "GID" -DirectoryObjectId "UID" -ErrorAction Stop
Search DL: Get-DistributionGroup -Filter "DisplayName -like '*NAME*'" -ErrorAction Stop
DL by email: Get-DistributionGroup -Identity "dl@domain.com" -ErrorAction Stop
DL members: Get-DistributionGroupMember -Identity "dl@domain.com" -ErrorAction Stop
Add DL member: Add-DistributionGroupMember -Identity "dl@domain.com" -Member "user@domain.com" -ErrorAction Stop
Remove DL member: Remove-DistributionGroupMember -Identity "dl@domain.com" -Member "user@domain.com" -Confirm:`$false -ErrorAction Stop
Mailbox: Get-Mailbox -Identity "user@domain.com" -ErrorAction Stop
Mailbox stats: Get-EXOMailboxStatistics -Identity "user@domain.com" -ErrorAction Stop
Add mailbox access: Add-MailboxPermission -Identity "mb" -User "u" -AccessRights FullAccess -AutoMapping `$true -ErrorAction Stop
Send As: Add-RecipientPermission -Identity "mb" -Trustee "u" -AccessRights SendAs -Confirm:`$false -ErrorAction Stop
Block sign-in: Update-MgUser -UserId "ID" -AccountEnabled:`$false -ErrorAction Stop
Revoke sessions: Revoke-MgUserSignInSession -UserId "ID" -ErrorAction Stop
Convert shared: Set-Mailbox -Identity "user" -Type Shared -ErrorAction Stop
Set forwarding: Set-Mailbox -Identity "user" -ForwardingSmtpAddress "smtp:target@domain.com" -DeliverToMailboxAndForward `$true -ErrorAction Stop
Set OOO: Set-MailboxAutoReplyConfiguration -Identity "user" -AutoReplyState Enabled -InternalMessage "msg" -ExternalMessage "msg" -ErrorAction Stop
Licenses: Get-MgUserLicenseDetail -UserId "ID"
All SKUs: Get-MgSubscribedSku -ErrorAction Stop | Select-Object SkuPartNumber,SkuId,ConsumedUnits,@{N='Total';E={`$_.PrepaidUnits.Enabled}}
Add license: Set-MgUserLicense -UserId "ID" -AddLicenses @(@{SkuId="GUID"}) -RemoveLicenses @() -ErrorAction Stop
Remove license: Set-MgUserLicense -UserId "ID" -AddLicenses @() -RemoveLicenses @("GUID") -ErrorAction Stop
Calendar perms: Get-MailboxFolderPermission -Identity "user:\Calendar" -ErrorAction Stop
Add calendar: Add-MailboxFolderPermission -Identity "user:\Calendar" -User "user2" -AccessRights Editor -ErrorAction Stop
Create user: New-MgUser -BodyParameter @{DisplayName="Name";GivenName="First";Surname="Last";UserPrincipalName="upn@domain.com";MailNickname="nick";AccountEnabled=`$true;PasswordProfile=@{Password="<GENERATED_TEMP_PASSWORD>";ForceChangePasswordNextSignIn=`$true};UsageLocation="US";JobTitle="Title";Department="Dept"} -ErrorAction Stop
Set manager: Set-MgUserManagerByRef -UserId "UID" -BodyParameter @{"@odata.id"="https://graph.microsoft.com/v1.0/users/MANAGER_ID"} -ErrorAction Stop
"@

# ============================================================
#  Config — API key is DPAPI-encrypted at rest (per-user, per-machine).
#  Stored values are prefixed:
#     DPAPI:<hex>  → Windows DPAPI ciphertext (preferred)
#     B64:<base64> → base64 fallback on non-Windows (NOT real encryption,
#                    just obfuscation — flagged with a warning at save time)
#     <anything else> → legacy plaintext, migrated on load
# ============================================================

function Protect-ApiKey {
    param([string]$PlainKey)
    if ([string]::IsNullOrWhiteSpace($PlainKey) -or $PlainKey -eq "none") { return $PlainKey }
    if ($PlainKey -like "DPAPI:*" -or $PlainKey -like "B64:*") { return $PlainKey }
    try {
        $secure = ConvertTo-SecureString $PlainKey -AsPlainText -Force
        $enc    = ConvertFrom-SecureString $secure
        return "DPAPI:$enc"
    } catch {
        Write-Warn "DPAPI unavailable on this host; API key will be base64-obfuscated only (NOT encrypted). Reason: $_"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainKey)
        return "B64:$([Convert]::ToBase64String($bytes))"
    }
}

function Unprotect-ApiKey {
    param([string]$StoredKey)
    if ([string]::IsNullOrWhiteSpace($StoredKey) -or $StoredKey -eq "none") { return $StoredKey }
    if ($StoredKey -like "DPAPI:*") {
        try {
            $secure = ConvertTo-SecureString $StoredKey.Substring(6)
            return [System.Net.NetworkCredential]::new("", $secure).Password
        } catch {
            throw "Failed to decrypt API key. Likely cause: config was encrypted by a different user account or on a different machine. Re-run AI setup (option 99 → /config)."
        }
    }
    if ($StoredKey -like "B64:*") {
        try {
            return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($StoredKey.Substring(4)))
        } catch { throw "Failed to decode API key (corrupt base64). Re-run AI setup." }
    }
    return $StoredKey  # legacy plaintext — caller is expected to re-save through Save-AIConfig to migrate
}

function Get-AIConfig {
    if (-not (Test-Path $script:AIConfigPath)) { return $null }
    $raw = $null
    try { $raw = Get-Content $script:AIConfigPath -Raw | ConvertFrom-Json } catch { return $null }
    if ($null -eq $raw) { return $null }

    # Normalize to hashtable so callers don't need the PSCustomObject branch
    $ht = @{}
    foreach ($prop in $raw.PSObject.Properties) {
        # Skip _comment_* documentation keys from the example file
        if ($prop.Name -like "_comment*") { continue }
        $ht[$prop.Name] = $prop.Value
    }

    # ---- Migrate legacy plaintext keys ----
    if ($ht.ContainsKey("ApiKey")) {
        $key = [string]$ht["ApiKey"]
        if ($key -and $key -ne "none" -and $key -notlike "DPAPI:*" -and $key -notlike "B64:*" -and $key -ne "REPLACE_ME") {
            Write-Warn "Plaintext API key detected in $($script:AIConfigPath). Encrypting in place..."
            $ht["ApiKey"] = Protect-ApiKey -PlainKey $key
            Save-AIConfig -Config $ht
            Write-Success "API key encrypted (DPAPI, per-user, non-portable)."
        }
    }

    # ---- Privacy section migration (commit 6) ----
    $privacyDefaults = @{
        ExternalRedaction       = 'Enabled'
        RedactInAuditLog        = 'Disabled'
        ExternalPayloadCapBytes = 8192
        TrustedProviders        = @()
    }
    $privacyHt = @{}
    $wroteDefaults = $false
    if ($ht.ContainsKey('Privacy') -and $null -ne $ht['Privacy']) {
        $p = $ht['Privacy']
        if ($p -is [hashtable]) {
            foreach ($k in $p.Keys) { if ($k -notlike '_comment*') { $privacyHt[$k] = $p[$k] } }
        } else {
            # PSCustomObject from JSON — normalize and drop _comment_* documentation keys
            foreach ($prop in $p.PSObject.Properties) {
                if ($prop.Name -like '_comment*') { continue }
                $privacyHt[$prop.Name] = $prop.Value
            }
        }
    } else {
        $wroteDefaults = $true
    }
    foreach ($k in $privacyDefaults.Keys) {
        if (-not $privacyHt.ContainsKey($k)) {
            $privacyHt[$k] = $privacyDefaults[$k]
            $wroteDefaults = $true
        }
    }
    # TrustedProviders may come back as a typed array — normalize to plain array of lowercase strings
    if ($privacyHt['TrustedProviders']) {
        $privacyHt['TrustedProviders'] = @($privacyHt['TrustedProviders'] | ForEach-Object { ([string]$_).ToLowerInvariant() })
    } else {
        $privacyHt['TrustedProviders'] = @()
    }
    $ht['Privacy'] = $privacyHt
    if ($wroteDefaults) {
        Save-AIConfig -Config $ht
    }

    return $ht
}

function Save-AIConfig {
    param([hashtable]$Config)
    $toSave = @{}
    foreach ($k in $Config.Keys) { $toSave[$k] = $Config[$k] }
    if ($toSave.ContainsKey("ApiKey")) {
        $toSave["ApiKey"] = Protect-ApiKey -PlainKey ([string]$toSave["ApiKey"])
    }
    $toSave | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:AIConfigPath -Encoding UTF8 -Force
}

function Setup-AIProvider {
    Write-SectionHeader "AI Assistant Setup"
    $provider = Show-Menu -Title "Select AI Provider" -Options @("Ollama (local)","Anthropic (Claude)","OpenAI (GPT)","Azure OpenAI","Custom endpoint") -BackLabel "Cancel"
    if ($provider -eq -1) { return $null }
    $config = @{}
    switch ($provider) {
        0 {
            $config["Provider"]="Ollama"; $url=Read-UserInput "Ollama URL (Enter for http://localhost:11434)"; $config["Endpoint"]=if($url){$url}else{"http://localhost:11434"}; $config["ApiKey"]="none"
            try { $models=Invoke-RestMethod -Uri "$($config['Endpoint'])/api/tags" -Method GET -TimeoutSec 5 -ErrorAction Stop
                if($models.models.Count -gt 0){Write-Success "Found $($models.models.Count) model(s)."; $labels=$models.models|ForEach-Object{"$($_.name) ($([math]::Round($_.size/1MB))MB)"}
                    $sel=Show-Menu -Title "Select model" -Options $labels -BackLabel "Type manually"; $config["Model"]=if($sel -eq -1){Read-UserInput "Model name"}else{$models.models[$sel].name}
                }else{Write-Warn "No models. Run: ollama pull llama3";$config["Model"]=Read-UserInput "Model name"}
            } catch { Write-ErrorMsg "Cannot reach Ollama."; $config["Model"]=Read-UserInput "Model name" }
            if([string]::IsNullOrWhiteSpace($config["Model"])){return $null}; Save-AIConfig -Config $config; return $config
        }
        1 { $config["Provider"]="Anthropic";$config["Endpoint"]="https://api.anthropic.com/v1/messages";$config["Model"]="claude-sonnet-4-20250514";$m=Read-UserInput "Model (Enter for default)";if($m){$config["Model"]=$m} }
        2 { $config["Provider"]="OpenAI";$config["Endpoint"]="https://api.openai.com/v1/chat/completions";$config["Model"]="gpt-4o";$m=Read-UserInput "Model (Enter for gpt-4o)";if($m){$config["Model"]=$m} }
        3 { $ep=Read-UserInput "Azure endpoint";if(!$ep){return $null};$d=Read-UserInput "Deployment";if(!$d){return $null};$v=Read-UserInput "API version (Enter for 2024-02-01)";if(!$v){$v="2024-02-01"};$config["Provider"]="AzureOpenAI";$config["Endpoint"]="$ep/openai/deployments/$d/chat/completions?api-version=$v";$config["Model"]=$d }
        4 { $ep=Read-UserInput "Endpoint URL";if(!$ep){return $null};$config["Provider"]="Custom";$config["Endpoint"]=$ep;$m=Read-UserInput "Model";$config["Model"]=if($m){$m}else{"default"} }
    }
    if($config["Provider"] -ne "Ollama"){$k=Read-UserInput "API Key";if([string]::IsNullOrWhiteSpace($k)){return $null};$config["ApiKey"]=$k}
    Save-AIConfig -Config $config; Write-Success "$($config['Provider']) / $($config['Model']) configured."; return $config
}

function Show-PrivacyMenu {
    <#
        Interactive editor for the Privacy section of ai_config.json.
        Toggles ExternalRedaction / RedactInAuditLog, prompts for an
        integer payload cap, edits the TrustedProviders list, and lets
        the operator reset the session token map. Persists via
        Save-AIConfig on exit.
    #>
    param([hashtable]$Config)
    if (-not $Config.ContainsKey('Privacy') -or -not ($Config['Privacy'] -is [hashtable])) {
        $Config['Privacy'] = @{
            ExternalRedaction       = 'Enabled'
            RedactInAuditLog        = 'Disabled'
            ExternalPayloadCapBytes = 8192
            TrustedProviders        = @()
        }
    }
    $priv = $Config['Privacy']
    $running = $true
    while ($running) {
        $tpDisplay = '(none)'
        if ($priv['TrustedProviders'] -and @($priv['TrustedProviders']).Count -gt 0) {
            $tpDisplay = ($priv['TrustedProviders'] -join ', ')
        }
        $tokens = 0
        if ($script:PrivacyMap -and $script:PrivacyMap.ByToken) { $tokens = $script:PrivacyMap.ByToken.Count }
        $cap = [int]$priv['ExternalPayloadCapBytes']
        $capLabel = if ($cap -eq 0) { "(no cap)" } else { "$cap bytes" }

        $sel = Show-Menu -Title "Privacy Settings" -Options @(
            "External redaction        : $($priv['ExternalRedaction'])",
            "Redact in audit log       : $($priv['RedactInAuditLog'])",
            "External payload cap      : $capLabel",
            "Trusted providers         : $tpDisplay",
            "Reset session token map   : $tokens token(s) currently stored"
        ) -BackLabel "Save and exit"

        switch ($sel) {
            0 {
                $priv['ExternalRedaction'] = if ($priv['ExternalRedaction'] -eq 'Enabled') { 'Disabled' } else { 'Enabled' }
            }
            1 {
                $priv['RedactInAuditLog'] = if ($priv['RedactInAuditLog'] -eq 'Enabled') { 'Disabled' } else { 'Enabled' }
            }
            2 {
                $v = Read-UserInput "Cap bytes (non-negative integer; 0 disables) [current $cap]"
                if (-not [string]::IsNullOrWhiteSpace($v)) {
                    $n = 0
                    if ([int]::TryParse($v, [ref]$n) -and $n -ge 0) {
                        $priv['ExternalPayloadCapBytes'] = $n
                    } else {
                        Write-ErrorMsg "Invalid integer. Keeping $cap."
                    }
                }
            }
            3 {
                $v = Read-UserInput "Trusted providers (comma-separated lowercase; blank to clear). Known: anthropic, openai, azure-openai, custom, ollama"
                if ($null -ne $v) {
                    $items = @($v -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
                    $priv['TrustedProviders'] = $items
                    if ($items.Count -gt 0) {
                        Write-Warn "Listed providers will receive RAW PII (secrets still scrubbed). Confirm each is in your compliance boundary."
                    }
                }
            }
            4 {
                $cleared = Reset-PrivacyMap
                Write-AIAuditEntry -EventType "CLEAR" -Detail ("privacy map manually reset ({0} tokens)" -f $cleared)
                Write-Success "Cleared $cleared token(s)."
            }
            -1 { $running = $false }
        }
    }
    $Config['Privacy'] = $priv
    Save-AIConfig -Config $Config
    Write-Success "Privacy settings saved."
}

# ============================================================
#  API Call
# ============================================================

function Invoke-AIChat {
    param([hashtable]$Config, [array]$Messages)

    $p = $Config["Provider"]
    try { $k = Unprotect-ApiKey -StoredKey ([string]$Config["ApiKey"]) }
    catch { return "[!] $_" }
    $ep = $Config["Endpoint"]
    $m  = $Config["Model"]

    # ---- Privacy classification ----
    $privacy = $null
    if ($Config.ContainsKey('Privacy') -and $Config['Privacy'] -is [hashtable]) {
        $privacy = $Config['Privacy']
    } else {
        $privacy = @{ ExternalRedaction='Enabled'; ExternalPayloadCapBytes=8192; TrustedProviders=@() }
    }
    $isExternal  = Test-IsExternalProvider -Provider $p -Endpoint $ep -TrustedProviders $privacy['TrustedProviders']
    $fullRedact  = $isExternal -and ($privacy['ExternalRedaction'] -eq 'Enabled')
    $secretsOnly = -not $fullRedact   # secrets are scrubbed even for local

    # ---- Tokenize messages + apply byte cap ----
    $counts = @{ JWT=0; SECRET=0; THUMB=0; UPN=0; GUID=0; TENANT=0; NAME=0 }
    $safeMessages = @()
    $cap = 0
    if ($isExternal) { $cap = [int]($privacy['ExternalPayloadCapBytes']) }
    $truncated = $false
    foreach ($msg in $Messages) {
        $content = [string]$msg.content
        if ($content) {
            $content = Convert-ToSafePayload -Text $content -SecretsOnly:$secretsOnly -Counts $counts
        }
        if ($cap -gt 0 -and $content.Length -gt $cap) {
            $orig = $content.Length
            $content = $content.Substring(0, $cap) + "`n...[TRUNCATED $($orig - $cap) BYTES AFTER REDACTION]"
            $truncated = $true
        }
        $safeMessages += @{ role = $msg.role; content = $content }
    }

    # ---- System prompt (with placeholder note appended for external providers) ----
    $safeSystemPrompt = $script:AISystemPrompt
    if ($fullRedact) { $safeSystemPrompt = $safeSystemPrompt + $script:PrivacySystemPromptAddendum }

    # ---- Audit ----
    $auditDetail = "provider={0} external={1} redact={2} cap={3} truncated={4} | {5}" -f `
        (Get-ProviderCanonicalName $p), $isExternal, $(if ($fullRedact) {'full'} else {'secrets-only'}), $cap, $truncated, (Format-CountsForAudit $counts)
    Write-AIAuditEntry -EventType "REDACT" -Detail $auditDetail

    # ---- Call provider ----
    try {
        $rawResponse = $null
        $usage = @{ InputTokens = 0; OutputTokens = 0 }
        switch ($p) {
            "Ollama" {
                # Real streaming -- print the response as it arrives.
                if (Get-Command Invoke-OllamaStream -ErrorAction SilentlyContinue) {
                    $stream = Invoke-OllamaStream -Endpoint $ep -Model $m -SystemPrompt $safeSystemPrompt -SafeMessages $safeMessages
                    $rawResponse = $stream.Text
                    $usage = $stream.Usage
                    $script:LastAIChatStreamed = $true
                } else {
                    $body = @{
                        model    = $m
                        stream   = $false
                        messages = @(@{ role='system'; content=$safeSystemPrompt }) + $safeMessages
                    } | ConvertTo-Json -Depth 10
                    $resp = Invoke-RestMethod -Uri "$ep/api/chat" -Method POST -ContentType "application/json" -Body $body -TimeoutSec 180 -ErrorAction Stop
                    $rawResponse = $resp.message.content
                    if ($resp.prompt_eval_count) { $usage.InputTokens  = [int]$resp.prompt_eval_count }
                    if ($resp.eval_count)        { $usage.OutputTokens = [int]$resp.eval_count }
                }
            }
            "Anthropic" {
                $body = @{
                    model      = $m
                    max_tokens = 4096
                    system     = $safeSystemPrompt
                    messages   = $safeMessages
                } | ConvertTo-Json -Depth 10
                $h = @{ "x-api-key"=$k; "anthropic-version"="2023-06-01"; "content-type"="application/json" }
                $resp = Invoke-RestMethod -Uri $ep -Method POST -Headers $h -Body $body -ErrorAction Stop
                $rawResponse = $resp.content[0].text
                if ($resp.usage) {
                    $usage.InputTokens  = [int]$resp.usage.input_tokens
                    $usage.OutputTokens = [int]$resp.usage.output_tokens
                }
            }
            default {
                $body = @{
                    model      = $m
                    messages   = @(@{ role='system'; content=$safeSystemPrompt }) + $safeMessages
                    max_tokens = 4096
                } | ConvertTo-Json -Depth 10
                $h = @{ "Content-Type"="application/json" }
                if ($p -eq "AzureOpenAI") { $h["api-key"] = $k } else { $h["Authorization"] = "Bearer $k" }
                $resp = Invoke-RestMethod -Uri $ep -Method POST -Headers $h -Body $body -ErrorAction Stop
                $rawResponse = $resp.choices[0].message.content
                if ($resp.usage) {
                    $usage.InputTokens  = [int]$resp.usage.prompt_tokens
                    $usage.OutputTokens = [int]$resp.usage.completion_tokens
                }
            }
        }

        # ---- Cost tracking ----
        if (Get-Command Add-AICostEvent -ErrorAction SilentlyContinue) {
            try {
                $script:LastAICostResult = Add-AICostEvent -Config $Config -Usage $usage -Reason 'chat'
            } catch { Write-Warn "Cost tracker: $($_.Exception.Message)" }
        }

        # ---- Restore placeholders so the operator sees real values and
        #      Extract-Commands sees real cmdlet arguments ----
        $restored = Restore-FromSafePayload -Text ([string]$rawResponse)
        return $restored
    } catch {
        $e = $_.Exception.Message
        if ($e -match "401|Unauth")            { return "[!] Invalid API key." }
        if ($e -match "429")                   { return "[!] Rate limited." }
        if ($e -match "resolve|connect|refused"){ return "[!] Cannot reach $(if($p -eq 'Ollama'){'Ollama'}else{'API'})." }
        return "[!] API error: $e"
    }
}

# ============================================================
#  Auto-connect services
# ============================================================

function Ensure-ServiceForCommand {
    param([string]$Command)
    if ($Command -match 'Get-Mg|Set-Mg|New-Mg|Update-Mg|Remove-Mg|Revoke-Mg|MgSubscribedSku|MgUserLicense|MgGroupMember|MgUserMember|Invoke-MgGraphRequest') {
        if (-not $script:SessionState.MgGraph) { Write-InfoMsg "Connecting Microsoft Graph..."; return (Connect-Graph) }; return $true
    }
    if ($Command -match 'Mailbox|DistributionGroup|RecipientPermission|EXO|ManagedFolderAssistant') {
        if (-not $script:SessionState.ExchangeOnline) { Write-InfoMsg "Connecting Exchange Online..."; return (Connect-EXO) }; return $true
    }
    if ($Command -match 'Compliance') {
        if (-not $script:SessionState.ComplianceCenter) { Write-InfoMsg "Connecting SCC..."; return (Connect-SCC) }; return $true
    }
    return $true
}

# ============================================================
#  Command Extraction with sanitization
# ============================================================

function Extract-Commands {
    param([string]$Response)
    $commands = @()

    # Pattern 1: RUN: command
    foreach ($line in ($Response -split "`n")) {
        $line = $line.Trim()
        if ($line -match '^RUN:\s*(.+)$') {
            $cmd = $Matches[1].Trim()
            $cmd = $cmd -replace '`$', '' -replace '^`+', '' -replace '^\$\s', '' -replace '</?cmd\s*>', '' -replace '\s*>$', ''
            if ($cmd.Length -gt 5) { $commands += $cmd }
        }
    }

    # Pattern 2: <cmd> tags fallback
    if ($commands.Count -eq 0) {
        $tagMatches = [regex]::Matches($Response, '<cmd>([\s\S]*?)</cmd>')
        foreach ($m in $tagMatches) { $cmd = $m.Groups[1].Value.Trim() -replace '</?cmd\s*>',''-replace '\s*>$',''; if($cmd.Length -gt 5){$commands += $cmd} }
    }

    # Pattern 3: ```powershell blocks fallback
    if ($commands.Count -eq 0) {
        $codeMatches = [regex]::Matches($Response, '```(?:powershell)?\s*\n([\s\S]*?)```')
        foreach ($m in $codeMatches) {
            foreach ($cline in ($m.Groups[1].Value.Trim() -split "`n")) {
                $cline = $cline.Trim()
                if ($cline -match '^(Get-|Set-|New-|Remove-|Add-|Update-|Revoke-|Enable-|Disable-|Start-)') { $commands += $cline }
            }
        }
    }

    # ---- AUTO-FIX: Replace Get-MgGroupMember with Graph API direct call for proper names ----
    for ($i = 0; $i -lt $commands.Count; $i++) {
        if ($commands[$i] -match 'Get-MgGroupMember.*-GroupId\s+"?([a-f0-9-]+)"?') {
            $groupId = $Matches[1]
            $commands[$i] = "Invoke-MgGraphRequest -Method GET -Uri `"https://graph.microsoft.com/v1.0/groups/$groupId/members?`$select=displayName,userPrincipalName,id&`$top=999`" -ErrorAction Stop | ForEach-Object { `$_.value } | Select-Object displayName, userPrincipalName, id"
        }
    }

    return $commands
}

function Get-CleanResponse {
    <# Strip commands AND hallucinated claims from AI text #>
    param([string]$Response)
    $clean = $Response -replace '(?m)^RUN:\s*.+$', ''
    $clean = $clean -replace '<cmd>[\s\S]*?</cmd>', ''
    $clean = $clean -replace '```(?:powershell)?[\s\S]*?```', ''

    # Strip hallucinated claims that appear before commands run
    $clean = $clean -replace '(?i)I found\s+.+?\([^)]*@[^)]*\)[^.]*\.?', ''
    $clean = $clean -replace '(?i)I have (successfully|already)\s+(removed|added|updated|created|set|enabled|disabled|revoked|blocked|converted)[^.]*\.?', ''
    $clean = $clean -replace '(?i)(successfully|already)\s+(removed|added|updated|created|set|enabled|disabled|revoked|blocked|converted)[^.]*\.?', ''
    $clean = $clean -replace '(?i)I found the following[^:]*:[^.]*\.?', ''
    $clean = $clean -replace '(?i)Please confirm if this is[^.]*\.?', ''

    return (($clean -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join "`n").Trim()
}

function Test-HasCommands { param([string]$Response); return ($Response -match '(?m)^RUN:\s' -or $Response -match '<cmd>' -or $Response -match '```powershell') }

function Limit-ResultsForAI {
    <# Truncates command output to keep the AI context manageable. Shows first/last lines. #>
    param([string]$Text, [int]$MaxLines = 30)
    $lines = $Text -split "`n"
    if ($lines.Count -le $MaxLines) { return $Text }
    $head = $lines[0..14] -join "`n"
    $tail = $lines[($lines.Count - 10)..($lines.Count - 1)] -join "`n"
    return "$head`n... ($($lines.Count) total lines, truncated) ...`n$tail"
}

# ============================================================
#  Privacy / PII Tokenization
#  Outbound payloads to non-local LLM providers are tokenized: each
#  piece of PII (UPN/email/GUID/tenant-ID/cert-thumbprint/JWT/API-key
#  /display-name) is replaced with a stable opaque placeholder
#  (<UPN_1>, <GUID_3>, <TENANT>, ...) backed by a session-scoped
#  reverse map. The AI's response is restored before display and
#  before scriptblock execution, so the operator sees real values and
#  generated commands target real objects. Map clears on /clear and
#  on assistant exit.
#
#  Secrets (JWT, sk-*, sk-ant-*, cert thumbprints) are ALWAYS
#  tokenized regardless of provider — that rule is hardcoded, not
#  controlled by config.
# ============================================================

$script:PrivacyPatterns = @(
    # Order matters — higher-specificity patterns first.
    # SecretsOnly=$true means these run even when ExternalRedaction is
    # Disabled and even for local providers — secrets must never leak.
    @{ Type='JWT';    Regex='(?<![\w.-])eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+'; SecretsOnly=$true  },
    @{ Type='SECRET'; Regex='sk-ant-[A-Za-z0-9_\-]{20,}';                                    SecretsOnly=$true  },
    @{ Type='SECRET'; Regex='sk-[A-Za-z0-9]{20,}';                                            SecretsOnly=$true  },
    @{ Type='THUMB';  Regex='\b[0-9A-Fa-f]{40}\b';                                            SecretsOnly=$true  },
    @{ Type='UPN';    Regex='\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b';           SecretsOnly=$false },
    @{ Type='GUID';   Regex='\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b'; SecretsOnly=$false }
)

$script:PrivacyMap = $null

function Reset-PrivacyMap {
    $count = 0
    if ($script:PrivacyMap -and $script:PrivacyMap.ByToken) {
        $count = $script:PrivacyMap.ByToken.Count
    }
    $script:PrivacyMap = @{
        ByValue  = @{}
        ByToken  = @{}
        Counters = @{ JWT=0; SECRET=0; THUMB=0; UPN=0; GUID=0; NAME=0; TENANT=0 }
    }
    return $count
}
Reset-PrivacyMap | Out-Null

function Get-OrCreatePrivacyToken {
    param([string]$Value, [string]$Type)
    if ([string]::IsNullOrEmpty($Value)) { return $Value }
    if ($script:PrivacyMap.ByValue.ContainsKey($Value)) {
        return $script:PrivacyMap.ByValue[$Value]
    }
    if (-not $script:PrivacyMap.Counters.ContainsKey($Type)) {
        $script:PrivacyMap.Counters[$Type] = 0
    }
    $script:PrivacyMap.Counters[$Type] = $script:PrivacyMap.Counters[$Type] + 1
    $tok = '<{0}_{1}>' -f $Type, $script:PrivacyMap.Counters[$Type]
    $script:PrivacyMap.ByValue[$Value] = $tok
    $script:PrivacyMap.ByToken[$tok]   = $Value
    return $tok
}

function Test-IsLocalEndpoint {
    param([string]$Endpoint)
    if ([string]::IsNullOrWhiteSpace($Endpoint)) { return $false }
    try {
        $uri = [Uri]$Endpoint
        $h = $uri.Host.ToLowerInvariant()
        if ($h -eq 'localhost' -or $h -eq '127.0.0.1' -or $h -eq '::1') { return $true }
        if ($h.EndsWith('.local')) { return $true }
        return $false
    } catch { return $false }
}

function Get-ProviderCanonicalName {
    param([string]$Provider)
    switch (([string]$Provider).ToLowerInvariant()) {
        'ollama'      { return 'ollama' }
        'anthropic'   { return 'anthropic' }
        'openai'      { return 'openai' }
        'azureopenai' { return 'azure-openai' }
        'custom'      { return 'custom' }
        default       { return ([string]$Provider).ToLowerInvariant() }
    }
}

function Test-IsExternalProvider {
    param([string]$Provider, [string]$Endpoint, [array]$TrustedProviders)
    $canonical = Get-ProviderCanonicalName -Provider $Provider
    $trusted = @()
    if ($TrustedProviders) { $trusted = @($TrustedProviders | ForEach-Object { ([string]$_).ToLowerInvariant() }) }
    if ($trusted -contains $canonical) { return $false }
    if ($canonical -eq 'ollama' -or $canonical -eq 'custom') {
        return -not (Test-IsLocalEndpoint $Endpoint)
    }
    return $true
}

function Convert-ToSafePayload {
    <#
        Tokenize PII in $Text. When $SecretsOnly is true only the
        always-on secret patterns run (JWT / sk-* / sk-ant-* / cert
        thumbprint). Mutates the session privacy map. $Counts is an
        optional hashtable accumulator: keys are token types
        (JWT/SECRET/THUMB/UPN/GUID/TENANT/NAME), values are ints.
    #>
    param(
        [string]   $Text,
        [bool]     $SecretsOnly = $false,
        [hashtable]$Counts = $null
    )
    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    # Tenant ID pre-pass: resolve to <TENANT> before the generic GUID
    # regex sees it, so tenant gets a stable distinguished placeholder.
    if (-not $SecretsOnly -and $script:SessionState -and $script:SessionState.TenantId) {
        $tid = [string]$script:SessionState.TenantId
        if ($tid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$' -and $Text.Contains($tid)) {
            if (-not $script:PrivacyMap.ByValue.ContainsKey($tid)) {
                $script:PrivacyMap.ByValue[$tid] = '<TENANT>'
                $script:PrivacyMap.ByToken['<TENANT>'] = $tid
            }
            $occurrences = ([regex]::Matches($Text, [regex]::Escape($tid))).Count
            $Text = $Text.Replace($tid, '<TENANT>')
            if ($Counts -and $occurrences -gt 0) { $Counts['TENANT'] = (($Counts['TENANT']) + $occurrences) }
        }
    }

    foreach ($pat in $script:PrivacyPatterns) {
        if ($SecretsOnly -and -not $pat.SecretsOnly) { continue }
        $patType = $pat.Type
        $countsRef = $Counts
        $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $tok = Get-OrCreatePrivacyToken -Value $m.Value -Type $patType
            if ($countsRef) { $countsRef[$patType] = (($countsRef[$patType]) + 1) }
            return $tok
        }.GetNewClosure()
        $Text = [regex]::Replace($Text, $pat.Regex, $evaluator)
    }

    if (-not $SecretsOnly) {
        # Display-name capture from common cmdlet idioms. Three forms,
        # tried in order — quoted always wins over bareword so we don't
        # half-tokenize values that are already inside quotes:
        #   (1)  displayName : "Name"  or  displayName : 'Name'
        #   (2)  -DisplayName / -MailNickname / -SamAccountName "Name"
        #   (3)  displayName : Bareword Until ClosingQuote
        # The bareword form is anchored on a closing " so it only fires
        # inside the typical -Search "displayName:..." idiom and won't
        # devour prose.
        $nameCountsRef = $Counts

        $nameEvalQuoted = [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $prefix = $m.Groups[1].Value
            $val = if ($m.Groups[2].Success -and $m.Groups[2].Value) { $m.Groups[2].Value } else { $m.Groups[3].Value }
            if (-not $val -or $val -like '<*_*>' -or $val -eq '<TENANT>') { return $m.Value }
            $tok = Get-OrCreatePrivacyToken -Value $val -Type 'NAME'
            if ($nameCountsRef) { $nameCountsRef['NAME'] = (($nameCountsRef['NAME']) + 1) }
            return ('{0}"{1}"' -f $prefix, $tok)
        }.GetNewClosure()

        $nameEvalBare = [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $prefix = $m.Groups[1].Value
            $val = $m.Groups[2].Value
            if (-not $val -or $val -like '<*_*>' -or $val -eq '<TENANT>') { return $m.Value }
            $tok = Get-OrCreatePrivacyToken -Value $val -Type 'NAME'
            if ($nameCountsRef) { $nameCountsRef['NAME'] = (($nameCountsRef['NAME']) + 1) }
            return ('{0}{1}' -f $prefix, $tok)
        }.GetNewClosure()

        $Text = [regex]::Replace($Text, '(?i)(displayName\s*:\s*)(?:"([^"]+)"|''([^'']+)'')', $nameEvalQuoted)
        $Text = [regex]::Replace($Text, '(?i)(-(?:DisplayName|MailNickname|SamAccountName)\s+)(?:"([^"]+)"|''([^'']+)'')', $nameEvalQuoted)
        $Text = [regex]::Replace($Text, '(?i)(displayName\s*:\s*)([A-Za-z][A-Za-z0-9 \-''\.]{0,80})(?=")', $nameEvalBare)
    }

    return $Text
}

function Restore-FromSafePayload {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    if (-not $script:PrivacyMap -or $script:PrivacyMap.ByToken.Count -eq 0) { return $Text }
    # Longest-token-first: belt + braces. The `>` terminator already
    # prevents <UPN_1> from matching inside <UPN_10>, but this guards
    # against future token shapes.
    $tokens = @($script:PrivacyMap.ByToken.Keys | Sort-Object -Property Length -Descending)
    foreach ($tok in $tokens) {
        $val = [string]$script:PrivacyMap.ByToken[$tok]
        $Text = $Text.Replace($tok, $val)
    }
    return $Text
}

function Format-CountsForAudit {
    param([hashtable]$Counts)
    if (-not $Counts -or $Counts.Count -eq 0) { return '0' }
    return (($Counts.GetEnumerator() | Where-Object { $_.Value -gt 0 } | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ' ')
}

$script:PrivacySystemPromptAddendum = @"


PRIVACY NOTE: Values like <UPN_1>, <GUID_3>, <TENANT>, <NAME_2>, <SECRET_4>, <JWT_1>, <THUMB_1> are opaque placeholders. Preserve them verbatim in any command you propose. Do NOT invent or substitute real-looking values. If a value you need has not been provided, ask the operator rather than guessing.
"@

$script:PlanningSystemPromptAddendum = @"


PLANNING: For tasks that need 3+ tool calls, OR any sequence of destructive operations, you SHOULD first call the special meta-tool 'submit_plan' to propose the full plan. The operator approves, edits, or rejects the plan before any step runs. Single-tool reads (Get-* / List-*) do NOT need a plan. The plan's steps[].tool field must exactly match a real tool name from the catalog (not 'submit_plan' / 'ask_operator'). Use 'dependsOn' to express ordering. Mark each step's 'destructive' field truthfully so the operator sees risk. If the operator asked you to '/plan' in their last message, ALWAYS submit_plan first. If they asked '/noplan', skip planning and call tools directly.
"@

$script:PlanForceSystemPromptAddendum = @"


PLAN MODE FORCED: The operator typed /plan. For this turn you MUST call submit_plan first (do not call any other tool directly). Build a complete plan of every tool call needed to accomplish the operator's request, then stop and wait for approval.
"@

$script:NoPlanSystemPromptAddendum = @"


PLAN MODE SKIPPED: The operator typed /noplan. Do NOT call submit_plan; call the actual tools directly even if the task involves multiple steps.
"@

# ============================================================
#  AI Command Audit + Allow-list
#  Every AI-proposed command is parsed via the PowerShell language
#  parser and checked against $script:AICmdAllowList before any prompt
#  is shown. Default-deny: anything not matched is rejected with no
#  confirmation dialog. Each phase (PROPOSE / SKIP / REJECT / RUN /
#  OK / ERROR) is written to a per-session log under
#  $env:LOCALAPPDATA\M365Manager\audit\ (or ~/.m365manager/audit/ on
#  non-Windows). Sensitive parameter values are redacted at log time.
# ============================================================

$script:AICmdAllowList = @(
    # Microsoft Graph PowerShell SDK
    '*-Mg*',
    'Invoke-MgGraphRequest',

    # Exchange Online — specific cmdlets the AI prompt teaches
    'Get-Mailbox', 'Set-Mailbox',
    'Get-EXOMailbox*', 'Get-MailboxStatistics',
    'Get-MailboxPermission', 'Add-MailboxPermission', 'Remove-MailboxPermission',
    'Get-MailboxFolderPermission', 'Add-MailboxFolderPermission',
    'Remove-MailboxFolderPermission', 'Set-MailboxFolderPermission',
    'Get-MailboxAutoReplyConfiguration', 'Set-MailboxAutoReplyConfiguration',
    'Get-DistributionGroup', 'Get-DistributionGroupMember', 'Set-DistributionGroup',
    'Add-DistributionGroupMember', 'Remove-DistributionGroupMember',
    'Get-Recipient', 'Get-EXORecipient',
    'Add-RecipientPermission', 'Remove-RecipientPermission',
    'Get-UnifiedGroup', 'Get-UnifiedGroupLinks',

    # Security & Compliance Center
    'Get-ComplianceSearch', 'New-ComplianceSearch', 'Start-ComplianceSearch',
    'Remove-ComplianceSearch', 'Get-ComplianceSearchAction',

    # Pure pipeline / formatting — no side effects
    'Select-Object', 'Where-Object', 'ForEach-Object', 'Sort-Object',
    'Group-Object', 'Measure-Object',
    'Format-Table', 'Format-List', 'Format-Wide', 'Out-String', 'Out-Default'
)

$script:AIAuditLogPath = $null

function Get-AIAuditLogPath {
    if ($script:AIAuditLogPath -and (Test-Path -LiteralPath (Split-Path $script:AIAuditLogPath -Parent))) {
        return $script:AIAuditLogPath
    }
    $base = $null
    $onWindows = $false
    if ($env:LOCALAPPDATA) {
        $base = Join-Path $env:LOCALAPPDATA 'M365Manager\audit'
        $onWindows = $true
    } elseif ($env:HOME) {
        $base = Join-Path $env:HOME '.m365manager/audit'
    } else {
        $base = Join-Path (Get-Location).Path 'audit'
    }
    $createdNow = $false
    try {
        if (-not (Test-Path -LiteralPath $base)) {
            New-Item -ItemType Directory -Path $base -Force | Out-Null
            $createdNow = $true
        }
    } catch { return $null }
    # Lock down the audit directory if we just created it. On Windows,
    # %LOCALAPPDATA% already has user-only NTFS ACLs and we inherit them.
    # On POSIX, set mode 0700 so the audit log can't be world/group-read.
    if ($createdNow -and -not $onWindows -and (Get-Command chmod -ErrorAction SilentlyContinue)) {
        try { & chmod 700 $base 2>$null | Out-Null } catch {}
    }
    $name = "mark-{0}-{1}.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'), $PID
    $script:AIAuditLogPath = Join-Path $base $name
    return $script:AIAuditLogPath
}

function Format-CommandForAudit {
    param([string]$Command)
    # Redact -ParamName <value> for sensitive parameter names. Value can be
    # 'single-quoted', "double-quoted", or bareword/variable up to next whitespace.
    $sensitive = 'Password|Credential|AccessToken|ClientSecret|AppSecret|ApiKey|Secret|Token'
    return ($Command -replace "(-(?:$sensitive))\s+('[^']*'|""[^""]*""|\S+)", '$1 ***REDACTED***')
}

function Write-AIAuditEntry {
    param([string]$EventType, [string]$Detail)
    $path = Get-AIAuditLogPath
    if (-not $path) { return }
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $EventType, $Detail
    try { Add-Content -LiteralPath $path -Value $line -ErrorAction Stop } catch {}
}

function Format-DetailForAudit {
    <#
        Compose audit text. Always strips -Password/-Token/etc parameter
        values (Format-CommandForAudit). Additionally tokenizes PII via
        Convert-ToSafePayload when Privacy.RedactInAuditLog is Enabled.
        Caller may pass $Config=$null for contexts where no config is
        available (e.g. session-startup events).
    #>
    param([string]$Detail, [hashtable]$Config)
    if ([string]::IsNullOrEmpty($Detail)) { return $Detail }
    $piiRedact = $false
    if ($Config -and $Config.ContainsKey('Privacy') -and $Config['Privacy'] -is [hashtable]) {
        $piiRedact = ($Config['Privacy']['RedactInAuditLog'] -eq 'Enabled')
    }
    $out = $Detail
    if ($piiRedact) {
        # SecretsOnly=$false applies the full pattern set; secrets are
        # always tokenized regardless.
        $out = Convert-ToSafePayload -Text $out -SecretsOnly:$false -Counts $null
    }
    return (Format-CommandForAudit -Command $out)
}

function Test-AICommandAllowed {
    <#
        Parses the proposed command via the PowerShell language parser and
        rejects if (a) it has parse errors, (b) it contains no named
        CommandAst, (c) any CommandAst's name does not match a glob in
        $script:AICmdAllowList. Returns a hashtable:
            Allowed  : bool
            Reason   : string
            Commands : array of cmdlet names that were inspected
    #>
    param([string]$Command)

    $tokens = $null; $errs = $null
    try {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Command, [ref]$tokens, [ref]$errs)
    } catch {
        return @{ Allowed=$false; Reason="parser exception: $($_.Exception.Message)"; Commands=@() }
    }
    if ($errs -and $errs.Count -gt 0) {
        return @{ Allowed=$false; Reason="parse error: $($errs[0].Message)"; Commands=@() }
    }

    $cmdAsts = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
    if (-not $cmdAsts -or $cmdAsts.Count -eq 0) {
        return @{ Allowed=$false; Reason="no command invocation found"; Commands=@() }
    }

    $found = @()
    foreach ($c in $cmdAsts) {
        $name = $c.GetCommandName()
        if (-not $name) {
            return @{ Allowed=$false; Reason="indirect or non-named invocation rejected"; Commands=$found }
        }
        $found += $name
        $allowed = $false
        foreach ($pattern in $script:AICmdAllowList) {
            if ($name -like $pattern) { $allowed = $true; break }
        }
        if (-not $allowed) {
            return @{ Allowed=$false; Reason="cmdlet '$name' not in allow-list"; Commands=$found }
        }
    }
    return @{ Allowed=$true; Reason=""; Commands=$found }
}

# ============================================================
#  Command Execution with [Y/N/A/E] and auto-fixes
# ============================================================

function Invoke-MarkCommands {
    param([string]$Response, [hashtable]$Config)

    # Deprecation notice: this regex-RUN: path is now the FALLBACK for
    # providers / models that don't support native tool calling. Phase 5
    # Commit A's AIToolDispatch.ps1 + the ai-tools/ registry is the
    # primary path. We emit one console warning per session.
    if (-not $script:AIToolingDeprecationWarned) {
        Write-Warn "[deprecated] Using regex RUN: extractor. Native tool calling preferred (provider lacked support, or capability cache says no)."
        $script:AIToolingDeprecationWarned = $true
    }

    $commands = Extract-Commands -Response $Response
    if ($commands.Count -eq 0) { return $null }

    $results = @(); $runAll = $false

    foreach ($command in $commands) {
        $redacted = Format-DetailForAudit -Detail $command -Config $Config

        # ---- Default-deny: AST parse + allow-list (no Y/N prompt on reject) ----
        $check = Test-AICommandAllowed -Command $command
        if (-not $check.Allowed) {
            Write-AIAuditEntry -EventType "REJECT" -Detail "$($check.Reason) | $redacted"
            Write-Host ""
            Write-ErrorMsg "Refused to run AI command: $($check.Reason)"
            Write-Host "    $command" -ForegroundColor DarkGray
            $results += "Command: $command`nError: rejected ($($check.Reason))"
            continue
        }
        Write-AIAuditEntry -EventType "PROPOSE" -Detail $redacted

        if (-not (Ensure-ServiceForCommand -Command $command)) {
            Write-AIAuditEntry -EventType "ERROR" -Detail "service connection failed | $redacted"
            $results += "Command: $command`nError: Service connection failed."
            continue
        }

        # Display command box
        Write-Host ""
        $b = $script:Box
        Write-Host ("  " + $b.TL + [string]::new($b.H, 56) + $b.TR) -ForegroundColor $script:Colors.Warning
        Write-Host ("  " + $b.V + " Mark wants to run:") -ForegroundColor $script:Colors.Warning
        $dc = $command; $lw = 52
        while ($dc.Length -gt 0) {
            $ch = if($dc.Length -gt $lw){$dc.Substring(0,$lw)}else{$dc}
            Write-Host ("  " + $b.V + "   " + $ch) -ForegroundColor "Cyan"
            $dc = if($dc.Length -gt $lw){$dc.Substring($lw)}else{""}
        }
        Write-Host ("  " + $b.BL + [string]::new($b.H, 56) + $b.BR) -ForegroundColor $script:Colors.Warning

        if ($command -match '^\s*(Set-|Remove-|Add-|New-|Update-|Revoke-|Enable-|Disable-|Start-)') { Write-Warn "This will MAKE CHANGES." }

        # ---- Confirmation loop with [E]xplain ----
        if (-not $runAll) {
            $confirmed = $false
            while (-not $confirmed) {
                Write-Host "  [Y]es  [N]o  [A]ll  [E]xplain" -ForegroundColor $script:Colors.Highlight -NoNewline; Write-Host ": " -NoNewline
                $ans = Read-Host
                if ($ans -match '^[Aa]') { $runAll = $true; $confirmed = $true }
                elseif ($ans -match '^[Yy]') { $confirmed = $true }
                elseif ($ans -match '^[Ee]') {
                    Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-Host "explaining..." -ForegroundColor "DarkGray"
                    $expl = Invoke-AIChat -Config $Config -Messages @(@{role="user";content="Explain this PowerShell command simply. What does each part do?`n`n$command"})
                    try{$t=$Host.UI.RawUI.CursorPosition.Y-1;$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t;Write-Host(" "*80);$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t}catch{}
                    Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse (Get-CleanResponse $expl); Write-Host ""
                }
                else { $results += "[Skipped: $command]"; $confirmed = $true; $ans = "n" }
            }
            if ($ans -match '^[Nn]') {
                Write-AIAuditEntry -EventType "SKIP" -Detail $redacted
                continue
            }
        }

        # ---- Execute via parsed ScriptBlock (no Invoke-Expression) ----
        Write-Host "  Executing..." -ForegroundColor $script:Colors.Info
        Write-AIAuditEntry -EventType "RUN" -Detail $redacted
        try {
            $sb = [scriptblock]::Create($command)
            $output = & $sb 2>&1
            $outputStr = ($output | Out-String).Trim()
            if ([string]::IsNullOrWhiteSpace($outputStr)) { $outputStr = "(completed, no output)" }
            Write-Host "  Result:" -ForegroundColor $script:Colors.Success
            $lines = $outputStr -split "`n"; $show = [math]::Min($lines.Count, 25)
            for ($i = 0; $i -lt $show; $i++) { Write-Host "    $($lines[$i])" -ForegroundColor White }
            if ($lines.Count -gt 25) { Write-Warn "($($lines.Count) lines, showing 25)" }
            $results += "Command: $command`nOutput: $outputStr"
            Write-AIAuditEntry -EventType "OK" -Detail ("{0} line(s) | {1}" -f $lines.Count, $redacted)
        } catch {
            Write-ErrorMsg "Failed: $_"
            $results += "Command: $command`nError: $_"
            $errMsg = Format-DetailForAudit -Detail $_.Exception.Message -Config $Config
            Write-AIAuditEntry -EventType "ERROR" -Detail ("{0} | {1}" -f $errMsg, $redacted)
        }
    }

    return ($results -join "`n---`n")
}

# ============================================================
#  Chat Interface
# ============================================================

function Start-AIAssistant {
    $b = $script:Box
    $config = Get-AIConfig
    if ($null -eq $config) {
        Write-Host ""; Write-Host ("  " + $b.DTL + [string]::new($b.DH,56) + $b.DTR) -ForegroundColor $script:Colors.Title
        Write-Host ("  " + $b.DV + "   Mark - AI Assistant (First Time Setup)              " + $b.DV) -ForegroundColor $script:Colors.Title
        Write-Host ("  " + $b.DBL + [string]::new($b.DH,56) + $b.DBR) -ForegroundColor $script:Colors.Title
        Write-Host ""; $config = Setup-AIProvider; if ($null -eq $config) { Write-ErrorMsg "Cancelled."; Pause-ForUser; return }
    }
    if ($config -is [PSCustomObject]) { $ht=@{}; $config.PSObject.Properties|ForEach-Object{$ht[$_.Name]=$_.Value}; $config=$ht }

    Clear-Host; Write-Host ""
    Write-Host ("  " + $b.DTL + [string]::new($b.DH,56) + $b.DTR) -ForegroundColor "Cyan"
    Write-Host ("  " + $b.DV + "   Mark - M365 AI Assistant                            " + $b.DV) -ForegroundColor "Cyan"
    $pd="$($config['Provider']) ($($config['Model']))"; if($pd.Length -gt 52){$pd=$pd.Substring(0,49)+"..."}
    Write-Host ("  " + $b.DV + "   $("{0,-52}" -f $pd)" + $b.DV) -ForegroundColor "Gray"
    Write-Host ("  " + $b.DBL + [string]::new($b.DH,56) + $b.DBR) -ForegroundColor "Cyan"
    Write-Host ""; Write-Host "  /help  /about  /tenants  /tenant <name>  /tools  /plan  /noplan  /dryrun  /cost  /costs  /list  /load  /save  /quit" -ForegroundColor "DarkGray"; Write-Host ""

    # Phase 5: pre-load the tool catalog and probe provider capability.
    if (Get-Command Get-AIToolCatalog -ErrorAction SilentlyContinue) { Get-AIToolCatalog | Out-Null }
    $useTooling = $false
    if (Get-Command Test-ProviderToolSupport -ErrorAction SilentlyContinue) {
        $useTooling = Test-ProviderToolSupport -Config $config
    }
    if ($useTooling) {
        Write-InfoMsg ("Native tool calling enabled ({0} tools loaded)." -f ($script:AIToolCatalog | Measure-Object).Count)
    } else {
        Write-Warn "Native tool calling unavailable -- falling back to RUN: regex path."
    }
    Write-Host ""

    $chatHistory = @()
    $ctx = "Services auto-connect. Tenant: $($script:SessionState.TenantMode)"
    if ($script:SessionState.TenantName -and $script:SessionState.TenantName -ne "Own Tenant") { $ctx += " ($($script:SessionState.TenantName))" }

    Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline
    Write-MarkResponse "Hey! I'm Mark. How can I help?"
    Write-Host ""

    $chatting = $true
    while ($chatting) {
        Write-Host "  You" -ForegroundColor $script:Colors.Highlight -NoNewline; Write-Host ": " -NoNewline
        $userMsg = Read-Host
        if ([string]::IsNullOrWhiteSpace($userMsg)) { continue }

        $cmd = $userMsg.Trim().ToLower()
        if ($cmd -match '^/(quit|exit|back)$') {
            # Auto-save unless /ephemeral
            if ((Get-Command Save-AISession -ErrorAction SilentlyContinue) -and $chatHistory.Count -gt 0 -and -not (Get-AISessionCurrent).Ephemeral) {
                $sid = Save-AISession -Config $config -History $chatHistory
                if ($sid) { Write-Host "  [auto-saved as $sid]" -ForegroundColor DarkGray }
            }
            Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": See you!" -ForegroundColor White; Write-Host ""; $chatting=$false; continue
        }
        if ($cmd -eq '/help') {
            Write-Host ""
            Write-Host "  /config /models /privacy /clear /context /tools" -ForegroundColor "DarkGray"
            Write-Host "  /about                 diagnostic snapshot" -ForegroundColor "DarkGray"
            Write-Host "  /tenants               list registered tenants" -ForegroundColor "DarkGray"
            Write-Host "  /tenant <name>         switch tenant context" -ForegroundColor "DarkGray"
            Write-Host "  /plan /noplan          plan-mode controls" -ForegroundColor "DarkGray"
            Write-Host "  /dryrun                toggle PREVIEW (no tenant changes)" -ForegroundColor "DarkGray"
            Write-Host "  /cost /costs           cost summary / history" -ForegroundColor "DarkGray"
            Write-Host "  /list /load <id>       saved sessions" -ForegroundColor "DarkGray"
            Write-Host "  /save [title]          persist this chat" -ForegroundColor "DarkGray"
            Write-Host "  /rename <id> <title>   rename" -ForegroundColor "DarkGray"
            Write-Host "  /delete <id>           delete" -ForegroundColor "DarkGray"
            Write-Host "  /ephemeral             do not auto-save on quit" -ForegroundColor "DarkGray"
            Write-Host "  /export <id> [path]    redacted JSON for sharing" -ForegroundColor "DarkGray"
            Write-Host "  /quit                  exit" -ForegroundColor "DarkGray"
            Write-Host ""; continue
        }
        if ($cmd -eq '/cost') {
            if (Get-Command Show-AICostSummary -ErrorAction SilentlyContinue) { Show-AICostSummary }
            else { Write-Warn "Cost tracker not loaded." }
            continue
        }
        if ($cmd -eq '/costs') {
            if (Get-Command Show-AICostHistory -ErrorAction SilentlyContinue) { Show-AICostHistory }
            else { Write-Warn "Cost tracker not loaded." }
            continue
        }
        if ($cmd -eq '/list') {
            if (Get-Command Show-AISessionList -ErrorAction SilentlyContinue) { Show-AISessionList }
            else { Write-Warn "Session store not loaded." }
            continue
        }
        if ($cmd -like '/load *') {
            if (-not (Get-Command Load-AISession -ErrorAction SilentlyContinue)) { Write-Warn "Session store not loaded."; continue }
            $target = $userMsg.Substring(6).Trim()
            $loaded = Load-AISession -IdOrPrefix $target
            if (-not $loaded) { Write-Warn "No session matches '$target'."; continue }
            $chatHistory = @($loaded.History)
            Write-InfoMsg ("Loaded session '{0}' ({1} message(s), {2:N4} USD)." -f $loaded.Title, $loaded.History.Count, $loaded.CostUsd)
            continue
        }
        if ($cmd -eq '/save' -or $cmd -like '/save *') {
            if (-not (Get-Command Save-AISession -ErrorAction SilentlyContinue)) { Write-Warn "Session store not loaded."; continue }
            $title = if ($cmd.Length -gt 5) { $userMsg.Substring(6).Trim() } else { $null }
            $sid = Save-AISession -Config $config -History $chatHistory -Title $title -Force
            if ($sid) { Write-InfoMsg ("Saved session '{0}' ({1})." -f $script:AISessionCurrent.Title, $sid) }
            continue
        }
        if ($cmd -like '/rename *') {
            $parts = $userMsg.Substring(8).Trim() -split '\s+', 2
            if ($parts.Count -ne 2) { Write-Warn "Usage: /rename <id-or-prefix> <new title>"; continue }
            if (Rename-AISession -IdOrPrefix $parts[0] -NewTitle $parts[1]) { Write-InfoMsg "Renamed." }
            continue
        }
        if ($cmd -like '/delete *') {
            $target = $userMsg.Substring(8).Trim()
            if (Remove-AISession -IdOrPrefix $target) { Write-InfoMsg "Deleted '$target'." }
            continue
        }
        if ($cmd -eq '/ephemeral') {
            Set-AISessionEphemeral -On $true
            Write-InfoMsg "Ephemeral mode ON -- this chat will NOT be auto-saved on /quit."
            continue
        }
        if ($cmd -eq '/about') {
            if (Get-Command Show-AIAbout -ErrorAction SilentlyContinue) { Show-AIAbout -Config $config }
            else { Write-Warn "AIUx module not loaded." }
            continue
        }
        if ($cmd -eq '/tenants') {
            if (Get-Command Show-TenantRegistry -ErrorAction SilentlyContinue) { Show-TenantRegistry }
            else { Write-Warn "Tenant registry not loaded." }
            continue
        }
        if ($cmd -like '/tenant *') {
            if (-not (Get-Command Switch-Tenant -ErrorAction SilentlyContinue)) { Write-Warn "TenantSwitch not loaded."; continue }
            $name = $userMsg.Substring(8).Trim()
            if (-not $name) { Write-Warn "Usage: /tenant <name>"; continue }
            Switch-Tenant -Name $name | Out-Null
            continue
        }
        if ($cmd -eq '/dryrun') {
            if (Get-Command Set-PreviewMode -ErrorAction SilentlyContinue) {
                $newState = -not (Get-PreviewMode)
                Set-PreviewMode -Enabled $newState
                if ($newState) { Write-Warn "PREVIEW MODE ON -- mutating tools will be logged but not executed." }
                else           { Write-InfoMsg "PREVIEW MODE OFF -- mutating tools will apply to the tenant." }
            } else { Write-Warn "Preview module not loaded." }
            continue
        }
        if ($cmd -like '/export *' -or $cmd -eq '/export') {
            $rest = if ($cmd.Length -gt 8) { $userMsg.Substring(8).Trim() } else { '' }
            $parts = $rest -split '\s+', 2
            $idArg = if ($parts.Count -ge 1 -and $parts[0]) { $parts[0] } else { $script:AISessionCurrent.Id }
            $dest  = if ($parts.Count -ge 2) { $parts[1] } else { $null }
            if (-not $idArg) { Write-Warn "Usage: /export <id-or-prefix> [destination-path]"; continue }
            Export-AISession -IdOrPrefix $idArg -DestinationPath $dest | Out-Null
            continue
        }
        if ($cmd -eq '/plan') {
            if (Get-Command Set-AIPlanMode -ErrorAction SilentlyContinue) {
                Set-AIPlanMode -Mode 'force'
                Write-InfoMsg "Plan mode forced for next prompt -- Mark will submit_plan before executing tools."
            } else { Write-Warn "Planner module not loaded." }
            Write-Host ""; continue
        }
        if ($cmd -eq '/noplan') {
            if (Get-Command Set-AIPlanMode -ErrorAction SilentlyContinue) {
                Set-AIPlanMode -Mode 'skip'
                Write-InfoMsg "Plan mode skipped for next prompt -- Mark will call tools directly."
            } else { Write-Warn "Planner module not loaded." }
            Write-Host ""; continue
        }
        if ($cmd -eq '/config') { $nc=Setup-AIProvider; if($nc){$config=$nc}else{$r=Get-AIConfig;if($r){if($r -is [PSCustomObject]){$ht=@{};$r.PSObject.Properties|ForEach-Object{$ht[$_.Name]=$_.Value};$config=$ht}else{$config=$r}}}; Write-Host ""; continue }
        if ($cmd -eq '/models') { if($config["Provider"] -eq "Ollama"){try{$ms=Invoke-RestMethod -Uri "$($config['Endpoint'])/api/tags" -Method GET -TimeoutSec 5;Write-Host "";foreach($m in $ms.models){$c=if($m.name -eq $config["Model"]){"<< current"}else{""};Write-Host "    $($m.name) ($([math]::Round($m.size/1MB))MB) $c" -ForegroundColor White}}catch{Write-ErrorMsg "Cannot reach Ollama."}}else{Write-InfoMsg "/models is for Ollama."}; Write-Host ""; continue }
        if ($cmd -eq '/privacy') { Show-PrivacyMenu -Config $config; $r=Get-AIConfig; if($r){$config=$r}; Write-Host ""; continue }
        if ($cmd -eq '/clear') {
            $cleared = Reset-PrivacyMap
            $chatHistory = @()
            Write-AIAuditEntry -EventType "CLEAR" -Detail ("chat history + privacy map ({0} tokens) cleared" -f $cleared)
            Write-Host "  [cleared — $cleared privacy token(s) dropped]" -ForegroundColor "DarkGray"; Write-Host ""; continue
        }
        if ($cmd -eq '/context') { Write-Host ""; Write-StatusLine "Tenant" "$($script:SessionState.TenantMode) $(if($script:SessionState.TenantName){"($($script:SessionState.TenantName))"})" "White"; Write-StatusLine "Graph" $(if($script:SessionState.MgGraph){"Connected"}else{"Auto"}) $(if($script:SessionState.MgGraph){"Green"}else{"Yellow"}); Write-StatusLine "EXO" $(if($script:SessionState.ExchangeOnline){"Connected"}else{"Auto"}) $(if($script:SessionState.ExchangeOnline){"Green"}else{"Yellow"}); Write-StatusLine "SCC" $(if($script:SessionState.ComplianceCenter){"Connected"}else{"Auto"}) $(if($script:SessionState.ComplianceCenter){"Green"}else{"Yellow"}); Write-StatusLine "AI" "$($config['Provider'])/$($config['Model'])" "Cyan"; Write-Host ""; continue }
        if ($cmd -eq '/tools') {
            if (Get-Command Show-AIToolCatalog -ErrorAction SilentlyContinue) { Show-AIToolCatalog }
            else { Write-Warn "Tool catalog module not loaded." }
            continue
        }

        # ---- Phase 5: native tool-calling path ----
        if ($useTooling -and (Get-Command Invoke-AIChatToolingTurn -ErrorAction SilentlyContinue)) {
            $sendMsg = $userMsg
            if ($chatHistory.Count -eq 0) { $sendMsg = "[CONTEXT: $ctx]`n`n$userMsg" }
            $chatHistory += @{ role = 'user'; content = $sendMsg }
            $hops = 0
            $maxHops = 8
            while ($hops -lt $maxHops) {
                $hops++
                Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-Host "thinking..." -ForegroundColor "DarkGray"
                $turn = Invoke-AIChatToolingTurn -Config $config -Messages $chatHistory
                try{$t=$Host.UI.RawUI.CursorPosition.Y-1;$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t;Write-Host(" "*80);$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t}catch{}

                if ($turn.Error) {
                    if ($turn.Error -eq 'ollama_no_tool_support') {
                        Write-Warn "This Ollama model lacks tool support -- switching to RUN: regex fallback for this session."
                        $useTooling = $false
                        # Pop the user message we already added so the regex path takes it fresh
                        $chatHistory = @($chatHistory | Select-Object -SkipLast 1)
                        break
                    }
                    Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": [!] $($turn.Error)" -ForegroundColor $script:Colors.Error
                    break
                }
                # ---- Cost tracking for this hop ----
                if ($turn.Usage -and (Get-Command Add-AICostEvent -ErrorAction SilentlyContinue)) {
                    try {
                        $costResult = Add-AICostEvent -Config $config -Usage $turn.Usage -Reason 'tooling-hop'
                        if (Get-Command Show-AICostFooter -ErrorAction SilentlyContinue) { Show-AICostFooter -Result $costResult }
                    } catch { Write-Warn "Cost tracker: $($_.Exception.Message)" }
                }
                # Persist assistant turn into history in Anthropic-shaped form
                # so subsequent tool_result blocks have somewhere to attach.
                if ($turn.AssistantContent) {
                    $chatHistory += @{ role = 'assistant'; content = $turn.AssistantContent }
                }
                if ($turn.Text) {
                    Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline
                    Write-MarkResponse $turn.Text
                }
                if ($turn.ToolUses.Count -eq 0) { break }
                # Operator gate (Y/A/N) + dispatch
                $results = New-Object System.Collections.ArrayList
                $runAll = $false
                # ---- Auto-plan: if AI returned >= threshold tool_uses on first hop
                # and didn't already submit_plan, nudge into plan mode for the *next*
                # turn. We don't synthesize a plan locally; we tell the AI to revise.
                $threshold = if ($script:AIAutoPlanThreshold) { [int]$script:AIAutoPlanThreshold } else { 3 }
                $hasSubmitPlan = $false
                foreach ($tu in $turn.ToolUses) { if ($tu.name -eq 'submit_plan') { $hasSubmitPlan = $true; break } }
                if ($hops -eq 1 -and -not $hasSubmitPlan -and $turn.ToolUses.Count -ge $threshold -and (Get-Command Set-AIPlanMode -ErrorAction SilentlyContinue) -and (Get-AIPlanMode) -ne 'skip') {
                    Write-Warn ("Mark proposed {0} tool calls on first turn (>= {1}). Routing through plan mode." -f $turn.ToolUses.Count, $threshold)
                    Set-AIPlanMode -Mode 'force'
                    foreach ($tu in $turn.ToolUses) {
                        [void]$results.Add(@{ id = $tu.id; content = '{"ok":false,"error":"auto_plan_required","note":"Operator requires submit_plan first for tasks needing >= ' + $threshold + ' tool calls."}'; isError = $true })
                    }
                } else {
                foreach ($tu in $turn.ToolUses) {
                    # ---- submit_plan intercept -- route to planner approval flow
                    if ($tu.name -eq 'submit_plan' -and (Get-Command Invoke-AIPlanApprovalFlow -ErrorAction SilentlyContinue)) {
                        Write-Host ""
                        $flow = Invoke-AIPlanApprovalFlow -PlanInput $tu.input -Config $config
                        # Reset plan-mode latch once a plan has been processed
                        if (Get-Command Set-AIPlanMode -ErrorAction SilentlyContinue) { Set-AIPlanMode -Mode 'auto' }
                        $body = @{ ok = ($flow.Status -ne 'rejected'); status = $flow.Status }
                        if ($flow.Result) {
                            $body.executed  = $flow.Result.Executed
                            $body.succeeded = $flow.Result.Succeeded
                            $body.failed    = $flow.Result.Failed
                            $body.skipped   = $flow.Result.Skipped
                            $body.steps     = @($flow.Result.StepResults | ForEach-Object {
                                @{ id = $_.id; tool = $_.tool; status = $_.status; error = $_.error }
                            })
                        }
                        if ($flow.ValidationErrors) { $body.validationErrors = @($flow.ValidationErrors) }
                        if ($flow.Reason) { $body.reason = $flow.Reason }
                        $payload = ($body | ConvertTo-Json -Depth 8 -Compress)
                        [void]$results.Add(@{ id = $tu.id; content = $payload; isError = ($flow.Status -eq 'rejected') })
                        continue
                    }
                    # ---- ask_operator intercept -- prompt the human directly
                    if ($tu.name -eq 'ask_operator') {
                        $q = [string]$tu.input.question
                        Write-Host ""
                        Write-Host "  Mark asks: $q" -ForegroundColor Cyan
                        Write-Host "  You" -ForegroundColor $script:Colors.Highlight -NoNewline; Write-Host ": " -NoNewline
                        $ansText = Read-Host
                        $payload = (@{ ok = $true; answer = $ansText } | ConvertTo-Json -Compress)
                        [void]$results.Add(@{ id = $tu.id; content = $payload; isError = $false })
                        continue
                    }
                    $toolDef = Get-AIToolByName -Name $tu.name
                    $destFlag = if ($toolDef -and $toolDef.destructive) { '[DESTRUCTIVE] ' } else { '' }
                    Write-Host ""
                    Write-Host ("  Mark wants to call: " + $destFlag + $tu.name) -ForegroundColor Yellow
                    foreach ($k in $tu.input.Keys) {
                        Write-Host ("      {0}: {1}" -f $k, ($tu.input[$k])) -ForegroundColor DarkGray
                    }
                    $ans = if ($runAll) { 'y' } else {
                        Write-Host "  [Y]es  [N]o  [A]ll  [Q]uit" -ForegroundColor $script:Colors.Highlight -NoNewline; Write-Host ": " -NoNewline
                        Read-Host
                    }
                    if ($ans -match '^[Aa]') { $runAll = $true; $ans = 'y' }
                    if ($ans -match '^[Qq]') { break }
                    if ($ans -notmatch '^[Yy]') {
                        # Ask for an optional reason so the AI gets context to adapt.
                        $note = if (Get-NonInteractiveMode) { '' } else {
                            Write-Host "  (optional) what should change? " -ForegroundColor DarkGray -NoNewline
                            Read-Host
                        }
                        $rejectPayload = if (Get-Command Build-RejectionToolResult -ErrorAction SilentlyContinue) {
                            Build-RejectionToolResult -ToolName $tu.name -Note $note
                        } else { '{"ok":false,"error":"operator_declined"}' }
                        [void]$results.Add(@{ id = $tu.id; content = $rejectPayload; isError = $true })
                        continue
                    }
                    $inputHt = @{}; foreach ($k in $tu.input.Keys) { $inputHt[$k] = $tu.input[$k] }
                    $spin = $null
                    if (Get-Command Start-AISpinner -ErrorAction SilentlyContinue) {
                        $spin = Start-AISpinner -Label ("running " + $tu.name)
                    }
                    try {
                        $out = Invoke-AIToolCall -ToolName $tu.name -Params $inputHt -Config $config
                    } finally {
                        if ($spin -and (Get-Command Stop-AISpinner -ErrorAction SilentlyContinue)) { Stop-AISpinner -Token $spin }
                    }
                    $payload = (Format-AIToolResultPayload -Result $out -MaxBytes 4000)
                    [void]$results.Add(@{ id = $tu.id; content = $payload; isError = (-not $out.ok) })
                }
                }
                if ($results.Count -eq 0) { break }
                $followup = Build-ToolResultMessage -Provider $config.Provider -ToolResults @($results)
                if ($followup -is [array]) { $chatHistory += $followup } else { $chatHistory += $followup }
            }
            Write-Host ""
            if ($chatHistory.Count -gt 60) { $chatHistory = $chatHistory[-60..-1] }
            # Reset the /plan or /noplan latch so it only applies to one prompt.
            if (Get-Command Set-AIPlanMode -ErrorAction SilentlyContinue) { Set-AIPlanMode -Mode 'auto' }
            if ($useTooling) { continue }   # done; let outer while-loop ask next prompt
        }

        $sendMsg = $userMsg
        if ($chatHistory.Count -eq 0) { $sendMsg = "[CONTEXT: $ctx]`n`n$userMsg" }
        $chatHistory += @{ role = "user"; content = $sendMsg }

        # ---- Call AI ----
        Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-Host "thinking..." -ForegroundColor "DarkGray"
        $response = Invoke-AIChat -Config $config -Messages $chatHistory
        try{$t=$Host.UI.RawUI.CursorPosition.Y-1;$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t;Write-Host(" "*80);$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t}catch{}

        if ($response -match "^\[!\]") { Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": $response" -ForegroundColor $script:Colors.Error; Write-Host ""; continue }

        # ---- Cost footer (regex path) ----
        if ($script:LastAICostResult -and (Get-Command Show-AICostFooter -ErrorAction SilentlyContinue)) {
            Show-AICostFooter -Result $script:LastAICostResult
            $script:LastAICostResult = $null
        }

        $hasCmd = Test-HasCommands -Response $response
        $cleanText = Get-CleanResponse $response

        # Show only the clean (non-hallucinated) text -- skip when we
        # already streamed it (Ollama path printed inline).
        if ($cleanText -and -not $script:LastAIChatStreamed) {
            Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse $cleanText
        }
        $script:LastAIChatStreamed = $false

        if ($hasCmd) {
            $cmdResults = Invoke-MarkCommands -Response $response -Config $config
            $chatHistory += @{ role = "assistant"; content = $response }

            if ($cmdResults) {
                # Truncate results to avoid overwhelming the AI context
                $truncated = Limit-ResultsForAI -Text $cmdResults
                $chatHistory += @{ role = "user"; content = "[COMMAND RESULTS]`n$truncated`n[/COMMAND RESULTS]`nBriefly tell me what happened. If there were multiple matches for a user, ask which one. If more steps are needed, do them with RUN: commands." }
                Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-Host "interpreting..." -ForegroundColor "DarkGray"
                $interp = Invoke-AIChat -Config $config -Messages $chatHistory
                try{$t=$Host.UI.RawUI.CursorPosition.Y-1;$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t;Write-Host(" "*80);$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t}catch{}

                $ci = Get-CleanResponse $interp
                if ($ci) { Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse $ci }

                # Execute follow-up commands
                if (Test-HasCommands -Response $interp) {
                    $more = Invoke-MarkCommands -Response $interp -Config $config
                    $chatHistory += @{role="assistant";content=$interp}
                    if ($more) {
                        $moreT = Limit-ResultsForAI -Text $more
                        $chatHistory += @{role="user";content="[COMMAND RESULTS]`n$moreT`n[/COMMAND RESULTS]`nBriefly summarize. If more steps needed, do them."}
                        $final = Invoke-AIChat -Config $config -Messages $chatHistory
                        $cf = Get-CleanResponse $final
                        if ($cf) { Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse $cf }
                        $chatHistory += @{role="assistant";content=$final}
                    }
                } else { $chatHistory += @{role="assistant";content=$interp} }
            }
        } else {
            if (-not $cleanText) { Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse $response }
            $chatHistory += @{role="assistant";content=$response}
        }
        Write-Host ""
        if ($chatHistory.Count -gt 40) { $chatHistory=$chatHistory[-40..-1] }
    }
}

function Write-MarkResponse {
    param([string]$Text)
    $maxW=72;$ind="        ";$first=$true
    foreach($para in ($Text -split "`n")){$para=$para.Trim();if([string]::IsNullOrWhiteSpace($para)){Write-Host "";continue}
        if($para -match '^\s*([-*]|\d+\.|```)'){if($first){Write-Host $para -ForegroundColor White;$first=$false}else{Write-Host "$ind$para" -ForegroundColor White};continue}
        $words=$para -split '\s+';$cur=""
        foreach($w in $words){if(($cur.Length+$w.Length+1) -gt $maxW){if($first){Write-Host $cur -ForegroundColor White;$first=$false}else{Write-Host "$ind$cur" -ForegroundColor White};$cur=$w}else{if($cur){$cur+=" $w"}else{$cur=$w}}}
        if($cur){if($first){Write-Host $cur -ForegroundColor White;$first=$false}else{Write-Host "$ind$cur" -ForegroundColor White}}
    }
}
