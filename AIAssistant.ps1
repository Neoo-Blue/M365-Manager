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

# ============================================================
#  API Call
# ============================================================

function Invoke-AIChat {
    param([hashtable]$Config,[array]$Messages)
    $p=$Config["Provider"]
    try { $k = Unprotect-ApiKey -StoredKey ([string]$Config["ApiKey"]) }
    catch { return "[!] $_" }
    $ep=$Config["Endpoint"];$m=$Config["Model"]
    try {
        switch($p){
            "Ollama" { $body=@{model=$m;stream=$false;messages=@(@{role="system";content=$script:AISystemPrompt})+$Messages}|ConvertTo-Json -Depth 10; return (Invoke-RestMethod -Uri "$ep/api/chat" -Method POST -ContentType "application/json" -Body $body -TimeoutSec 180 -ErrorAction Stop).message.content }
            "Anthropic" { $body=@{model=$m;max_tokens=4096;system=$script:AISystemPrompt;messages=$Messages}|ConvertTo-Json -Depth 10; $h=@{"x-api-key"=$k;"anthropic-version"="2023-06-01";"content-type"="application/json"}; return (Invoke-RestMethod -Uri $ep -Method POST -Headers $h -Body $body -ErrorAction Stop).content[0].text }
            default { $body=@{model=$m;messages=@(@{role="system";content=$script:AISystemPrompt})+$Messages;max_tokens=4096}|ConvertTo-Json -Depth 10; $h=@{"Content-Type"="application/json"}; if($p -eq "AzureOpenAI"){$h["api-key"]=$k}else{$h["Authorization"]="Bearer $k"}; return (Invoke-RestMethod -Uri $ep -Method POST -Headers $h -Body $body -ErrorAction Stop).choices[0].message.content }
        }
    } catch { $e=$_.Exception.Message; if($e -match "401|Unauth"){return "[!] Invalid API key."}; if($e -match "429"){return "[!] Rate limited."}; if($e -match "resolve|connect|refused"){return "[!] Cannot reach $(if($p -eq 'Ollama'){'Ollama'}else{'API'})."}; return "[!] API error: $e" }
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
#  Command Execution with [Y/N/A/E] and auto-fixes
# ============================================================

function Invoke-MarkCommands {
    param([string]$Response, [hashtable]$Config)

    $commands = Extract-Commands -Response $Response
    if ($commands.Count -eq 0) { return $null }

    $results = @(); $runAll = $false

    foreach ($command in $commands) {
        if (-not (Ensure-ServiceForCommand -Command $command)) { $results += "Command: $command`nError: Service connection failed."; continue }

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
            if ($ans -match '^[Nn]') { continue }
        }

        # ---- Execute ----
        Write-Host "  Executing..." -ForegroundColor $script:Colors.Info
        try {
            $output = Invoke-Expression $command 2>&1
            $outputStr = ($output | Out-String).Trim()
            if ([string]::IsNullOrWhiteSpace($outputStr)) { $outputStr = "(completed, no output)" }
            Write-Host "  Result:" -ForegroundColor $script:Colors.Success
            $lines = $outputStr -split "`n"; $show = [math]::Min($lines.Count, 25)
            for ($i = 0; $i -lt $show; $i++) { Write-Host "    $($lines[$i])" -ForegroundColor White }
            if ($lines.Count -gt 25) { Write-Warn "($($lines.Count) lines, showing 25)" }
            $results += "Command: $command`nOutput: $outputStr"
        } catch {
            Write-ErrorMsg "Failed: $_"
            $results += "Command: $command`nError: $_"
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
    Write-Host ""; Write-Host "  /help  /config  /models  /clear  /context  /quit" -ForegroundColor "DarkGray"; Write-Host ""

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
        if ($cmd -match '^/(quit|exit|back)$') { Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": See you!" -ForegroundColor White; Write-Host ""; $chatting=$false; continue }
        if ($cmd -eq '/help') { Write-Host ""; Write-Host "  /config /models /clear /context /quit" -ForegroundColor "DarkGray"; Write-Host ""; continue }
        if ($cmd -eq '/config') { $nc=Setup-AIProvider; if($nc){$config=$nc}else{$r=Get-AIConfig;if($r){if($r -is [PSCustomObject]){$ht=@{};$r.PSObject.Properties|ForEach-Object{$ht[$_.Name]=$_.Value};$config=$ht}else{$config=$r}}}; Write-Host ""; continue }
        if ($cmd -eq '/models') { if($config["Provider"] -eq "Ollama"){try{$ms=Invoke-RestMethod -Uri "$($config['Endpoint'])/api/tags" -Method GET -TimeoutSec 5;Write-Host "";foreach($m in $ms.models){$c=if($m.name -eq $config["Model"]){"<< current"}else{""};Write-Host "    $($m.name) ($([math]::Round($m.size/1MB))MB) $c" -ForegroundColor White}}catch{Write-ErrorMsg "Cannot reach Ollama."}}else{Write-InfoMsg "/models is for Ollama."}; Write-Host ""; continue }
        if ($cmd -eq '/clear') { $chatHistory=@(); Write-Host "  [cleared]" -ForegroundColor "DarkGray"; Write-Host ""; continue }
        if ($cmd -eq '/context') { Write-Host ""; Write-StatusLine "Tenant" "$($script:SessionState.TenantMode) $(if($script:SessionState.TenantName){"($($script:SessionState.TenantName))"})" "White"; Write-StatusLine "Graph" $(if($script:SessionState.MgGraph){"Connected"}else{"Auto"}) $(if($script:SessionState.MgGraph){"Green"}else{"Yellow"}); Write-StatusLine "EXO" $(if($script:SessionState.ExchangeOnline){"Connected"}else{"Auto"}) $(if($script:SessionState.ExchangeOnline){"Green"}else{"Yellow"}); Write-StatusLine "SCC" $(if($script:SessionState.ComplianceCenter){"Connected"}else{"Auto"}) $(if($script:SessionState.ComplianceCenter){"Green"}else{"Yellow"}); Write-StatusLine "AI" "$($config['Provider'])/$($config['Model'])" "Cyan"; Write-Host ""; continue }

        $sendMsg = $userMsg
        if ($chatHistory.Count -eq 0) { $sendMsg = "[CONTEXT: $ctx]`n`n$userMsg" }
        $chatHistory += @{ role = "user"; content = $sendMsg }

        # ---- Call AI ----
        Write-Host ""; Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-Host "thinking..." -ForegroundColor "DarkGray"
        $response = Invoke-AIChat -Config $config -Messages $chatHistory
        try{$t=$Host.UI.RawUI.CursorPosition.Y-1;$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t;Write-Host(" "*80);$Host.UI.RawUI.CursorPosition=New-Object System.Management.Automation.Host.Coordinates 0,$t}catch{}

        if ($response -match "^\[!\]") { Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": $response" -ForegroundColor $script:Colors.Error; Write-Host ""; continue }

        $hasCmd = Test-HasCommands -Response $response
        $cleanText = Get-CleanResponse $response

        # Show only the clean (non-hallucinated) text
        if ($cleanText) { Write-Host "  Mark" -ForegroundColor "Cyan" -NoNewline; Write-Host ": " -NoNewline; Write-MarkResponse $cleanText }

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
