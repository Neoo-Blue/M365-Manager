# ============================================================
#  AISessionStore.ps1 -- persistent chat session storage
#
#  Each named session is one DPAPI-encrypted JSON blob under
#  <stateDir>/chat-sessions/<id>.session containing:
#    - title, timestamps, provider, model, tenant
#    - privacyMap (ByValue / ByToken / Counters)
#    - history (full chat message array, including tool-use /
#      tool-result blocks)
#    - costRolledUpUsd snapshot at save time
#  An unencrypted index.json lists session metadata so /list
#  doesn't have to decrypt anything.
#
#  /save and /load wire into Start-AIAssistant -- a chat is
#  auto-saved on /quit unless /ephemeral was toggled on. /export
#  writes a redacted plaintext version (placeholders, not real
#  values) for sharing.
# ============================================================

if ($null -eq (Get-Variable -Name AISessionCurrent -Scope Script -ErrorAction SilentlyContinue)) {
    $script:AISessionCurrent = @{
        Id             = $null
        Title          = $null
        Ephemeral      = $false   # set by /ephemeral; suppresses auto-save
        DirtySinceSave = $false
    }
}

function Get-AISessionDir {
    $base = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'M365Manager' } else { Join-Path $HOME '.m365manager' }
    $p = Join-Path $base 'chat-sessions'
    if (-not (Test-Path $p)) {
        New-Item -Path $p -ItemType Directory -Force | Out-Null
        if ($IsLinux -or $IsMacOS) { try { & chmod 700 $p 2>$null } catch {} }
    }
    return $p
}

function Get-AISessionIndexPath { return (Join-Path (Get-AISessionDir) 'index.json') }

function Get-AISessionIndex {
    $p = Get-AISessionIndexPath
    if (-not (Test-Path $p)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        return @($raw)
    } catch { return @() }
}

function Save-AISessionIndex {
    param([Parameter(Mandatory)][array]$Index)
    Set-Content -LiteralPath (Get-AISessionIndexPath) -Value ($Index | ConvertTo-Json -Depth 6) -Encoding UTF8 -Force
}

function New-AISessionId {
    return ((Get-Date).ToString('yyyyMMdd-HHmmss') + '-' + ([guid]::NewGuid().ToString().Substring(0,8)))
}

function Protect-AISessionPayload {
    <#
        DPAPI-encrypt (CurrentUser scope) on Windows. On non-Windows
        we write plain text and stamp 'plaintext' so /load knows not
        to attempt decryption. The operator gets a warning at save
        time on non-Windows so they can choose to /ephemeral instead.
    #>
    param([Parameter(Mandatory)][string]$PlainText)
    if ($IsLinux -or $IsMacOS) {
        return @{ Mode = 'plaintext'; Data = $PlainText }
    }
    try {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        $enc   = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return @{ Mode = 'dpapi'; Data = [Convert]::ToBase64String($enc) }
    } catch {
        Write-Warn "DPAPI encryption failed ($($_.Exception.Message)); falling back to plaintext."
        return @{ Mode = 'plaintext'; Data = $PlainText }
    }
}

function Unprotect-AISessionPayload {
    param(
        [Parameter(Mandatory)][string]$Mode,
        [Parameter(Mandatory)][string]$Data
    )
    if ($Mode -eq 'plaintext') { return $Data }
    if ($Mode -eq 'dpapi') {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $bytes = [Convert]::FromBase64String($Data)
        $dec   = [System.Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($dec)
    }
    throw "Unknown payload mode '$Mode'."
}

function Save-AISession {
    <#
        Persist the current chat. Title defaults to the first user
        message's leading 60 chars when not provided. Returns the
        session id. Does nothing when /ephemeral is on (and the
        operator did not explicitly type /save).
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Config,
        [Parameter(Mandatory)][array]$History,
        [string]$Title,
        [string]$Id,
        [switch]$Force
    )
    if ($script:AISessionCurrent.Ephemeral -and -not $Force) { return $null }
    if ($null -eq $History -or $History.Count -eq 0) { return $null }

    if (-not $Id) { $Id = if ($script:AISessionCurrent.Id) { $script:AISessionCurrent.Id } else { New-AISessionId } }
    if (-not $Title) {
        $first = $History | Where-Object { $_.role -eq 'user' } | Select-Object -First 1
        if ($first -and $first.content -is [string]) {
            $t = $first.content -replace '\s+', ' '
            $t = $t.Trim()
            if ($t.Length -gt 60) { $t = $t.Substring(0, 57) + '...' }
            $Title = $t
        } else { $Title = "session-$Id" }
    }

    $tenantId = if ($script:SessionState) { $script:SessionState.TenantDomain } else { 'unknown' }
    $costUsd  = if ($script:AICostState) { [double]$script:AICostState.SessionUsd } else { 0.0 }

    $blob = [ordered]@{
        schemaVersion  = 1
        id             = $Id
        title          = $Title
        createdUtc     = (Get-Date).ToUniversalTime().ToString('o')
        lastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        provider       = [string]$Config.Provider
        model          = [string]$Config.Model
        tenant         = $tenantId
        costRolledUpUsd= $costUsd
        privacyMap     = $script:PrivacyMap
        history        = $History
    }
    $plain = ($blob | ConvertTo-Json -Depth 16)
    $payload = Protect-AISessionPayload -PlainText $plain
    $envelope = [ordered]@{
        id          = $Id
        title       = $Title
        mode        = $payload.Mode
        data        = $payload.Data
    }
    $file = Join-Path (Get-AISessionDir) ("{0}.session" -f $Id)
    Set-Content -LiteralPath $file -Value ($envelope | ConvertTo-Json -Depth 4) -Encoding UTF8 -Force

    # Update index (skip privacyMap + history; index stays small + readable)
    $idx = Get-AISessionIndex
    $other = @($idx | Where-Object { $_.id -ne $Id })
    $meta = @{
        id             = $Id
        title          = $Title
        lastUpdatedUtc = $blob.lastUpdatedUtc
        provider       = $blob.provider
        model          = $blob.model
        tenant         = $blob.tenant
        costUsd        = $blob.costRolledUpUsd
        messages       = $History.Count
        encrypted      = ($payload.Mode -eq 'dpapi')
    }
    Save-AISessionIndex -Index (@($meta) + $other)

    $script:AISessionCurrent.Id             = $Id
    $script:AISessionCurrent.Title          = $Title
    $script:AISessionCurrent.DirtySinceSave = $false
    return $Id
}

function Load-AISession {
    <#
        Returns @{ Title; History; PrivacyMap; Provider; Model;
        Tenant; CostUsd } or $null on miss. Resolves -IdOrPrefix
        against id exact match first, then title prefix.
    #>
    param([Parameter(Mandatory)][string]$IdOrPrefix)
    $idx = Get-AISessionIndex
    $hit = $idx | Where-Object { $_.id -eq $IdOrPrefix } | Select-Object -First 1
    if (-not $hit) { $hit = $idx | Where-Object { $_.title -and $_.title.StartsWith($IdOrPrefix) } | Select-Object -First 1 }
    if (-not $hit) { return $null }

    $file = Join-Path (Get-AISessionDir) ("{0}.session" -f $hit.id)
    if (-not (Test-Path $file)) { return $null }
    try {
        $env = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json -ErrorAction Stop
        $plain = Unprotect-AISessionPayload -Mode $env.mode -Data $env.data
        $blob  = $plain | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Failed to load $($hit.id): $($_.Exception.Message)"
        return $null
    }

    # Rebuild PrivacyMap as a hashtable graph
    $pm = @{ ByValue=@{}; ByToken=@{}; Counters=@{ JWT=0; SECRET=0; THUMB=0; UPN=0; GUID=0; NAME=0; TENANT=0 } }
    if ($blob.privacyMap) {
        if ($blob.privacyMap.ByValue)  { foreach ($p in $blob.privacyMap.ByValue.PSObject.Properties)  { $pm.ByValue[[string]$p.Name]  = [string]$p.Value } }
        if ($blob.privacyMap.ByToken)  { foreach ($p in $blob.privacyMap.ByToken.PSObject.Properties)  { $pm.ByToken[[string]$p.Name]  = [string]$p.Value } }
        if ($blob.privacyMap.Counters) { foreach ($p in $blob.privacyMap.Counters.PSObject.Properties) { $pm.Counters[[string]$p.Name] = [int]$p.Value }    }
    }

    # Rebuild history as hashtables for proper downstream handling
    $history = @()
    foreach ($m in @($blob.history)) {
        $entry = @{ role = [string]$m.role }
        if ($m.content -is [string]) { $entry.content = [string]$m.content }
        else { $entry.content = $m.content }   # arrays of blocks pass through
        $history += $entry
    }

    $script:AISessionCurrent.Id             = $hit.id
    $script:AISessionCurrent.Title          = $hit.title
    $script:AISessionCurrent.Ephemeral      = $false
    $script:AISessionCurrent.DirtySinceSave = $false
    $script:PrivacyMap = $pm

    return @{
        Title      = $hit.title
        History    = $history
        PrivacyMap = $pm
        Provider   = [string]$blob.provider
        Model      = [string]$blob.model
        Tenant     = [string]$blob.tenant
        CostUsd    = [double]$blob.costRolledUpUsd
        Id         = $hit.id
    }
}

function Remove-AISession {
    param([Parameter(Mandatory)][string]$IdOrPrefix)
    $idx = Get-AISessionIndex
    $hit = $idx | Where-Object { $_.id -eq $IdOrPrefix } | Select-Object -First 1
    if (-not $hit) { $hit = $idx | Where-Object { $_.title -and $_.title.StartsWith($IdOrPrefix) } | Select-Object -First 1 }
    if (-not $hit) { Write-Warn "No session matches '$IdOrPrefix'."; return $false }
    $file = Join-Path (Get-AISessionDir) ("{0}.session" -f $hit.id)
    if (Test-Path $file) { Remove-Item -LiteralPath $file -Force }
    Save-AISessionIndex -Index @($idx | Where-Object { $_.id -ne $hit.id })
    if ($script:AISessionCurrent.Id -eq $hit.id) {
        $script:AISessionCurrent.Id = $null
        $script:AISessionCurrent.Title = $null
    }
    return $true
}

function Rename-AISession {
    param(
        [Parameter(Mandatory)][string]$IdOrPrefix,
        [Parameter(Mandatory)][string]$NewTitle
    )
    $idx = Get-AISessionIndex
    $hit = $idx | Where-Object { $_.id -eq $IdOrPrefix } | Select-Object -First 1
    if (-not $hit) { $hit = $idx | Where-Object { $_.title -and $_.title.StartsWith($IdOrPrefix) } | Select-Object -First 1 }
    if (-not $hit) { Write-Warn "No session matches '$IdOrPrefix'."; return $false }
    # Load + re-save under the new title (re-encrypts with new index entry).
    $blob = Load-AISession -IdOrPrefix $hit.id
    if (-not $blob) { return $false }
    $cfg = @{ Provider = $blob.Provider; Model = $blob.Model }
    Save-AISession -Config $cfg -History $blob.History -Title $NewTitle -Id $hit.id -Force | Out-Null
    return $true
}

function Show-AISessionList {
    <#
        /list -- compact table of saved sessions, newest first.
    #>
    $idx = Get-AISessionIndex
    if ($idx.Count -eq 0) { Write-Host "  (no saved sessions)" -ForegroundColor DarkGray; Write-Host ""; return }
    Write-Host ""
    Write-Host "  SAVED SESSIONS" -ForegroundColor $script:Colors.Title
    $sorted = $idx | Sort-Object { [datetime]::Parse($_.lastUpdatedUtc) } -Descending
    foreach ($s in $sorted) {
        $enc = if ($s.encrypted) { '[enc]' } else { '[plain]' }
        $tag = if ($script:AISessionCurrent.Id -eq $s.id) { ' <-- current' } else { '' }
        Write-Host ("  {0}  {1}  {2,-32}  {3,-30}  {4:N4} USD  {5} msg{6}" -f `
            $enc, $s.lastUpdatedUtc, ([string]$s.title), ([string]$s.id), [double]$s.costUsd, [int]$s.messages, $tag) -ForegroundColor White
    }
    Write-Host ""
}

function Set-AISessionEphemeral {
    param([bool]$On = $true)
    $script:AISessionCurrent.Ephemeral = $On
}

function Export-AISession {
    <#
        Write a *redacted* JSON file safe for sharing. Uses
        Convert-ToSafePayload on every string so real UPNs / tenant
        IDs / GUIDs / secrets become placeholders. The privacy map
        is INTENTIONALLY excluded from the export.
    #>
    param(
        [Parameter(Mandatory)][string]$IdOrPrefix,
        [string]$DestinationPath
    )
    $blob = Load-AISession -IdOrPrefix $IdOrPrefix
    if (-not $blob) { Write-Warn "No session matches '$IdOrPrefix'."; return $false }

    if (-not $DestinationPath) { $DestinationPath = Join-Path (Get-Location).Path ("session-export-" + $blob.Id + ".json") }
    $counts = @{ JWT=0; SECRET=0; THUMB=0; UPN=0; GUID=0; TENANT=0; NAME=0 }
    $redactedHistory = @()
    foreach ($m in $blob.History) {
        if ($m.content -is [string]) {
            $redactedHistory += @{ role = $m.role; content = (Convert-ToSafePayload -Text $m.content -SecretsOnly:$false -Counts $counts) }
        } else {
            # leave non-string content (tool_use blocks etc.) but stringify-then-redact embedded fields
            $serialized = ($m.content | ConvertTo-Json -Depth 12 -Compress)
            $redactedHistory += @{ role = $m.role; content = (Convert-ToSafePayload -Text $serialized -SecretsOnly:$false -Counts $counts) }
        }
    }
    $out = [ordered]@{
        schemaVersion = 1
        exportType    = 'redacted'
        title         = $blob.Title
        id            = $blob.Id
        provider      = $blob.Provider
        model         = $blob.Model
        costUsd       = $blob.CostUsd
        replacedCounts= $counts
        history       = $redactedHistory
    }
    Set-Content -LiteralPath $DestinationPath -Value ($out | ConvertTo-Json -Depth 16) -Encoding UTF8 -Force
    Write-InfoMsg ("Exported (redacted) to {0}" -f $DestinationPath)
    return $true
}

function Test-AISessionDirty {
    return [bool]$script:AISessionCurrent.DirtySinceSave
}

function Set-AISessionDirty { $script:AISessionCurrent.DirtySinceSave = $true }

function Get-AISessionCurrent { return $script:AISessionCurrent }
