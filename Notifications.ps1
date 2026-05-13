# ============================================================
#  Notifications.ps1 -- email + Teams webhook dispatch
#
#  Three exit channels:
#    Send-Email          Graph /me/sendMail or /users/{id}/sendMail
#    Send-TeamsWebhook   POST to a webhook URL (incoming webhook
#                        connector, Power Automate flow URL, etc.)
#    Send-Notification   dispatcher; pulls config from
#                        ai_config.json's Notifications section.
#
#  Webhook URLs are sensitive (anyone with the URL can post to
#  the channel) so they're DPAPI-protected at rest via the
#  generic Protect-Secret / Unprotect-Secret helpers added in
#  this commit (thin wrappers over the existing Protect-ApiKey
#  / Unprotect-ApiKey from Phase 0.5 Commit 2).
#
#  Redaction model: notification recipient addresses are NEVER
#  passed through Phase 0.5 Commit 6's Convert-ToSafePayload --
#  they're legitimate destinations, not data that needs to be
#  obscured. The AI privacy layer only ever sees outbound chat
#  payloads, never notification destinations, so this is
#  automatic; the README documents the contract for any future
#  AI integration.
# ============================================================

# ============================================================
#  Generic secret protect/unprotect (wrappers; delegate to the
#  existing AIAssistant.ps1 helpers if loaded, else re-implement
#  locally so this module is standalone).
# ============================================================

function Protect-Secret {
    param([Parameter(Mandatory)][string]$PlainText)
    if (Get-Command Protect-ApiKey -ErrorAction SilentlyContinue) {
        return (Protect-ApiKey -PlainKey $PlainText)
    }
    if ([string]::IsNullOrWhiteSpace($PlainText)) { return $PlainText }
    if ($PlainText -like 'DPAPI:*' -or $PlainText -like 'B64:*') { return $PlainText }
    try {
        $secure = ConvertTo-SecureString $PlainText -AsPlainText -Force
        $enc    = ConvertFrom-SecureString $secure
        return "DPAPI:$enc"
    } catch {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
        return "B64:$([Convert]::ToBase64String($bytes))"
    }
}

function Unprotect-Secret {
    param([Parameter(Mandatory)][string]$Stored)
    if (Get-Command Unprotect-ApiKey -ErrorAction SilentlyContinue) {
        return (Unprotect-ApiKey -StoredKey $Stored)
    }
    if ([string]::IsNullOrWhiteSpace($Stored)) { return $Stored }
    if ($Stored -like 'DPAPI:*') {
        $secure = ConvertTo-SecureString $Stored.Substring(6)
        return [System.Net.NetworkCredential]::new("", $secure).Password
    }
    if ($Stored -like 'B64:*') {
        return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Stored.Substring(4)))
    }
    return $Stored
}

# ============================================================
#  Notifications config
#  Lives inside ai_config.json's "Notifications" block. We
#  reuse Get-AIConfig + Save-AIConfig from AIAssistant.ps1 so
#  the encryption + comment-stripping infra is shared.
# ============================================================

$script:NotificationsDefault = @{
    DefaultEmailFrom         = ''
    SecurityTeamRecipients   = @()
    OperationsTeamRecipients = @()
    TeamsWebhookSecurity     = ''
    TeamsWebhookOperations   = ''
    DryRunNotifications      = $false
}

function Get-NotificationsConfig {
    if (-not (Get-Command Get-AIConfig -ErrorAction SilentlyContinue)) { return $script:NotificationsDefault }
    $cfg = Get-AIConfig
    if (-not $cfg) { return $script:NotificationsDefault }
    if (-not $cfg.ContainsKey('Notifications') -or -not ($cfg['Notifications'] -is [hashtable])) {
        # Migration: inject defaults + write back
        $cfg['Notifications'] = @{}
        foreach ($k in $script:NotificationsDefault.Keys) { $cfg['Notifications'][$k] = $script:NotificationsDefault[$k] }
        if (Get-Command Save-AIConfig -ErrorAction SilentlyContinue) { Save-AIConfig -Config $cfg | Out-Null }
        return $cfg['Notifications']
    }
    $n = $cfg['Notifications']
    foreach ($k in $script:NotificationsDefault.Keys) { if (-not $n.ContainsKey($k)) { $n[$k] = $script:NotificationsDefault[$k] } }
    return $n
}

function Set-NotificationsConfig {
    param([Parameter(Mandatory)][hashtable]$Updates)
    if (-not (Get-Command Get-AIConfig -ErrorAction SilentlyContinue)) { Write-ErrorMsg "Get-AIConfig not loaded."; return }
    $cfg = Get-AIConfig
    if (-not $cfg) { $cfg = @{} }
    if (-not $cfg.ContainsKey('Notifications')) { $cfg['Notifications'] = @{} }
    foreach ($k in $Updates.Keys) { $cfg['Notifications'][$k] = $Updates[$k] }
    # Encrypt webhooks if a plain URL was passed
    foreach ($k in 'TeamsWebhookSecurity','TeamsWebhookOperations') {
        $v = [string]$cfg['Notifications'][$k]
        if ($v -and $v -notlike 'DPAPI:*' -and $v -notlike 'B64:*' -and $v -like 'http*') {
            $cfg['Notifications'][$k] = Protect-Secret -PlainText $v
        }
    }
    if (Get-Command Save-AIConfig -ErrorAction SilentlyContinue) { Save-AIConfig -Config $cfg }
}

# ============================================================
#  Channels
# ============================================================

function Send-Email {
    <#
        Graph /me/sendMail (or /users/{id}/sendMail when -From
        differs from the connected account). Body is HTML.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$To,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [string]$From,
        [ValidateSet('Normal','High')][string]$Importance = 'Normal'
    )
    $cfg = Get-NotificationsConfig
    if ($cfg.DryRunNotifications) {
        Write-Host ("  [NOTIFY DRY-RUN] email to {0} -- {1}" -f ($To -join ', '), $Subject) -ForegroundColor Yellow
        Write-AuditEntry -EventType 'NOTIFY_DRYRUN' -Detail "email to $($To -join ',') -- $Subject" -ActionType 'NotifyEmail' -Target @{ to = $To; subject = $Subject } -Result 'preview' | Out-Null
        return $true
    }

    $payload = @{
        message = @{
            subject = $Subject
            body = @{ contentType = 'HTML'; content = $Body }
            toRecipients = @($To | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
            importance = $Importance
        }
        saveToSentItems = $true
    } | ConvertTo-Json -Depth 10

    $uri = if ($From -and $From -ne (Get-MgContext).Account) {
        "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    } else {
        "https://graph.microsoft.com/v1.0/me/sendMail"
    }

    return Invoke-Action `
        -Description ("Send email to {0} -- {1}" -f ($To -join ','), $Subject) `
        -ActionType 'NotifyEmail' `
        -Target @{ to = $To; from = $From; subject = $Subject; importance = $Importance } `
        -NoUndoReason 'Email send is irreversible.' `
        -Action {
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body $payload -ContentType 'application/json' -ErrorAction Stop | Out-Null
            $true
        }
}

function Send-TeamsWebhook {
    <#
        POST a simple HTML card to an incoming-webhook URL.
        Accepts either a plain webhook URL or a DPAPI/B64-encrypted
        one (Protect-Secret); decrypts automatically.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WebhookUrl,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Body,
        [string]$ThemeColor = '0078D4'
    )
    $cfg = Get-NotificationsConfig
    if ($cfg.DryRunNotifications) {
        Write-Host ("  [NOTIFY DRY-RUN] teams webhook -- {0}" -f $Title) -ForegroundColor Yellow
        Write-AuditEntry -EventType 'NOTIFY_DRYRUN' -Detail "teams webhook -- $Title" -ActionType 'NotifyTeams' -Target @{ title = $Title } -Result 'preview' | Out-Null
        return $true
    }

    $url = $WebhookUrl
    if ($url -like 'DPAPI:*' -or $url -like 'B64:*') { $url = Unprotect-Secret -Stored $url }
    if ([string]::IsNullOrWhiteSpace($url)) { Write-Warn "Empty webhook URL after decrypt."; return $false }

    $card = @{
        '@type'        = 'MessageCard'
        '@context'     = 'https://schema.org/extensions'
        themeColor     = $ThemeColor
        summary        = $Title
        title          = $Title
        text           = $Body
    } | ConvertTo-Json -Depth 6

    return Invoke-Action `
        -Description ("POST Teams webhook -- {0}" -f $Title) `
        -ActionType 'NotifyTeams' `
        -Target @{ title = $Title; themeColor = $ThemeColor } `
        -NoUndoReason 'Webhook delivery is irreversible.' `
        -Action {
            Invoke-RestMethod -Uri $url -Method POST -ContentType 'application/json' -Body $card -ErrorAction Stop | Out-Null
            $true
        }
}

function Send-Notification {
    <#
        Dispatcher. Routes a single notification to one or more
        channels, picking recipients / webhooks from the config
        based on -Severity (Info / Warning / Critical map to
        Operations / Security accordingly).

          Info     -> Operations
          Warning  -> Operations + Security
          Critical -> Security (high importance)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$Channels,       # email, teams, both
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$Body,
        [ValidateSet('Info','Warning','Critical')][string]$Severity = 'Info',
        [string[]]$To,
        [string]$WebhookUrl
    )
    $cfg = Get-NotificationsConfig
    $themeColor = switch ($Severity) { 'Critical' { 'B00020' } 'Warning' { 'F0A030' } default { '0078D4' } }
    $importance = if ($Severity -eq 'Critical') { 'High' } else { 'Normal' }

    # Recipient resolution
    $emailTo = $To
    if (-not $emailTo -or $emailTo.Count -eq 0) {
        $emailTo = switch ($Severity) {
            'Critical' { @($cfg.SecurityTeamRecipients) }
            'Warning'  { @($cfg.SecurityTeamRecipients) + @($cfg.OperationsTeamRecipients) | Where-Object { $_ } | Sort-Object -Unique }
            default    { @($cfg.OperationsTeamRecipients) }
        }
    }
    $webhook = $WebhookUrl
    if (-not $webhook) {
        $webhook = if ($Severity -eq 'Critical' -or $Severity -eq 'Warning') { $cfg.TeamsWebhookSecurity } else { $cfg.TeamsWebhookOperations }
    }

    $any = $false
    foreach ($ch in $Channels) {
        switch ($ch) {
            'email' {
                if (-not $emailTo -or $emailTo.Count -eq 0) {
                    Write-Warn "Email channel requested but no recipients configured/provided."
                } else {
                    $from = if ($cfg.DefaultEmailFrom) { $cfg.DefaultEmailFrom } else { $null }
                    Send-Email -To $emailTo -Subject $Subject -Body $Body -From $from -Importance $importance | Out-Null
                    $any = $true
                }
            }
            'teams' {
                if (-not $webhook) {
                    Write-Warn "Teams channel requested but no webhook configured/provided."
                } else {
                    Send-TeamsWebhook -WebhookUrl $webhook -Title $Subject -Body $Body -ThemeColor $themeColor | Out-Null
                    $any = $true
                }
            }
            'both' {
                if ($emailTo -and $emailTo.Count -gt 0) { $from = if ($cfg.DefaultEmailFrom) { $cfg.DefaultEmailFrom } else { $null }; Send-Email -To $emailTo -Subject $Subject -Body $Body -From $from -Importance $importance | Out-Null; $any = $true }
                if ($webhook) { Send-TeamsWebhook -WebhookUrl $webhook -Title $Subject -Body $Body -ThemeColor $themeColor | Out-Null; $any = $true }
            }
        }
    }
    return $any
}

# ============================================================
#  Smoke test + setup
# ============================================================

function Test-NotificationChannels {
    <#
        Send a known-safe ping to each configured channel.
        Tagged [TEST] in the subject so receivers can filter
        them out.
    #>
    $cfg = Get-NotificationsConfig
    $stamp = (Get-Date).ToUniversalTime().ToString('o')
    $subject = "[TEST] M365 Manager notification smoke test"
    $body = "<p>Test ping from <code>Test-NotificationChannels</code> at $stamp. If you're seeing this, the channel is configured correctly.</p>"

    $sent = @{}
    foreach ($key in 'SecurityTeamRecipients','OperationsTeamRecipients') {
        $rcpts = @($cfg.$key)
        if ($rcpts.Count -gt 0) {
            Send-Email -To $rcpts -Subject $subject -Body $body -From $cfg.DefaultEmailFrom | Out-Null
            $sent["email:$key"] = $rcpts.Count
        }
    }
    foreach ($key in 'TeamsWebhookSecurity','TeamsWebhookOperations') {
        $w = [string]$cfg.$key
        if ($w) {
            Send-TeamsWebhook -WebhookUrl $w -Title $subject -Body $body -ThemeColor '0078D4' | Out-Null
            $sent["teams:$key"] = 1
        }
    }
    if ($sent.Count -eq 0) {
        Write-Warn "No channels configured. Use Start-NotificationsSetup to add recipients / webhooks."
    } else {
        foreach ($k in $sent.Keys) { Write-Success "$k -> $($sent[$k]) target(s)" }
    }
}

function Start-NotificationsSetup {
    $cfg = Get-NotificationsConfig
    Write-SectionHeader "Notifications setup"
    Write-StatusLine "DefaultEmailFrom"         "$($cfg.DefaultEmailFrom)" 'White'
    Write-StatusLine "SecurityTeamRecipients"   "$($cfg.SecurityTeamRecipients -join ', ')" 'White'
    Write-StatusLine "OperationsTeamRecipients" "$($cfg.OperationsTeamRecipients -join ', ')" 'White'
    Write-StatusLine "TeamsWebhookSecurity"     $(if ($cfg.TeamsWebhookSecurity)   { '<set>' } else { '<empty>' }) 'White'
    Write-StatusLine "TeamsWebhookOperations"   $(if ($cfg.TeamsWebhookOperations) { '<set>' } else { '<empty>' }) 'White'
    Write-StatusLine "DryRunNotifications"      "$($cfg.DryRunNotifications)" 'White'

    while ($true) {
        $sel = Show-Menu -Title "Notifications" -Options @(
            "Set DefaultEmailFrom",
            "Set SecurityTeamRecipients (comma-separated)",
            "Set OperationsTeamRecipients (comma-separated)",
            "Set Teams webhook (Security)",
            "Set Teams webhook (Operations)",
            "Toggle DryRunNotifications",
            "Send test pings to every configured channel"
        ) -BackLabel "Done"
        switch ($sel) {
            0 { $v = Read-UserInput "DefaultEmailFrom"; if ($null -ne $v) { Set-NotificationsConfig -Updates @{ DefaultEmailFrom = $v.Trim() } } }
            1 { $v = Read-UserInput "Recipients"; if ($null -ne $v) { Set-NotificationsConfig -Updates @{ SecurityTeamRecipients = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } } }
            2 { $v = Read-UserInput "Recipients"; if ($null -ne $v) { Set-NotificationsConfig -Updates @{ OperationsTeamRecipients = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } } }
            3 { $v = Read-UserInput "Teams webhook URL (Security)"; if ($null -ne $v) { Set-NotificationsConfig -Updates @{ TeamsWebhookSecurity = $v.Trim() } } }
            4 { $v = Read-UserInput "Teams webhook URL (Operations)"; if ($null -ne $v) { Set-NotificationsConfig -Updates @{ TeamsWebhookOperations = $v.Trim() } } }
            5 { Set-NotificationsConfig -Updates @{ DryRunNotifications = (-not [bool]$cfg.DryRunNotifications) } }
            6 { Test-NotificationChannels }
            -1 { return }
        }
        $cfg = Get-NotificationsConfig
    }
}
