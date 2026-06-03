# ============================================================
#  IncidentResponse.ps1 -- compromised-account playbook
#
#  Invoke-CompromisedAccountResponse drives a 13-step sequence:
#  snapshot (forensic baseline) -> contain (block, revoke
#  sessions, wipe auth methods, force password) -> clean up
#  (disable inbox rules, clear forwarding) -> audit (24h, 7d
#  sent mail, 7d outbound shares) -> optional purge -> notify
#  -> report.
#
#  Every state-mutating step goes through Invoke-Action so it
#  audits, respects PREVIEW, and gets a reverse recipe where
#  reversal is possible. The snapshot is written BEFORE any
#  mutation so the operator always has the pre-incident state
#  on disk even if a later step fails.
#
#  Severity gating:
#    Low      : steps 1, 8, 9, 10, 13           (forensic only)
#    Medium   : Low  + 2, 3, 4, 5                (contain)
#    High     : Med  + 6, 7, 12                  (default)
#    Critical : High + 11 default-on prompt      (full)
#
#  Tenant-scoped: every artifact lives under
#  <stateDir>\<tenant-slug>\incidents\<incident-id>\.
# ============================================================

# ============================================================
#  Paths + helpers
# ============================================================

function Get-IncidentTenantSlug {
    <#
        Stable filesystem-safe slug for the current tenant. Mirrors
        the slug Audit.ps1 uses for its session-log filename.
        Returns 'default' when no tenant is set (which is the case
        for the first-run direct-admin path).
    #>
    if (-not $script:SessionState -or -not $script:SessionState.TenantName) {
        return 'default'
    }
    $slug = ($script:SessionState.TenantName -replace '[^A-Za-z0-9]+','_').ToLower()
    if (-not $slug) { return 'default' }
    return $slug
}

function Get-IncidentsDirectory {
    <#
        <stateDir>\<tenant-slug>\incidents\
        Created with mode 700 on POSIX. Per-tenant scoping closes
        one of the deferred Phase 6 retrofits (sessions / break-
        glass / incidents).
    #>
    $base = Get-StateDirectory
    if (-not $base) { return $null }
    $tenant = Get-IncidentTenantSlug
    $dir = Join-Path $base (Join-Path $tenant 'incidents')
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { return $null }
        if (-not $env:LOCALAPPDATA -and (Get-Command chmod -ErrorAction SilentlyContinue)) {
            try { & chmod 700 $dir 2>$null | Out-Null } catch {}
        }
    }
    return $dir
}

function Get-IncidentRegistryPath {
    <#
        <stateDir>\<tenant-slug>\incidents.jsonl  -- append-only,
        one record per incident. Companion to per-incident dirs.
    #>
    $base = Get-StateDirectory
    if (-not $base) { return $null }
    $tenant = Get-IncidentTenantSlug
    $tenantDir = Join-Path $base $tenant
    if (-not (Test-Path -LiteralPath $tenantDir)) {
        try { New-Item -ItemType Directory -Path $tenantDir -Force | Out-Null } catch { return $null }
    }
    return Join-Path $tenantDir 'incidents.jsonl'
}

function New-IncidentId {
    <#
        INC-YYYY-MM-DD-xxxx where xxxx is the first 4 chars of a
        fresh GUID. Date prefix makes the ids sortable and tells
        operators at a glance whether an incident is recent.
    #>
    $date = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd')
    $rand = ([guid]::NewGuid().ToString().Substring(0,4))
    return ("INC-{0}-{1}" -f $date, $rand)
}

function Get-IncidentDirectory {
    param([Parameter(Mandatory)][string]$IncidentId)
    $root = Get-IncidentsDirectory
    if (-not $root) { return $null }
    $dir = Join-Path $root $IncidentId
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        if (-not $env:LOCALAPPDATA -and (Get-Command chmod -ErrorAction SilentlyContinue)) {
            try { & chmod 700 $dir 2>$null | Out-Null } catch {}
        }
    }
    return $dir
}

function Write-IncidentArtifact {
    <#
        Persist a JSON or HTML artifact into the incident dir. JSON
        objects go through ConvertTo-Json -Depth 12; everything else
        is written verbatim. Returns the path written.
    #>
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$Filename,
        [Parameter(Mandatory)]$Content
    )
    $dir = Get-IncidentDirectory -IncidentId $IncidentId
    if (-not $dir) { return $null }
    $path = Join-Path $dir $Filename
    if ($Content -is [string]) {
        Set-Content -LiteralPath $path -Value $Content -Encoding UTF8 -Force
    } else {
        $json = $Content | ConvertTo-Json -Depth 12
        Set-Content -LiteralPath $path -Value $json -Encoding UTF8 -Force
    }
    return $path
}

function Write-IncidentRegistryRecord {
    <#
        Append one incident record to <tenant>/incidents.jsonl.
        Uses Add-Content so two parallel writers can both append
        safely (jsonl is line-atomic).
    #>
    param([Parameter(Mandatory)][hashtable]$Record)
    $path = Get-IncidentRegistryPath
    if (-not $path) { return }
    $json = $Record | ConvertTo-Json -Depth 8 -Compress
    try { Add-Content -LiteralPath $path -Value $json -ErrorAction Stop } catch {
        Write-Warn "Could not append incident registry record: $($_.Exception.Message)"
    }
}

function Write-IncidentAuditEntry {
    <#
        Wrapper around Write-AuditEntry that injects the incidentId
        into the Target hashtable + tags ActionType with an Incident
        prefix so AuditViewer filtering can find every entry tied
        to one incident in one query.
    #>
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$EventType,
        [Parameter(Mandatory)][string]$Detail,
        [string]$ActionType,
        [hashtable]$Target,
        [string]$Result = 'info',
        [string]$ErrorMessage
    )
    if (-not $Target) { $Target = @{} }
    $Target.incidentId = $IncidentId
    $args = @{
        EventType = $EventType
        Detail    = $Detail
        Target    = $Target
        Result    = $Result
    }
    if ($ActionType)   { $args.ActionType   = $ActionType }
    if ($ErrorMessage) { $args.ErrorMessage = $ErrorMessage }
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry @args | Out-Null
    }
}

function New-IncidentPassword {
    <#
        24-char cryptographically-strong password meeting the M365
        complexity rules. Returns the plain string -- the caller is
        responsible for handing it to the operator securely.
    #>
    $bytes = New-Object byte[] 18
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $base64 = [Convert]::ToBase64String($bytes)
    # Force at least one of each character class M365 wants.
    return ($base64 -replace '[+/=]','x') + '!Aa1'
}

# ============================================================
#  Snapshot (step 1) -- forensic baseline, read-only
# ============================================================

function Get-IncidentSnapshot {
    <#
        Capture every fact about the user we'd want for forensics
        BEFORE we mutate anything. Returns the snapshot hashtable
        and writes snapshot.json to the incident dir.

        Read-only -- no Invoke-Action wrapping needed.
    #>
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN
    )
    $snap = [ordered]@{
        capturedAtUtc      = (Get-Date).ToUniversalTime().ToString('o')
        upn                = $UPN
        userId             = $null
        accountEnabled     = $null
        signInActivity     = $null
        manager            = $null
        groupMemberships   = @()
        licenses           = @()
        mfaMethods         = @()
        mailboxForwarding  = $null
        deliverToMailboxAndForward = $null
        inboxRules         = @()
        recentSignIns      = @()
        capturedErrors     = @()
    }

    # User core
    try {
        $u = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/$UPN" + '?$select=id,accountEnabled,signInActivity,displayName,userPrincipalName') -ErrorAction Stop
        $snap.userId         = [string]$u.id
        $snap.accountEnabled = [bool]$u.accountEnabled
        $snap.signInActivity = $u.signInActivity
        $snap.displayName    = [string]$u.displayName
    } catch { $snap.capturedErrors += "user: $($_.Exception.Message)" }

    # Manager
    try {
        $m = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/manager?`$select=id,userPrincipalName,displayName" -ErrorAction Stop
        if ($m) { $snap.manager = @{ id = [string]$m.id; upn = [string]$m.userPrincipalName; displayName = [string]$m.displayName } }
    } catch {
        if ($_.Exception.Message -notmatch 'Not Found|404') {
            $snap.capturedErrors += "manager: $($_.Exception.Message)"
        }
    }

    # Group memberships
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/memberOf?`$select=id,displayName,@odata.type" -ErrorAction Stop
        foreach ($g in @($resp.value)) {
            $snap.groupMemberships += @{ id = [string]$g.id; displayName = [string]$g.displayName; type = [string]$g.'@odata.type' }
        }
    } catch { $snap.capturedErrors += "groups: $($_.Exception.Message)" }

    # Licenses
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/licenseDetails" -ErrorAction Stop
        foreach ($l in @($resp.value)) {
            $snap.licenses += @{ skuId = [string]$l.skuId; skuPartNumber = [string]$l.skuPartNumber }
        }
    } catch { $snap.capturedErrors += "licenses: $($_.Exception.Message)" }

    # MFA methods (via existing helper if loaded)
    if (Get-Command Get-UserAuthMethods -ErrorAction SilentlyContinue) {
        try {
            $methods = @(Get-UserAuthMethods -User $UPN)
            foreach ($m in $methods) { $snap.mfaMethods += @{ id = [string]$m.Id; type = [string]$m.Label; urlSegment = [string]$m.UrlSegment } }
        } catch { $snap.capturedErrors += "mfaMethods: $($_.Exception.Message)" }
    }

    # Mailbox forwarding (EXO)
    if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
        try {
            $mbx = Get-Mailbox -Identity $UPN -ErrorAction Stop
            $snap.mailboxForwarding         = if ($mbx.ForwardingSmtpAddress) { [string]$mbx.ForwardingSmtpAddress } elseif ($mbx.ForwardingAddress) { [string]$mbx.ForwardingAddress } else { $null }
            $snap.deliverToMailboxAndForward = [bool]$mbx.DeliverToMailboxAndForward
        } catch { $snap.capturedErrors += "mailbox: $($_.Exception.Message)" }
    }

    # Inbox rules
    try {
        $rules = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/mailFolders/inbox/messageRules" -ErrorAction Stop
        foreach ($r in @($rules.value)) {
            $snap.inboxRules += @{
                id          = [string]$r.id
                displayName = [string]$r.displayName
                isEnabled   = [bool]$r.isEnabled
                actions     = $r.actions
                conditions  = $r.conditions
            }
        }
    } catch { $snap.capturedErrors += "inboxRules: $($_.Exception.Message)" }

    # Recent sign-ins (last 24h via Search-SignIns if loaded)
    if (Get-Command Search-SignIns -ErrorAction SilentlyContinue) {
        try {
            $signIns = @(Search-SignIns -User $UPN -From (Get-Date).AddDays(-1) -MaxResults 100)
            foreach ($s in $signIns) {
                $snap.recentSignIns += @{
                    ts        = [string]$s.CreatedDateTime
                    ipAddress = [string]$s.IpAddress
                    location  = [string]$s.Location
                    appName   = [string]$s.AppDisplayName
                    status    = [string]$s.Status
                    risk      = [string]$s.RiskLevel
                }
            }
        } catch { $snap.capturedErrors += "signIns: $($_.Exception.Message)" }
    }

    Write-IncidentArtifact -IncidentId $IncidentId -Filename 'snapshot.json' -Content $snap | Out-Null
    return $snap
}

# ============================================================
#  Step implementations (each one wraps Invoke-Action)
# ============================================================

function Invoke-IncidentBlockSignIn {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)]$Snapshot
    )
    return (Invoke-Action `
        -Description ("Incident {0}: Block sign-in for {1}" -f $IncidentId, $UPN) `
        -ActionType ("Incident:BlockSignIn") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN; userId = [string]$Snapshot.userId } `
        -ReverseType 'UnblockSignIn' `
        -ReverseDescription ("Re-enable sign-in for {0} (incident {1})" -f $UPN, $IncidentId) `
        -ReverseTarget @{ userUpn = $UPN; userId = [string]$Snapshot.userId } `
        -Action {
            $body = @{ accountEnabled = $false } | ConvertTo-Json
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$UPN" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
            $true
        })
}

function Invoke-IncidentRevokeSessions {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)]$Snapshot
    )
    return (Invoke-Action `
        -Description ("Incident {0}: Revoke all sign-in sessions for {1}" -f $IncidentId, $UPN) `
        -ActionType ("Incident:RevokeSessions") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN; userId = [string]$Snapshot.userId } `
        -NoUndoReason 'Session revocation is irreversible by design -- the user must re-authenticate on next access.' `
        -Action {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$UPN/revokeSignInSessions" -ErrorAction Stop | Out-Null
            $true
        })
}

function Invoke-IncidentRevokeAuthMethods {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN
    )
    return (Invoke-Action `
        -Description ("Incident {0}: Revoke ALL auth methods for {1}" -f $IncidentId, $UPN) `
        -ActionType ("Incident:RevokeAuthMethods") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN } `
        -NoUndoReason 'Auth method revocation cannot be undone via API; recovery requires operator-driven TAP + re-enrollment.' `
        -Action {
            if (-not (Get-Command Remove-AllAuthMethods -ErrorAction SilentlyContinue)) {
                throw "Remove-AllAuthMethods not loaded (MFAManager.ps1 missing?)."
            }
            $n = Remove-AllAuthMethods -User $UPN
            $n
        })
}

function Invoke-IncidentForcePasswordChange {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [switch]$NonInteractive
    )
    $newPwd = New-IncidentPassword
    $ok = Invoke-Action `
        -Description ("Incident {0}: Force password change + reset password for {1}" -f $IncidentId, $UPN) `
        -ActionType ("Incident:ForcePasswordChange") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN } `
        -NoUndoReason 'Password reset is irreversible -- operator must communicate the new credential to the legitimate user via a side channel.' `
        -Action {
            $body = @{
                passwordProfile = @{
                    forceChangePasswordNextSignIn = $true
                    password                       = $newPwd
                }
            } | ConvertTo-Json -Depth 4
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$UPN" -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
            $true
        }
    if ($ok) {
        # Deliver password to operator. Interactive: clipboard. Non-interactive: write
        # to incident dir with restrictive ACL (creator only on Windows / chmod 600 on POSIX).
        if (-not $NonInteractive -and (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) {
            try {
                Set-Clipboard -Value $newPwd
                Write-InfoMsg "Temporary password copied to clipboard. Deliver via a secure side-channel and clear the clipboard."
            } catch {
                Write-Warn "Could not write to clipboard ($($_.Exception.Message)); writing to incident dir instead."
                $NonInteractive = $true
            }
        }
        if ($NonInteractive -or -not (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) {
            $pwdFile = Write-IncidentArtifact -IncidentId $IncidentId -Filename 'temp-password.txt' -Content $newPwd
            if (-not $env:LOCALAPPDATA -and (Get-Command chmod -ErrorAction SilentlyContinue) -and $pwdFile) {
                try { & chmod 600 $pwdFile 2>$null | Out-Null } catch {}
            }
            Write-InfoMsg "Temporary password written to $pwdFile (delete after delivery)."
        }
    }
    return $ok
}

function Invoke-IncidentDisableInboxRules {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)]$Snapshot
    )
    # Snapshot already captured the rules. Persist a focused
    # inbox-rules.json so the operator can re-enable specific ones
    # later, and so the snapshot.json's potentially large dump isn't
    # the only forensic copy.
    $rules = @($Snapshot.inboxRules)
    Write-IncidentArtifact -IncidentId $IncidentId -Filename 'inbox-rules.json' -Content $rules | Out-Null

    if ($rules.Count -eq 0) {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'INFO' -Detail "No inbox rules to disable for $UPN." -ActionType 'Incident:DisableInboxRules' -Target @{ userUpn = $UPN } -Result 'info'
        return $true
    }

    $disabled = 0
    foreach ($r in $rules) {
        $ruleId   = [string]$r.id
        $ruleName = [string]$r.displayName
        $wasEnabled = [bool]$r.isEnabled
        if (-not $wasEnabled) { continue }   # Don't bother disabling already-disabled rules

        $ok = Invoke-Action `
            -Description ("Incident {0}: Disable inbox rule '{1}' on {2}" -f $IncidentId, $ruleName, $UPN) `
            -ActionType ("Incident:DisableInboxRule") `
            -Target @{ incidentId = $IncidentId; userUpn = $UPN; ruleId = $ruleId; ruleName = $ruleName } `
            -ReverseType 'EnableInboxRule' `
            -ReverseDescription ("Re-enable inbox rule '{0}' on {1}" -f $ruleName, $UPN) `
            -ReverseTarget @{ userUpn = $UPN; ruleId = $ruleId; ruleName = $ruleName } `
            -Action {
                $body = @{ isEnabled = $false } | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH -Uri ("https://graph.microsoft.com/v1.0/users/{0}/mailFolders/inbox/messageRules/{1}" -f $UPN, $ruleId) -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $true
            }
        if ($ok) { $disabled++ }
    }
    return $disabled
}

function Invoke-IncidentClearForwarding {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)]$Snapshot
    )
    if (-not (Get-Command Set-Mailbox -ErrorAction SilentlyContinue)) {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'SKIP' -Detail "Set-Mailbox not loaded -- skipping forwarding clear." -ActionType 'Incident:ClearForwarding' -Target @{ userUpn = $UPN } -Result 'info'
        return $true
    }
    $hadForwarding = ($Snapshot.mailboxForwarding -or $Snapshot.deliverToMailboxAndForward)
    if (-not $hadForwarding) {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'INFO' -Detail "No mailbox forwarding to clear for $UPN." -ActionType 'Incident:ClearForwarding' -Target @{ userUpn = $UPN } -Result 'info'
        return $true
    }

    $prior = @{
        userUpn                      = $UPN
        forwardingAddress            = $Snapshot.mailboxForwarding
        deliverToMailboxAndForward   = [bool]$Snapshot.deliverToMailboxAndForward
    }
    return (Invoke-Action `
        -Description ("Incident {0}: Clear mailbox forwarding for {1} (was -> {2})" -f $IncidentId, $UPN, $Snapshot.mailboxForwarding) `
        -ActionType ("Incident:ClearForwarding") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN; priorForwarding = $Snapshot.mailboxForwarding } `
        -ReverseType 'SetForwarding' `
        -ReverseDescription ("Restore forwarding {0} -> {1} (incident {2})" -f $UPN, $Snapshot.mailboxForwarding, $IncidentId) `
        -ReverseTarget $prior `
        -Action {
            Set-Mailbox -Identity $UPN -ForwardingAddress $null -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $false -ErrorAction Stop
            $true
        })
}

function Invoke-IncidentAudit24h {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN
    )
    $bundle = [ordered]@{
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        upn           = $UPN
        windowHours   = 24
        signIns       = @()
        unifiedAudit  = @()
        errors        = @()
    }
    if (Get-Command Search-SignIns -ErrorAction SilentlyContinue) {
        try { $bundle.signIns = @(Search-SignIns -User $UPN -From (Get-Date).AddDays(-1) -MaxResults 200) }
        catch { $bundle.errors += "signIns: $($_.Exception.Message)" }
    }
    if (Get-Command Search-UAL -ErrorAction SilentlyContinue) {
        try { $bundle.unifiedAudit = @(Search-UAL -UserId $UPN -From (Get-Date).AddDays(-1)) }
        catch { $bundle.errors += "ual: $($_.Exception.Message)" }
    }
    Write-IncidentArtifact -IncidentId $IncidentId -Filename 'audit-24h.json' -Content $bundle | Out-Null
    Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'INFO' -Detail ("24h audit captured for {0} ({1} sign-ins, {2} UAL rows)" -f $UPN, @($bundle.signIns).Count, @($bundle.unifiedAudit).Count) -ActionType 'Incident:Audit24h' -Target @{ userUpn = $UPN } -Result 'info'
    return $bundle
}

function Invoke-IncidentAuditSentMail {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN
    )
    $bundle = [ordered]@{
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        upn           = $UPN
        windowDays    = 7
        messages      = @()
        externalRecipients = @()
        errors        = @()
    }
    try {
        $cutoff = (Get-Date).AddDays(-7).ToUniversalTime().ToString('o')
        $uri = "https://graph.microsoft.com/v1.0/users/$UPN/mailFolders/sentitems/messages?`$top=200&`$select=id,subject,toRecipients,ccRecipients,bccRecipients,sentDateTime&`$filter=sentDateTime ge $cutoff"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $userDomain = ($UPN -split '@')[-1].ToLowerInvariant()
        $externalSet = @{}
        foreach ($m in @($resp.value)) {
            $recipients = @()
            foreach ($collection in @($m.toRecipients, $m.ccRecipients, $m.bccRecipients)) {
                foreach ($r in @($collection)) {
                    $addr = [string]$r.emailAddress.address
                    if (-not $addr) { continue }
                    $recipients += $addr
                    $dom = ($addr -split '@')[-1].ToLowerInvariant()
                    if ($dom -and $dom -ne $userDomain) { $externalSet[$addr] = $true }
                }
            }
            $bundle.messages += @{
                id            = [string]$m.id
                subject       = [string]$m.subject
                sentUtc       = [string]$m.sentDateTime
                recipients    = $recipients
            }
        }
        $bundle.externalRecipients = @($externalSet.Keys | Sort-Object)
    } catch { $bundle.errors += "sentmail: $($_.Exception.Message)" }

    Write-IncidentArtifact -IncidentId $IncidentId -Filename 'mail-sent-7d.json' -Content $bundle | Out-Null
    Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'INFO' -Detail ("7d sent-mail audit captured for {0} ({1} messages, {2} external recipients)" -f $UPN, @($bundle.messages).Count, @($bundle.externalRecipients).Count) -ActionType 'Incident:AuditSentMail' -Target @{ userUpn = $UPN } -Result 'info'
    return $bundle
}

function Invoke-IncidentAuditShares {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN
    )
    $bundle = [ordered]@{
        capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        upn           = $UPN
        windowDays    = 7
        shares        = @()
        externalDomains = @()
        errors        = @()
    }
    if (-not (Get-Command Get-UserOutboundShares -ErrorAction SilentlyContinue)) {
        $bundle.errors += 'Get-UserOutboundShares not loaded (SharePoint.ps1 missing?).'
    } else {
        try {
            $shares = @(Get-UserOutboundShares -UPN $UPN -LookbackDays 7)
            $userDomain = ($UPN -split '@')[-1].ToLowerInvariant()
            $domSet = @{}
            foreach ($s in $shares) {
                $bundle.shares += @{
                    ts          = [string]$s.SharedAtUtc
                    target      = [string]$s.TargetUserOrEmail
                    item        = [string]$s.ItemName
                    site        = [string]$s.SiteUrl
                    permission  = [string]$s.Permission
                }
                $addr = [string]$s.TargetUserOrEmail
                if ($addr -and $addr.Contains('@')) {
                    $dom = ($addr -split '@')[-1].ToLowerInvariant()
                    if ($dom -and $dom -ne $userDomain) { $domSet[$dom] = $true }
                }
            }
            $bundle.externalDomains = @($domSet.Keys | Sort-Object)
        } catch { $bundle.errors += "shares: $($_.Exception.Message)" }
    }
    Write-IncidentArtifact -IncidentId $IncidentId -Filename 'shares-7d.json' -Content $bundle | Out-Null
    Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'INFO' -Detail ("7d outbound-shares audit captured for {0} ({1} shares to {2} external domains)" -f $UPN, @($bundle.shares).Count, @($bundle.externalDomains).Count) -ActionType 'Incident:AuditShares' -Target @{ userUpn = $UPN } -Result 'info'
    return $bundle
}

function Invoke-IncidentQuarantineSentMail {
    <#
        Step 11 (Critical only by default). Compliance-purge of the
        last 7 days of sent mail. ALWAYS requires explicit operator
        confirmation -- in NonInteractive mode this short-circuits
        with a "manual step required" entry rather than auto-purging.
    #>
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [switch]$NonInteractive
    )
    if ($NonInteractive) {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'SKIP' -Detail "Quarantine requested but NonInteractive mode -- manual step required." -ActionType 'Incident:QuarantineManualRequired' -Target @{ userUpn = $UPN } -Result 'warn'
        Write-Warn "Quarantine step skipped in non-interactive mode. Run manually: New-ComplianceSearch + Purge."
        return $false
    }

    Write-Host ""
    Write-Host "  Quarantine the last 7 days of sent mail from $UPN?" -ForegroundColor Yellow
    Write-Host "  This is IRREVERSIBLE -- purged mail cannot be restored." -ForegroundColor Red
    $ans = Read-Host "  Type 'PURGE' to proceed, anything else to skip"
    if ($ans -ne 'PURGE') {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'SKIP' -Detail "Operator declined quarantine." -ActionType 'Incident:QuarantineDeclined' -Target @{ userUpn = $UPN } -Result 'info'
        Write-InfoMsg "Quarantine declined."
        return $false
    }

    return (Invoke-Action `
        -Description ("Incident {0}: Purge last-7d sent mail from {1}" -f $IncidentId, $UPN) `
        -ActionType ("Incident:QuarantineSentMail") `
        -Target @{ incidentId = $IncidentId; userUpn = $UPN; windowDays = 7 } `
        -NoUndoReason 'Compliance purge is irreversible -- the messages are removed from the tenant permanently.' `
        -Action {
            if (-not (Get-Command New-ComplianceSearch -ErrorAction SilentlyContinue)) {
                throw "New-ComplianceSearch not available (connect to Security & Compliance Center first)."
            }
            $searchName = "Incident-$IncidentId-$([guid]::NewGuid().ToString().Substring(0,8))"
            $kql = "Sender:$UPN AND Date>=$((Get-Date).AddDays(-7).ToString('yyyy-MM-dd'))"
            New-ComplianceSearch -Name $searchName -ExchangeLocation $UPN -ContentMatchQuery $kql -ErrorAction Stop | Out-Null
            Start-ComplianceSearch -Identity $searchName -ErrorAction Stop | Out-Null
            New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete -Confirm:$false -ErrorAction Stop | Out-Null
            $true
        })
}

function Invoke-IncidentNotify {
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$Severity,
        [string]$ReportPath
    )
    if (-not (Get-Command Send-Notification -ErrorAction SilentlyContinue)) {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'SKIP' -Detail 'Send-Notification not loaded -- skipping notify step.' -ActionType 'Incident:Notify' -Target @{ userUpn = $UPN } -Result 'info'
        return $false
    }
    $subj = "Compromised-account response: $UPN [$Severity] [$IncidentId]"
    $body = @"
<html><body style='font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222'>
<h2 style='color:#a00'>Compromised-account response triggered</h2>
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px'>
  <tr><th align='left'>Incident</th><td>$IncidentId</td></tr>
  <tr><th align='left'>UPN</th><td>$([System.Net.WebUtility]::HtmlEncode($UPN))</td></tr>
  <tr><th align='left'>Severity</th><td>$Severity</td></tr>
  <tr><th align='left'>Tenant</th><td>$([System.Net.WebUtility]::HtmlEncode($script:SessionState.TenantName))</td></tr>
  <tr><th align='left'>Triggered (UTC)</th><td>$((Get-Date).ToUniversalTime().ToString('o'))</td></tr>
</table>
<p>Full report and forensic snapshot: <code>$ReportPath</code></p>
<p style='color:#666;font-size:12px'>Sent automatically by M365 Manager incident-response playbook.</p>
</body></html>
"@
    try {
        Send-Notification -Channels SecurityTeam -Severity Critical -Subject $subj -Body $body | Out-Null
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'OK' -Detail "Security team notified of $IncidentId." -ActionType 'Incident:Notify' -Target @{ userUpn = $UPN; severity = $Severity } -Result 'success'
        return $true
    } catch {
        Write-IncidentAuditEntry -IncidentId $IncidentId -EventType 'ERROR' -Detail "Notify failed: $($_.Exception.Message)" -ActionType 'Incident:Notify' -Target @{ userUpn = $UPN } -Result 'failure' -ErrorMessage $_.Exception.Message
        Write-Warn "Notify step failed: $($_.Exception.Message)"
        return $false
    }
}

function New-IncidentReport {
    <#
        Step 13. Single-page HTML summarizing the whole incident:
        header (id, upn, severity, time), step outcomes, findings
        from each audit, recommended next steps, links to the
        per-step JSON artifacts.
    #>
    param(
        [Parameter(Mandatory)][string]$IncidentId,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][string]$Severity,
        [Parameter(Mandatory)]$Snapshot,
        [Parameter(Mandatory)][array]$StepResults,
        [hashtable]$Audit24h,
        [hashtable]$AuditSentMail,
        [hashtable]$AuditShares
    )
    $rows = New-Object System.Collections.ArrayList
    foreach ($s in $StepResults) {
        $statusCell = if ($s.Status -eq 'success' -or $s.Status -eq $true) { '<span style=color:#080>OK</span>' } elseif ($s.Status -eq 'skipped') { '<span style=color:#888>skipped</span>' } else { '<span style=color:#a00>failed</span>' }
        [void]$rows.Add(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f `
            [int]$s.StepNumber, [System.Net.WebUtility]::HtmlEncode([string]$s.StepName), $statusCell, [System.Net.WebUtility]::HtmlEncode([string]$s.Detail)))
    }
    $stepRowsHtml = ($rows -join "`n")

    $findingsLines = New-Object System.Collections.ArrayList
    if ($Audit24h) {
        [void]$findingsLines.Add(("Last 24h: {0} sign-ins, {1} unified-audit rows." -f @($Audit24h.signIns).Count, @($Audit24h.unifiedAudit).Count))
    }
    if ($AuditSentMail) {
        [void]$findingsLines.Add(("Sent mail (7d): {0} messages to {1} external recipient(s)." -f @($AuditSentMail.messages).Count, @($AuditSentMail.externalRecipients).Count))
        if (@($AuditSentMail.externalRecipients).Count -gt 0) {
            [void]$findingsLines.Add("External recipients to investigate: " + (@($AuditSentMail.externalRecipients) -join ', '))
        }
    }
    if ($AuditShares) {
        [void]$findingsLines.Add(("Outbound shares (7d): {0} shares to {1} external domain(s)." -f @($AuditShares.shares).Count, @($AuditShares.externalDomains).Count))
        if (@($AuditShares.externalDomains).Count -gt 0) {
            [void]$findingsLines.Add("External share domains: " + (@($AuditShares.externalDomains) -join ', '))
        }
    }
    $findingsHtml = ($findingsLines | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }) -join "`n"

    $nextSteps = @(
        "Communicate the incident to the user via a side channel (call, in-person) -- do NOT email the affected account."
        "Audit downstream systems the user had access to (CRM, code repos, finance systems, partner portals)."
        "If phishing was sent from this account, notify recipients to disregard and report received messages."
        "Schedule a credential rotation for any shared service accounts the user managed."
        "Review session activity in connected SaaS apps (Slack, GitHub, Salesforce, etc.) for the same time window."
        "Close the incident with Close-Incident -Id $IncidentId once recovery is complete."
    )
    $nextStepsHtml = ($nextSteps | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }) -join "`n"

    $tenant = if ($script:SessionState -and $script:SessionState.TenantName) { $script:SessionState.TenantName } else { 'unknown' }

    $html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<title>Incident $IncidentId -- $UPN</title>
<style>
  body { font-family: -apple-system, Segoe UI, Helvetica, Arial, sans-serif; color: #222; max-width: 980px; margin: 1em auto; padding: 0 1em; }
  h1 { color: #a00; border-bottom: 2px solid #a00; padding-bottom: 4px; }
  h2 { color: #444; margin-top: 1.5em; }
  table { border-collapse: collapse; font-size: 13px; width: 100%; }
  th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; }
  th { background: #f4f4f4; }
  .meta th { width: 12em; }
  ul li { margin: 4px 0; }
  code { background: #f4f4f4; padding: 1px 5px; border-radius: 3px; font-size: 12px; }
</style></head><body>
<h1>Compromised-account response: $([System.Net.WebUtility]::HtmlEncode($UPN))</h1>
<table class='meta'>
  <tr><th>Incident</th><td><code>$IncidentId</code></td></tr>
  <tr><th>Severity</th><td>$Severity</td></tr>
  <tr><th>Tenant</th><td>$([System.Net.WebUtility]::HtmlEncode($tenant))</td></tr>
  <tr><th>Started (UTC)</th><td>$((Get-Date).ToUniversalTime().ToString('o'))</td></tr>
  <tr><th>User account enabled (pre-incident)</th><td>$($Snapshot.accountEnabled)</td></tr>
  <tr><th>Groups (pre-incident)</th><td>$(@($Snapshot.groupMemberships).Count)</td></tr>
  <tr><th>Licenses (pre-incident)</th><td>$(@($Snapshot.licenses).Count)</td></tr>
  <tr><th>MFA methods (pre-incident)</th><td>$(@($Snapshot.mfaMethods).Count)</td></tr>
  <tr><th>Mailbox forwarding (pre-incident)</th><td>$([System.Net.WebUtility]::HtmlEncode([string]$Snapshot.mailboxForwarding))</td></tr>
  <tr><th>Inbox rules (pre-incident)</th><td>$(@($Snapshot.inboxRules).Count)</td></tr>
</table>

<h2>Step outcomes</h2>
<table>
  <tr><th>#</th><th>Step</th><th>Status</th><th>Detail</th></tr>
$stepRowsHtml
</table>

<h2>Findings</h2>
<ul>
$findingsHtml
</ul>

<h2>Recommended next steps</h2>
<ol>
$nextStepsHtml
</ol>

<h2>Artifacts</h2>
<ul>
  <li><code>snapshot.json</code> -- pre-incident state</li>
  <li><code>audit-24h.json</code> -- sign-in + UAL activity (last 24h)</li>
  <li><code>mail-sent-7d.json</code> -- sent mail audit</li>
  <li><code>shares-7d.json</code> -- outbound shares audit</li>
  <li><code>inbox-rules.json</code> -- captured rules at time of incident</li>
</ul>

<p style='color:#888;font-size:12px;margin-top:2em'>Generated by M365 Manager incident-response playbook. Snapshot dir: <code>$(Get-IncidentDirectory -IncidentId $IncidentId)</code></p>
</body></html>
"@
    $path = Write-IncidentArtifact -IncidentId $IncidentId -Filename 'report.html' -Content $html
    return $path
}

# ============================================================
#  Severity gate
# ============================================================

function Get-IncidentSteps {
    <#
        Return the ordered step list for the given severity.
        Step set is closed (no plugins) -- adding a new step means
        updating this function.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Low','Medium','High','Critical')][string]$Severity,
        [switch]$QuarantineSentMail
    )
    $steps = @()
    # Step 1 always runs (snapshot is the forensic baseline)
    $steps += @{ Number = 1;  Name = 'Snapshot'           }
    # Containment (Medium+)
    if ($Severity -in 'Medium','High','Critical') {
        $steps += @{ Number = 2; Name = 'BlockSignIn'           }
        $steps += @{ Number = 3; Name = 'RevokeSessions'        }
        $steps += @{ Number = 4; Name = 'RevokeAuthMethods'     }
        $steps += @{ Number = 5; Name = 'ForcePasswordChange'   }
    }
    # Cleanup (High+)
    if ($Severity -in 'High','Critical') {
        $steps += @{ Number = 6; Name = 'DisableInboxRules' }
        $steps += @{ Number = 7; Name = 'ClearForwarding'   }
    }
    # Audit (always runs -- forensic value at every severity)
    $steps += @{ Number = 8;  Name = 'Audit24h'      }
    $steps += @{ Number = 9;  Name = 'AuditSentMail' }
    $steps += @{ Number = 10; Name = 'AuditShares'   }
    # Quarantine (Critical default-on prompt + explicit flag)
    if ($QuarantineSentMail) {
        $steps += @{ Number = 11; Name = 'QuarantineSentMail' }
    }
    # Notify (High+)
    if ($Severity -in 'High','Critical') {
        $steps += @{ Number = 12; Name = 'Notify' }
    }
    # Report (always)
    $steps += @{ Number = 13; Name = 'Report' }
    return $steps
}

# ============================================================
#  Main entry point
# ============================================================

function Invoke-CompromisedAccountResponse {
    <#
        Run the compromised-account playbook end-to-end. Returns an
        incident id like INC-2026-05-14-a3f2. All downstream
        view/replay/undo functions key off this id.

        See module header for severity matrix + step order.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$UPN,
        [ValidateSet('Low','Medium','High','Critical')][string]$Severity = 'High',
        [string]$Reason,
        [switch]$QuarantineSentMail,
        [switch]$WhatIf,
        [switch]$NonInteractive
    )

    # WhatIf is a convenience flag that forces PREVIEW for the
    # duration of the call without touching the session-wide mode.
    $prevPreview = $false
    if ($WhatIf -and (Get-Command Set-PreviewMode -ErrorAction SilentlyContinue)) {
        $prevPreview = [bool](Get-PreviewMode)
        Set-PreviewMode -Enabled $true
    }

    try {
        if (-not (Connect-ForTask 'Incident')) {
            Write-ErrorMsg "Connect-ForTask failed; cannot run incident response."
            return $null
        }

        $incidentId = New-IncidentId
        $startUtc   = (Get-Date).ToUniversalTime().ToString('o')
        $modeLabel  = if (Get-PreviewMode) { 'PREVIEW' } else { 'LIVE' }

        Write-SectionHeader "Compromised-account response: $UPN"
        Write-StatusLine "Incident"   $incidentId 'White'
        Write-StatusLine "Severity"   $Severity   $(if ($Severity -in 'High','Critical') { 'Red' } else { 'Yellow' })
        Write-StatusLine "Mode"       $modeLabel  $(if ($modeLabel -eq 'PREVIEW') { 'Yellow' } else { 'Red' })
        if ($Reason) { Write-StatusLine "Reason" $Reason 'White' }

        $steps = Get-IncidentSteps -Severity $Severity -QuarantineSentMail:$QuarantineSentMail
        $stepResults = New-Object System.Collections.ArrayList
        $snapshot    = $null
        $audit24h    = $null
        $auditSent   = $null
        $auditShares = $null
        $reportPath  = $null

        # Open registry record up-front (status='running') so a crashed
        # run still leaves a trace.
        $registryRecord = [ordered]@{
            id            = $incidentId
            upn           = $UPN
            severity      = $Severity
            reason        = $Reason
            tenantSlug    = (Get-IncidentTenantSlug)
            tenantName    = if ($script:SessionState) { [string]$script:SessionState.TenantName } else { $null }
            mode          = $modeLabel
            quarantine    = [bool]$QuarantineSentMail
            startedUtc    = $startUtc
            status        = 'running'
            stepsPlanned  = @($steps | ForEach-Object { $_.Number })
            createdBy     = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'unknown' }
        }
        Write-IncidentRegistryRecord -Record $registryRecord
        Write-IncidentAuditEntry -IncidentId $incidentId -EventType 'INCIDENT_START' -Detail ("Compromised-account response started for {0} (severity {1})" -f $UPN, $Severity) -ActionType 'Incident:Start' -Target @{ userUpn = $UPN; severity = $Severity; reason = $Reason } -Result 'info'

        foreach ($step in $steps) {
            $stepName = $step.Name
            $stepNo   = $step.Number
            Write-Host ""
            Write-Host ("  [{0,2}/13] {1}..." -f $stepNo, $stepName) -ForegroundColor Cyan
            $detail = ''
            $status = 'success'
            try {
                switch ($stepName) {
                    'Snapshot' {
                        $snapshot = Get-IncidentSnapshot -IncidentId $incidentId -UPN $UPN
                        $detail = "snapshot.json written; {0} groups, {1} licenses, {2} mfa methods, {3} inbox rules" -f `
                            @($snapshot.groupMemberships).Count, @($snapshot.licenses).Count, @($snapshot.mfaMethods).Count, @($snapshot.inboxRules).Count
                    }
                    'BlockSignIn'         { $r = Invoke-IncidentBlockSignIn -IncidentId $incidentId -UPN $UPN -Snapshot $snapshot;            $detail = "accountEnabled=false"; if (-not $r) { $status = 'failed' } }
                    'RevokeSessions'      { $r = Invoke-IncidentRevokeSessions -IncidentId $incidentId -UPN $UPN -Snapshot $snapshot;         $detail = "all sign-in sessions revoked"; if (-not $r) { $status = 'failed' } }
                    'RevokeAuthMethods'   { $r = Invoke-IncidentRevokeAuthMethods -IncidentId $incidentId -UPN $UPN;                          $detail = "removed $r methods (snapshot has originals)"; if ($null -eq $r) { $status = 'failed' } }
                    'ForcePasswordChange' { $r = Invoke-IncidentForcePasswordChange -IncidentId $incidentId -UPN $UPN -NonInteractive:$NonInteractive; $detail = "forceChangeNextSignIn + new random password"; if (-not $r) { $status = 'failed' } }
                    'DisableInboxRules'   { $r = Invoke-IncidentDisableInboxRules -IncidentId $incidentId -UPN $UPN -Snapshot $snapshot;      $detail = "$r rule(s) disabled" }
                    'ClearForwarding'     { $r = Invoke-IncidentClearForwarding -IncidentId $incidentId -UPN $UPN -Snapshot $snapshot;        $detail = "forwarding cleared" }
                    'Audit24h'            { $audit24h = Invoke-IncidentAudit24h -IncidentId $incidentId -UPN $UPN;                            $detail = "audit-24h.json written" }
                    'AuditSentMail'       { $auditSent = Invoke-IncidentAuditSentMail -IncidentId $incidentId -UPN $UPN;                       $detail = "{0} messages, {1} external recipients" -f @($auditSent.messages).Count, @($auditSent.externalRecipients).Count }
                    'AuditShares'         { $auditShares = Invoke-IncidentAuditShares -IncidentId $incidentId -UPN $UPN;                       $detail = "{0} shares, {1} external domains" -f @($auditShares.shares).Count, @($auditShares.externalDomains).Count }
                    'QuarantineSentMail'  { $r = Invoke-IncidentQuarantineSentMail -IncidentId $incidentId -UPN $UPN -NonInteractive:$NonInteractive; $detail = if ($r) { 'purged' } else { 'declined or non-interactive' }; if (-not $r) { $status = 'skipped' } }
                    'Notify'              { $r = Invoke-IncidentNotify -IncidentId $incidentId -UPN $UPN -Severity $Severity -ReportPath $reportPath; $detail = if ($r) { 'security team notified' } else { 'skipped or failed' }; if (-not $r) { $status = 'skipped' } }
                    'Report'              {
                        $reportPath = New-IncidentReport -IncidentId $incidentId -UPN $UPN -Severity $Severity -Snapshot $snapshot -StepResults @($stepResults) -Audit24h $audit24h -AuditSentMail $auditSent -AuditShares $auditShares
                        $detail = "report.html written"
                    }
                    default { $detail = "unknown step '$stepName'"; $status = 'failed' }
                }
            } catch {
                $status = 'failed'
                $detail = "exception: $($_.Exception.Message)"
                Write-ErrorMsg ("Step {0} ({1}) raised: {2}" -f $stepNo, $stepName, $_.Exception.Message)
                Write-IncidentAuditEntry -IncidentId $incidentId -EventType 'ERROR' -Detail $detail -ActionType ("Incident:$stepName") -Target @{ userUpn = $UPN } -Result 'failure' -ErrorMessage $_.Exception.Message
            }
            [void]$stepResults.Add(@{ StepNumber = $stepNo; StepName = $stepName; Status = $status; Detail = $detail })
            $color = switch ($status) { 'success' { 'Green' } 'skipped' { 'DarkGray' } default { 'Red' } }
            Write-Host ("       -> {0}: {1}" -f $status, $detail) -ForegroundColor $color
        }

        # Close registry record
        $endUtc = (Get-Date).ToUniversalTime().ToString('o')
        $finalRecord = [ordered]@{
            id            = $incidentId
            upn           = $UPN
            severity      = $Severity
            reason        = $Reason
            tenantSlug    = (Get-IncidentTenantSlug)
            tenantName    = if ($script:SessionState) { [string]$script:SessionState.TenantName } else { $null }
            mode          = $modeLabel
            quarantine    = [bool]$QuarantineSentMail
            startedUtc    = $startUtc
            completedUtc  = $endUtc
            status        = 'completed'
            stepsRan      = @($stepResults | ForEach-Object { @{ number=$_.StepNumber; name=$_.StepName; status=$_.Status; detail=$_.Detail } })
            reportPath    = $reportPath
            createdBy     = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { 'unknown' }
        }
        Write-IncidentRegistryRecord -Record $finalRecord
        Write-IncidentAuditEntry -IncidentId $incidentId -EventType 'INCIDENT_END' -Detail ("Compromised-account response completed for {0}" -f $UPN) -ActionType 'Incident:End' -Target @{ userUpn = $UPN; severity = $Severity; status = 'completed' } -Result 'info'

        Write-Host ""
        Write-Success "Incident $incidentId complete."
        Write-StatusLine "Report" $reportPath 'White'
        Write-StatusLine "Dir"    (Get-IncidentDirectory -IncidentId $incidentId) 'DarkGray'
        return $incidentId
    } finally {
        if ($WhatIf -and (Get-Command Set-PreviewMode -ErrorAction SilentlyContinue)) {
            Set-PreviewMode -Enabled $prevPreview
        }
    }
}
