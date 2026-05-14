# ============================================================
#  BreakGlass.ps1 -- break-glass / emergency-access tracking
#
#  Microsoft's recommended posture for break-glass:
#    - Excluded from EVERY conditional-access policy that would
#      block the account during an outage (MFA/legacy-auth/etc).
#    - Strong, unique credential (FIDO2 key recommended).
#    - Quarterly attestation by the security team.
#    - Posture-test cadence: at minimum monthly.
#    - Sign-in activity alerted on -- normal state is "never used".
#
#  This module owns the local registry of designated accounts
#  (<stateDir>/breakglass-accounts.json), the per-account posture
#  predicates, and the alert hook used by the scheduled
#  health-breakglass-signins.ps1 check.
# ============================================================

$script:BGPasswordAgeWarnDays = 180

function Get-BreakGlassStatePath {
    $dir = Get-StateDirectory
    if (-not $dir) { return $null }
    return Join-Path $dir 'breakglass-accounts.json'
}

function Read-BreakGlassRegistry {
    $p = Get-BreakGlassStatePath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return @() }
    try { return @((Get-Content -LiteralPath $p -Raw | ConvertFrom-Json)) } catch { return @() }
}

function Write-BreakGlassRegistry {
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Records)
    $p = Get-BreakGlassStatePath
    if (-not $p) { return }
    # -AsArray (PS 6+) forces [...] even for a single record so the
    # round-trip Read-BreakGlassRegistry -> @(ConvertFrom-Json) keeps
    # an array shape. Without it a 1-record file serializes as a
    # bare object and the next caller hits op_Addition on +=.
    try { ($Records | ConvertTo-Json -Depth 5 -AsArray) | Set-Content -LiteralPath $p -Encoding UTF8 -Force }
    catch { Write-Warn "Could not write break-glass registry: $_" }
}

function Get-BreakGlassAccounts {
    return @(Read-BreakGlassRegistry)
}

function Register-BreakGlassAccount {
    <#
        Designate a UPN as break-glass. Records who designated
        it + when + the attestation contact email. Idempotent
        on UPN (updates the existing record).
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [string]$AttestationEmail
    )
    # @(...) on the return value: PowerShell unwraps single-element
    # arrays during assignment, so without this Read-BreakGlassRegistry
    # returns a PSCustomObject when there's exactly one registered
    # account -- and `$regs += ...` then fails with op_Addition because
    # PSObject doesn't define operator+.
    $regs = @(Read-BreakGlassRegistry)
    $existing = $regs | Where-Object { $_.UPN -eq $UPN } | Select-Object -First 1
    $designatedBy = 'unknown'
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.Account) { $designatedBy = $ctx.Account }
    }
    $now = (Get-Date).ToUniversalTime().ToString('o')
    if ($existing) {
        $existing.AttestationEmail = $AttestationEmail
        $existing.UpdatedAt = $now
        $regs = @($regs | ForEach-Object { if ($_.UPN -eq $UPN) { $existing } else { $_ } })
    } else {
        $regs += [PSCustomObject]@{
            UPN              = $UPN
            DesignatedBy     = $designatedBy
            DesignatedAt     = $now
            LastAttestedAt   = $null
            LastAttestedBy   = $null
            AttestationEmail = $AttestationEmail
            UpdatedAt        = $now
        }
    }
    Write-BreakGlassRegistry -Records $regs
    Write-Success "Registered '$UPN' as break-glass (attest contact: $AttestationEmail)."
}

function Unregister-BreakGlassAccount {
    param([Parameter(Mandatory)][string]$UPN)
    $regs = @(Read-BreakGlassRegistry)
    $remaining = @($regs | Where-Object { $_.UPN -ne $UPN })
    if ($remaining.Count -eq $regs.Count) { Write-Warn "No break-glass record for '$UPN'."; return }
    Write-BreakGlassRegistry -Records $remaining
    Write-Success "Unregistered '$UPN' from break-glass list."
}

# ============================================================
#  Posture checks
# ============================================================

function Test-BreakGlassPosture {
    <#
        Run the per-account posture check. Returns a hashtable
        of @{ warning -> message } for everything that's NOT
        ideal. Empty hashtable means the account is in good
        shape.
    #>
    param([Parameter(Mandatory)][string]$UPN)
    $warnings = @{}

    $user = $null
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN?`$select=id,userPrincipalName,accountEnabled,lastPasswordChangeDateTime,signInActivity" -ErrorAction Stop
    } catch {
        return @{ lookupFailed = "Could not resolve $UPN -- $($_.Exception.Message)" }
    }
    if (-not $user.accountEnabled) {
        $warnings['accountDisabled'] = "Account is currently disabled -- break-glass must be enabled to be useful."
    }

    # Password age
    if ($user.lastPasswordChangeDateTime) {
        $age = ((Get-Date).ToUniversalTime() - ([DateTime]$user.lastPasswordChangeDateTime).ToUniversalTime()).TotalDays
        if ($age -gt $script:BGPasswordAgeWarnDays) {
            $warnings['passwordAge'] = "Password is $([int]$age) days old (warn threshold $script:BGPasswordAgeWarnDays days)."
        }
    } else {
        $warnings['passwordAge'] = "lastPasswordChangeDateTime missing from Graph response."
    }

    # Recent sign-in -- break-glass should normally never sign in
    if ($user.signInActivity.lastSignInDateTime) {
        $daysSince = ((Get-Date).ToUniversalTime() - ([DateTime]$user.signInActivity.lastSignInDateTime).ToUniversalTime()).TotalDays
        if ($daysSince -lt 30) {
            $warnings['recentSignIn'] = "Last sign-in was $([int]$daysSince) days ago -- expected near-zero use for break-glass."
        }
    }

    # FIDO2 + MFA method registration
    if (Get-Command Get-UserAuthMethods -ErrorAction SilentlyContinue) {
        $methods = Get-UserAuthMethods -User $UPN
        $hasFido = ($methods | Where-Object { $_.Label -eq 'FIDO2 key' }).Count -gt 0
        $real    = @($methods | Where-Object { $_.Label -ne 'Password' })
        if ($real.Count -eq 0) { $warnings['noMfaRegistered'] = "No strong auth methods registered." }
        elseif (-not $hasFido) { $warnings['noFido2']        = "No FIDO2 key registered -- recommended for break-glass." }
    }

    # Conditional-access exclusion check
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
        $caExclusions = New-Object System.Collections.ArrayList
        $caRiskyIncludes = New-Object System.Collections.ArrayList
        foreach ($p in $resp.value) {
            if ($p.state -ne 'enabled') { continue }
            $excluded = ($p.conditions.users.excludeUsers -contains $user.id)
            $allUsers = ($p.conditions.users.includeUsers -contains 'All')
            $includesUs = ($p.conditions.users.includeUsers -contains $user.id)
            $criticalShape = ($p.grantControls.builtInControls -contains 'mfa') -or
                             ($p.grantControls.builtInControls -contains 'block')
            if ($criticalShape -and ($allUsers -or $includesUs) -and -not $excluded) {
                [void]$caRiskyIncludes.Add($p.displayName)
            }
            if ($excluded) { [void]$caExclusions.Add($p.displayName) }
        }
        if ($caRiskyIncludes.Count -gt 0) {
            $warnings['caRiskyInclude'] = "Account is INCLUDED in $($caRiskyIncludes.Count) MFA/block CA policies that are NOT excluded: $($caRiskyIncludes -join '; ')"
        }
    } catch { $warnings['caQueryFailed'] = "Could not enumerate CA policies: $($_.Exception.Message)" }

    return $warnings
}

function Test-AllBreakGlassPosture {
    $regs = @(Read-BreakGlassRegistry)
    if ($regs.Count -eq 0) { Write-Warn "No break-glass accounts registered."; return @() }
    $out = New-Object System.Collections.ArrayList
    foreach ($r in $regs) {
        $w = Test-BreakGlassPosture -UPN $r.UPN
        [void]$out.Add([PSCustomObject]@{
            UPN          = $r.UPN
            WarningCount = $w.Count
            Warnings     = ($w.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ' | '
            LastAttested = $r.LastAttestedAt
        })
    }
    return @($out)
}

function Get-BreakGlassSignInActivity {
    <#
        For each registered account, return sign-ins inside the
        last -Days days. Normal state is no rows. Used by the
        scheduled health-breakglass-signins.ps1 check.
    #>
    param([int]$Days = 30)
    $regs = @(Read-BreakGlassRegistry)
    if ($regs.Count -eq 0) { return @() }
    if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) {
        Write-Warn "SignInLookup.ps1 not loaded."
        return @()
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($r in $regs) {
        $rows = @(Search-SignIns -UPN $r.UPN -From ((Get-Date).AddDays(-$Days)) -To (Get-Date) -MaxResults 200)
        foreach ($s in $rows) {
            [void]$out.Add([PSCustomObject]@{
                BreakGlassUpn = $r.UPN
                TimeUtc       = $s.TimeUtc
                App           = $s.App
                IP            = $s.IpAddress
                Outcome       = $s.Outcome
            })
        }
    }
    return @($out | Sort-Object TimeUtc -Descending)
}

function Send-BreakGlassAlert {
    <#
        Fire a security alert via the Notifications framework
        (Phase 4 Commit D). Falls back to console-only when
        Send-Notification isn't loaded yet.
    #>
    param(
        [Parameter(Mandatory)][string]$Account,
        [Parameter(Mandatory)][string]$Event
    )
    $subject = "[BREAK-GLASS ALERT] $Account"
    $body    = "<b>Account:</b> $Account<br><b>Event:</b> $Event<br><b>Time:</b> $((Get-Date).ToUniversalTime().ToString('o'))"
    if (Get-Command Send-Notification -ErrorAction SilentlyContinue) {
        Send-Notification -Channels @('email','teams') -Subject $subject -Body $body -Severity 'Critical' | Out-Null
    } else {
        Write-Warn "[BREAK-GLASS] $Account :: $Event"
        Write-AuditEntry -EventType 'BREAKGLASS_ALERT' -Detail $subject -ActionType 'BreakGlassAlert' -Target @{ upn = $Account; event = $Event } -Result 'info' | Out-Null
    }
}

function Invoke-QuarterlyBreakGlassAttestation {
    <#
        Runs the posture check on every registered account,
        emails the attestation contact for sign-off, and stamps
        LastAttestedAt on the registry. Sign-off is captured
        manually -- this just fires the campaign + records
        intent.
    #>
    $regs = @(Read-BreakGlassRegistry)
    if ($regs.Count -eq 0) { Write-Warn "No break-glass accounts registered."; return }
    foreach ($r in $regs) {
        $w = Test-BreakGlassPosture -UPN $r.UPN
        $body  = "<p>Quarterly attestation for break-glass account <b>$($r.UPN)</b>.</p>"
        if ($w.Count -eq 0) {
            $body += "<p>Posture: GOOD -- no warnings.</p>"
        } else {
            $body += "<p>Posture warnings (please review):</p><ul>"
            foreach ($k in $w.Keys) { $body += "<li><b>$k</b>: $($w[$k])</li>" }
            $body += "</ul>"
        }
        $body += "<p>Reply <b>ATTESTED</b> to confirm this account is still required and that the warnings above are acknowledged.</p>"
        if (Get-Command Send-Notification -ErrorAction SilentlyContinue) {
            Send-Notification -Channels @('email') -Subject "[Break-Glass attestation] $($r.UPN)" -Body $body -Severity 'Warning' -To @($r.AttestationEmail) | Out-Null
        }
        $r.LastAttestedAt = (Get-Date).ToUniversalTime().ToString('o')
        $r.LastAttestedBy = 'pending-reply'
    }
    Write-BreakGlassRegistry -Records $regs
    Write-Success "Quarterly attestation campaign emailed; LastAttestedAt stamps recorded (responses tracked manually)."
}

# ============================================================
#  Menu (lives under Audit & Reporting in Commit F's wiring;
#  for now exposed as Start-BreakGlassMenu)
# ============================================================

function Start-BreakGlassMenu {
    while ($true) {
        $sel = Show-Menu -Title "Break-Glass Accounts" -Options @(
            "List registered accounts",
            "Register an account...",
            "Unregister an account...",
            "Run posture check on one account...",
            "Run posture check on ALL accounts",
            "Sign-in activity (last 30 days)",
            "Run quarterly attestation campaign"
        ) -BackLabel "Back"
        switch ($sel) {
            0 { Get-BreakGlassAccounts | Format-Table -AutoSize; Pause-ForUser }
            1 {
                $u = Read-UserInput "UPN"; if (-not $u) { continue }
                $a = Read-UserInput "Attestation email"
                if ($u) { Register-BreakGlassAccount -UPN $u.Trim() -AttestationEmail $a.Trim() }
                Pause-ForUser
            }
            2 {
                $u = Read-UserInput "UPN to remove"
                if ($u -and (Confirm-Action "Unregister '$u'?")) { Unregister-BreakGlassAccount -UPN $u.Trim() }
                Pause-ForUser
            }
            3 {
                $u = Read-UserInput "UPN"; if (-not $u) { continue }
                if (-not (Connect-ForTask 'Report')) { Pause-ForUser; continue }
                $w = Test-BreakGlassPosture -UPN $u.Trim()
                if ($w.Count -eq 0) { Write-Success "No posture warnings." }
                else {
                    foreach ($k in $w.Keys) { Write-Warn "$k :: $($w[$k])" }
                }
                Pause-ForUser
            }
            4 {
                if (-not (Connect-ForTask 'Report')) { Pause-ForUser; continue }
                Test-AllBreakGlassPosture | Format-Table -AutoSize
                Pause-ForUser
            }
            5 {
                if (-not (Connect-ForTask 'Report')) { Pause-ForUser; continue }
                Get-BreakGlassSignInActivity -Days 30 | Format-Table -AutoSize
                Pause-ForUser
            }
            6 {
                if (-not (Connect-ForTask 'Report')) { Pause-ForUser; continue }
                Invoke-QuarterlyBreakGlassAttestation
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
