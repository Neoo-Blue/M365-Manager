# ============================================================
#  IncidentTriggers.ps1 -- auto-detection of suspicious activity
#
#  Pluggable detectors that surface a Finding (struct: TriggerType,
#  UPN, Severity, Evidence, RecommendedAction). Findings open a
#  Low-severity incident automatically (forensic-only snapshot)
#  and alert the security team via Notifications. Operator decides
#  whether to escalate to a containment response -- no detector
#  auto-runs the destructive playbook unless
#  IncidentResponse.AutoExecuteOnSeverity is set to Critical or
#  HighAndCritical (defaults to None for the conservative default).
# ============================================================

# ============================================================
#  Helpers
# ============================================================

function Get-IncidentTriggerConfig {
    <#
        Return the merged IncidentResponse config hashtable -- the
        ai_config.json IncidentResponse block plus the per-tenant
        overrides resolved via Get-EffectiveConfig. Keys with
        defaults baked in:
          AutoExecuteOnSeverity          = None
          UseAIForNarrative              = Disabled
          SnapshotRetentionDays          = 365
          DetectorIntervalMinutes        = 15
          ImpossibleTravelMaxKmPerHour   = 900
          MassDownloadFileCount          = 50
          MassDownloadWindowMinutes      = 5
          MassShareCount                 = 20
          MassShareWindowMinutes         = 60
          MFAFatigueRejectCount          = 10
          MFAFatigueWindowMinutes        = 60
          AnomalousLocationLookbackDays  = 90
          TabletopUPN                    = $null
    #>
    $defaults = @{
        AutoExecuteOnSeverity         = 'None'
        UseAIForNarrative             = 'Disabled'
        SnapshotRetentionDays         = 365
        DetectorIntervalMinutes       = 15
        ImpossibleTravelMaxKmPerHour  = 900
        MassDownloadFileCount         = 50
        MassDownloadWindowMinutes     = 5
        MassShareCount                = 20
        MassShareWindowMinutes        = 60
        MFAFatigueRejectCount         = 10
        MFAFatigueWindowMinutes       = 60
        AnomalousLocationLookbackDays = 90
        TabletopUPN                   = $null
    }
    $merged = @{}
    foreach ($k in $defaults.Keys) {
        $v = $null
        if (Get-Command Get-EffectiveConfig -ErrorAction SilentlyContinue) {
            $v = Get-EffectiveConfig -Key ("IncidentResponse." + $k)
        }
        if ($null -eq $v -or $v -eq '') { $v = $defaults[$k] }
        $merged[$k] = $v
    }
    return $merged
}

function New-IncidentFinding {
    <#
        Shape every detector's output. Optional Evidence is a free
        hashtable with detector-specific structured fields.
    #>
    param(
        [Parameter(Mandatory)][string]$TriggerType,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][ValidateSet('Low','Medium','High','Critical')][string]$Severity,
        [hashtable]$Evidence,
        [string]$RecommendedAction
    )
    return @{
        TriggerType       = $TriggerType
        UPN               = $UPN
        Severity          = $Severity
        Evidence          = if ($Evidence) { $Evidence } else { @{} }
        RecommendedAction = $RecommendedAction
        DetectedUtc       = (Get-Date).ToUniversalTime().ToString('o')
    }
}

function Get-CountryDistanceKm {
    <#
        Rough country-centroid distance lookup. We don't ship a
        full geocoder; just a few known city/country anchors so
        Detect-ImpossibleTravel can sanity-check obvious cases
        (Lagos vs Sao Paulo, Seattle vs Moscow). Returns $null
        when we don't have anchors for both -- the caller treats
        $null as "can't disprove, don't alert".
    #>
    param([string]$Country1, [string]$Country2)
    if (-not $Country1 -or -not $Country2) { return $null }
    if ($Country1 -eq $Country2) { return 0 }
    $anchors = @{
        'United States'  = @{ Lat = 39.0; Lon = -98.0 }
        'US'             = @{ Lat = 39.0; Lon = -98.0 }
        'Canada'         = @{ Lat = 56.0; Lon = -106.0 }
        'United Kingdom' = @{ Lat = 54.0; Lon = -2.0 }
        'UK'             = @{ Lat = 54.0; Lon = -2.0 }
        'Germany'        = @{ Lat = 51.0; Lon = 10.0 }
        'France'         = @{ Lat = 46.0; Lon = 2.0 }
        'India'          = @{ Lat = 21.0; Lon = 78.0 }
        'China'          = @{ Lat = 35.0; Lon = 105.0 }
        'Japan'          = @{ Lat = 36.0; Lon = 138.0 }
        'Australia'      = @{ Lat = -25.0; Lon = 134.0 }
        'Brazil'         = @{ Lat = -14.0; Lon = -52.0 }
        'Nigeria'        = @{ Lat = 9.0; Lon = 8.0 }
        'Russia'         = @{ Lat = 60.0; Lon = 100.0 }
        'Mexico'         = @{ Lat = 23.0; Lon = -102.0 }
        'South Africa'   = @{ Lat = -29.0; Lon = 24.0 }
        'Egypt'          = @{ Lat = 26.0; Lon = 30.0 }
    }
    $a = $anchors[$Country1]; $b = $anchors[$Country2]
    if (-not $a -or -not $b) { return $null }
    # Haversine on the centroids -- not precise but good enough
    # for "10000km apart in 1 hour is impossible" filtering.
    $r  = 6371.0   # km
    $la1 = $a.Lat * [Math]::PI / 180.0
    $la2 = $b.Lat * [Math]::PI / 180.0
    $dla = ($b.Lat - $a.Lat) * [Math]::PI / 180.0
    $dlo = ($b.Lon - $a.Lon) * [Math]::PI / 180.0
    $h  = [Math]::Sin($dla/2) * [Math]::Sin($dla/2) + [Math]::Cos($la1) * [Math]::Cos($la2) * [Math]::Sin($dlo/2) * [Math]::Sin($dlo/2)
    return $r * 2 * [Math]::Asin([Math]::Min(1.0, [Math]::Sqrt($h)))
}

# ============================================================
#  Detectors
# ============================================================

function Detect-AnomalousLocationSignIn {
    <#
        Sign-in from a country never seen for that user in the
        last AnomalousLocationLookbackDays days. Returns the
        finding hashtable or $null when nothing anomalous.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [array]$SignIns   # optional pre-fetched signIns
    )
    $cfg = Get-IncidentTriggerConfig
    if (-not $SignIns) {
        if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) { return $null }
        try { $SignIns = @(Search-SignIns -User $UPN -From (Get-Date).AddDays(-1 * [int]$cfg.AnomalousLocationLookbackDays) -MaxResults 500) }
        catch { return $null }
    }
    if (@($SignIns).Count -lt 2) { return $null }

    # Treat the most recent sign-in as "today"; everything older
    # is the baseline.
    $sorted = @($SignIns | Sort-Object { [DateTime]$_.CreatedDateTime })
    $latest = $sorted[-1]
    $baseline = $sorted[0..($sorted.Count - 2)]

    # Accept both PSCustomObject (Graph SDK shape) and hashtable
    # (test fixture shape). Both support .Location and .countryOrRegion
    # member access; we just have to handle the nested-hashtable case
    # where Search-SignIns surfaces Location as a structured object.
    $latestCountry = Get-IncidentSignInCountry -SignIn $latest
    if (-not $latestCountry) { return $null }
    $baselineCountries = @{}
    foreach ($s in $baseline) {
        $c = Get-IncidentSignInCountry -SignIn $s
        if ($c) { $baselineCountries[$c] = $true }
    }
    if ($baselineCountries.ContainsKey($latestCountry)) { return $null }

    return New-IncidentFinding -TriggerType 'AnomalousLocationSignIn' -UPN $UPN -Severity 'Low' `
        -Evidence @{
            latestCountry = $latestCountry
            latestUtc     = [string]$latest.CreatedDateTime
            baselineCountries = @($baselineCountries.Keys | Sort-Object)
            lookbackDays  = [int]$cfg.AnomalousLocationLookbackDays
        } `
        -RecommendedAction "Verify the user is travelling. If not, run /incident $UPN High."
}

function Get-IncidentSignInCountry {
    <#
        Extract a country string from a sign-in record, tolerating
        the three shapes we see in practice:
          1. Search-SignIns / Graph SDK shape: .Location is a
             hashtable with .countryOrRegion / .city
          2. Pre-normalized shape: .Location is a string already
          3. Snapshot.json round-trip via ConvertFrom-Json:
             a PSCustomObject with Location as either string or
             nested object.
    #>
    param($SignIn)
    if (-not $SignIn) { return '' }
    $loc = $null
    try { $loc = $SignIn.Location } catch { return '' }
    if ($null -eq $loc) { return '' }
    if ($loc -is [string]) { return $loc }
    # Nested hashtable / PSCustomObject -- look for countryOrRegion
    if ($loc -is [hashtable]) {
        if ($loc.ContainsKey('countryOrRegion')) { return [string]$loc['countryOrRegion'] }
        if ($loc.ContainsKey('country'))         { return [string]$loc['country'] }
        return ''
    }
    try {
        $c = $loc.countryOrRegion
        if ($c) { return [string]$c }
    } catch {}
    return ''
}

function Detect-ImpossibleTravel {
    <#
        Two sign-ins separated by more time-distance than
        physically possible at MaxKmPerHour (default 900 km/h --
        commercial flight buffer). Looks at the last 7 days.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [array]$SignIns
    )
    $cfg = Get-IncidentTriggerConfig
    if (-not $SignIns) {
        if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) { return $null }
        try { $SignIns = @(Search-SignIns -User $UPN -From (Get-Date).AddDays(-7) -MaxResults 500) } catch { return $null }
    }
    if (@($SignIns).Count -lt 2) { return $null }
    $sorted = @($SignIns | Sort-Object { [DateTime]$_.CreatedDateTime })
    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $a = $sorted[$i - 1]; $b = $sorted[$i]
        $cA = Get-IncidentSignInCountry -SignIn $a
        $cB = Get-IncidentSignInCountry -SignIn $b
        if (-not $cA -or -not $cB -or $cA -eq $cB) { continue }
        $km = Get-CountryDistanceKm -Country1 $cA -Country2 $cB
        if ($null -eq $km) { continue }
        $hrs = ([DateTime]$b.CreatedDateTime - [DateTime]$a.CreatedDateTime).TotalHours
        if ($hrs -le 0) { continue }
        $impliedSpeed = $km / $hrs
        if ($impliedSpeed -gt [int]$cfg.ImpossibleTravelMaxKmPerHour) {
            return New-IncidentFinding -TriggerType 'ImpossibleTravel' -UPN $UPN -Severity 'High' `
                -Evidence @{
                    fromCountry        = $cA
                    toCountry          = $cB
                    fromUtc            = [string]$a.CreatedDateTime
                    toUtc              = [string]$b.CreatedDateTime
                    kmApart            = [int]$km
                    hoursApart         = [Math]::Round($hrs, 2)
                    impliedKmPerHour   = [int]$impliedSpeed
                    thresholdKmPerHour = [int]$cfg.ImpossibleTravelMaxKmPerHour
                } `
                -RecommendedAction "Strong compromise indicator. Run /incident $UPN High immediately."
        }
    }
    return $null
}

function Detect-HighRiskSignIn {
    <#
        Identity Protection flagged the sign-in as risk='high'.
        We don't reach into the Identity Protection API directly;
        we rely on Search-SignIns surfacing the RiskLevel field.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [array]$SignIns
    )
    if (-not $SignIns) {
        if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) { return $null }
        try { $SignIns = @(Search-SignIns -User $UPN -From (Get-Date).AddHours(-24) -MaxResults 100) } catch { return $null }
    }
    foreach ($s in @($SignIns)) {
        $risk = [string]$s.RiskLevel
        if ($risk -and $risk.ToLowerInvariant() -eq 'high') {
            return New-IncidentFinding -TriggerType 'HighRiskSignIn' -UPN $UPN -Severity 'High' `
                -Evidence @{
                    signInUtc = [string]$s.CreatedDateTime
                    ipAddress = [string]$s.IpAddress
                    location  = [string]$s.Location
                    appName   = [string]$s.AppDisplayName
                    riskLevel = $risk
                } `
                -RecommendedAction "Identity Protection flagged this. Run /incident $UPN High."
        }
    }
    return $null
}

function Detect-MassFileDownload {
    <#
        >N file-download events from one user within M minutes.
        Reads UAL FileDownloaded operations.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN
    )
    $cfg = Get-IncidentTriggerConfig
    if (-not (Get-Command Search-UAL -ErrorAction SilentlyContinue)) { return $null }
    $windowMin = [int]$cfg.MassDownloadWindowMinutes
    $threshold = [int]$cfg.MassDownloadFileCount
    try {
        $rows = @(Search-UAL -UserId $UPN -From (Get-Date).AddHours(-24) -Operations @('FileDownloaded'))
    } catch { return $null }
    if (@($rows).Count -lt $threshold) { return $null }
    $sorted = @($rows | Sort-Object { [DateTime]$_.CreationDate })
    # Sliding window
    $start = 0
    for ($end = 0; $end -lt $sorted.Count; $end++) {
        while ($start -lt $end -and (([DateTime]$sorted[$end].CreationDate - [DateTime]$sorted[$start].CreationDate).TotalMinutes -gt $windowMin)) {
            $start++
        }
        if (($end - $start + 1) -ge $threshold) {
            return New-IncidentFinding -TriggerType 'MassFileDownload' -UPN $UPN -Severity 'High' `
                -Evidence @{
                    fileCount       = ($end - $start + 1)
                    windowMinutes   = $windowMin
                    windowStartUtc  = [string]$sorted[$start].CreationDate
                    windowEndUtc    = [string]$sorted[$end].CreationDate
                    threshold       = $threshold
                } `
                -RecommendedAction "Possible data exfiltration. Run /incident $UPN High and engage legal/HR if user is leaving."
        }
    }
    return $null
}

function Detect-MassExternalShare {
    <#
        >N outbound shares to external recipients within MassShare
        WindowMinutes. Reuses SharePoint.Get-UserOutboundShares.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN
    )
    $cfg = Get-IncidentTriggerConfig
    if (-not (Get-Command Get-UserOutboundShares -ErrorAction SilentlyContinue)) { return $null }
    try { $shares = @(Get-UserOutboundShares -UPN $UPN -LookbackDays 1) } catch { return $null }
    $userDomain = ($UPN -split '@')[-1].ToLowerInvariant()
    $external = @($shares | Where-Object {
        $t = [string]$_.TargetUserOrEmail
        $t -and $t.Contains('@') -and (($t -split '@')[-1].ToLowerInvariant() -ne $userDomain)
    })
    if ($external.Count -lt [int]$cfg.MassShareCount) { return $null }
    # Find the densest window
    $sorted = @($external | Sort-Object { [DateTime]$_.SharedAtUtc })
    $windowMin = [int]$cfg.MassShareWindowMinutes
    $start = 0
    for ($end = 0; $end -lt $sorted.Count; $end++) {
        while ($start -lt $end -and (([DateTime]$sorted[$end].SharedAtUtc - [DateTime]$sorted[$start].SharedAtUtc).TotalMinutes -gt $windowMin)) {
            $start++
        }
        if (($end - $start + 1) -ge [int]$cfg.MassShareCount) {
            return New-IncidentFinding -TriggerType 'MassExternalShare' -UPN $UPN -Severity 'High' `
                -Evidence @{
                    shareCount     = ($end - $start + 1)
                    windowMinutes  = $windowMin
                    windowStartUtc = [string]$sorted[$start].SharedAtUtc
                    windowEndUtc   = [string]$sorted[$end].SharedAtUtc
                    threshold      = [int]$cfg.MassShareCount
                } `
                -RecommendedAction "Possible data exfiltration via SPO. Run /incident $UPN High."
        }
    }
    return $null
}

function Detect-SuspiciousInboxRule {
    <#
        Rule with the classic AiTM-kit signature: forward-to-
        external + delete-from-sent + a degenerate name (one
        character, '.', or 'rules'). Read-only via Graph.
    #>
    param([Parameter(Mandatory)][string]$UPN)
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UPN/mailFolders/inbox/messageRules" -ErrorAction Stop
    } catch { return $null }
    $userDomain = ($UPN -split '@')[-1].ToLowerInvariant()
    foreach ($r in @($resp.value)) {
        if (-not [bool]$r.isEnabled) { continue }
        $name = [string]$r.displayName
        $nameSuspicious = ($name.Length -le 1) -or ($name -eq '.') -or ($name -eq 'rules') -or ($name -eq '..') -or ($name -eq ' ')
        $forwardsExternal = $false
        if ($r.actions -and $r.actions.forwardTo) {
            foreach ($f in @($r.actions.forwardTo)) {
                $addr = [string]$f.emailAddress.address
                if ($addr -and $addr.Contains('@') -and (($addr -split '@')[-1].ToLowerInvariant() -ne $userDomain)) {
                    $forwardsExternal = $true
                }
            }
        }
        $movesOutOfSight = ($r.actions.delete -eq $true) -or ($r.actions.permanentDelete -eq $true) -or ($r.actions.markAsRead -eq $true -and $r.actions.moveToFolder)

        if (($nameSuspicious -and $forwardsExternal) -or ($forwardsExternal -and $movesOutOfSight)) {
            return New-IncidentFinding -TriggerType 'SuspiciousInboxRule' -UPN $UPN -Severity 'Critical' `
                -Evidence @{
                    ruleId       = [string]$r.id
                    ruleName     = $name
                    nameSuspicious   = $nameSuspicious
                    forwardsExternal = $forwardsExternal
                    movesOutOfSight  = $movesOutOfSight
                } `
                -RecommendedAction "Classic AiTM rule signature. Run /incident $UPN Critical -QuarantineSentMail and consider OAuth grant revocation."
        }
    }
    return $null
}

function Detect-MFAFatigue {
    <#
        >N rejected MFA prompts in MFAFatigueWindowMinutes. UAL
        records this as UserStrongAuthClientAuthNRequiredInterrupt
        or sign-in failures with status 50158 (Authentication
        cancelled by the user).
    #>
    param([Parameter(Mandatory)][string]$UPN)
    $cfg = Get-IncidentTriggerConfig
    if (-not (Get-Command Search-SignIns -ErrorAction SilentlyContinue)) { return $null }
    try { $signIns = @(Search-SignIns -User $UPN -From (Get-Date).AddHours(-1 * ([int]$cfg.MFAFatigueWindowMinutes / 60.0 + 1)) -OnlyFailures -MaxResults 500) } catch { return $null }
    $rejected = @($signIns | Where-Object {
        $status = [string]$_.Status
        # Common codes for user-cancelled MFA / push-bomb rejections
        $status -match '(50158|50074|500121)' -or ($_.RiskLevel -eq 'high')
    })
    if ($rejected.Count -lt [int]$cfg.MFAFatigueRejectCount) { return $null }
    return New-IncidentFinding -TriggerType 'MFAFatigue' -UPN $UPN -Severity 'High' `
        -Evidence @{
            rejectCount    = $rejected.Count
            windowMinutes  = [int]$cfg.MFAFatigueWindowMinutes
            threshold      = [int]$cfg.MFAFatigueRejectCount
        } `
        -RecommendedAction "MFA-fatigue attack signature. Run /incident $UPN High and rotate the user's MFA methods."
}

# ============================================================
#  Driver
# ============================================================

function Invoke-IncidentDetectors {
    <#
        Run every detector against the given list of UPNs (or
        every active user if -All is set). Each Finding opens
        a Low-severity forensic incident automatically and
        alerts the security team. AutoExecuteOnSeverity=None by
        default -- escalation to the containment playbook is
        always operator-driven.
    #>
    [CmdletBinding()]
    param(
        [string[]]$UPNs,
        [switch]$All,
        [switch]$NonInteractive
    )
    if ($All -and -not $UPNs) {
        if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
            Write-ErrorMsg "Graph SDK not loaded -- pass -UPNs explicitly."; return $null
        }
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users" + '?$filter=accountEnabled eq true&$select=userPrincipalName&$top=500') -ErrorAction Stop
            $UPNs = @($resp.value | ForEach-Object { [string]$_.userPrincipalName })
        } catch {
            Write-ErrorMsg "Could not enumerate users: $($_.Exception.Message)"; return $null
        }
    }
    if (-not $UPNs -or $UPNs.Count -eq 0) {
        Write-Warn "No UPNs to scan."; return @()
    }

    $cfg = Get-IncidentTriggerConfig
    $autoExec = [string]$cfg.AutoExecuteOnSeverity
    $findings = New-Object System.Collections.ArrayList

    $detectors = @(
        @{ Name = 'AnomalousLocationSignIn'; Fn = ${function:Detect-AnomalousLocationSignIn} }
        @{ Name = 'ImpossibleTravel';        Fn = ${function:Detect-ImpossibleTravel}        }
        @{ Name = 'HighRiskSignIn';          Fn = ${function:Detect-HighRiskSignIn}          }
        @{ Name = 'MassFileDownload';        Fn = ${function:Detect-MassFileDownload}        }
        @{ Name = 'MassExternalShare';       Fn = ${function:Detect-MassExternalShare}       }
        @{ Name = 'SuspiciousInboxRule';     Fn = ${function:Detect-SuspiciousInboxRule}     }
        @{ Name = 'MFAFatigue';              Fn = ${function:Detect-MFAFatigue}              }
    )

    Write-SectionHeader "Incident detectors"
    Write-StatusLine "Users"                $UPNs.Count       'White'
    Write-StatusLine "Detectors"            $detectors.Count  'White'
    Write-StatusLine "AutoExecute"          $autoExec         $(if ($autoExec -eq 'None') { 'Green' } else { 'Yellow' })

    foreach ($upn in $UPNs) {
        foreach ($d in $detectors) {
            try {
                $finding = & $d.Fn -UPN $upn
            } catch {
                Write-Warn ("Detector {0} on {1} raised: {2}" -f $d.Name, $upn, $_.Exception.Message)
                continue
            }
            if ($finding) {
                [void]$findings.Add($finding)
                Write-Host ("  [{0}] {1} on {2} (severity {3})" -f $finding.TriggerType, $finding.RecommendedAction, $finding.UPN, $finding.Severity) -ForegroundColor Yellow
                # Open a forensic-only Low incident automatically
                $incidentId = $null
                try {
                    $incidentId = Invoke-CompromisedAccountResponse `
                        -UPN $upn `
                        -Severity 'Low' `
                        -Reason ("Auto-trigger: " + $finding.TriggerType) `
                        -NonInteractive
                } catch { Write-Warn "Auto-open Low incident failed: $($_.Exception.Message)" }
                $finding.IncidentId = $incidentId
                # Optional auto-escalation per config
                $shouldEscalate = switch ($autoExec) {
                    'Critical'        { $finding.Severity -eq 'Critical' }
                    'HighAndCritical' { $finding.Severity -in 'High','Critical' }
                    default           { $false }
                }
                if ($shouldEscalate) {
                    Write-Warn ("AUTO-ESCALATING via config (AutoExecuteOnSeverity={0}). Running {1} response for {2}." -f $autoExec, $finding.Severity, $upn)
                    try { Invoke-CompromisedAccountResponse -UPN $upn -Severity $finding.Severity -Reason ("Auto-escalated: " + $finding.TriggerType) -NonInteractive | Out-Null }
                    catch { Write-ErrorMsg "Auto-escalation failed: $($_.Exception.Message)" }
                } else {
                    # Notify only -- operator decides
                    if (Get-Command Send-Notification -ErrorAction SilentlyContinue) {
                        $body = ($finding | ConvertTo-Json -Depth 6)
                        try { Send-Notification -Channels SecurityTeam -Severity $finding.Severity -Subject ("Incident detector: {0} on {1}" -f $finding.TriggerType, $upn) -Body $body | Out-Null }
                        catch { Write-Warn "Notify failed: $($_.Exception.Message)" }
                    }
                }
            }
        }
    }

    Write-Host ""
    Write-Success ("Detector sweep complete: {0} finding(s)." -f $findings.Count)
    return @($findings)
}
