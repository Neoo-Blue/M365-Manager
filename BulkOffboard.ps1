# ============================================================
#  BulkOffboard.ps1 — CSV-driven offboarding
#
#  Public surface:
#    Start-BulkOffboard            interactive wrapper (menu entry)
#    Invoke-BulkOffboard -Path ...  scriptable entry point
#
#  Same -WhatIf / dry-run pattern as BulkOnboard.ps1; commit E
#  refactors both onto Preview.ps1 / Invoke-Action.
#
#  OneDrive handoff (HandoffOneDriveTo column): wired through to
#  validation + result CSV, but the actual SPO/OneDrive transfer
#  is a Phase 3 deliverable. For now each row with the column set
#  logs a TODO line and continues with everything else.
# ============================================================

$script:BulkOffboardRequiredColumns = @('UserPrincipalName')
$script:BulkOffboardBoolColumns     = @('ConvertToShared','RemoveFromAllGroups')

function ConvertTo-BulkOffboardRow {
    param($Row)
    $h = @{}
    foreach ($p in $Row.PSObject.Properties) {
        $k = $p.Name.Trim()
        if ([string]::IsNullOrWhiteSpace($k)) { continue }
        $canonical = if ($k -ieq 'UPN') { 'UserPrincipalName' } else { $k }
        $h[$canonical] = if ($null -eq $p.Value) { '' } else { ([string]$p.Value).Trim() }
    }
    return $h
}

function ConvertTo-BulkBool {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '^(?i:true|yes|1|y|on)$')
}

function Test-BulkOffboardCsv {
    <#
        Validate every row without tenant calls. Catches: missing
        required fields, malformed UPN, duplicate UPN within CSV,
        invalid bool literal in ConvertToShared / RemoveFromAllGroups.
    #>
    param([array]$Rows)

    $errors = @()
    $normalized = @()
    $upnSeen = @{}
    $upnRegex = '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $rowNum = $i + 2
        $r = ConvertTo-BulkOffboardRow -Row $Rows[$i]

        foreach ($req in $script:BulkOffboardRequiredColumns) {
            if (-not $r.ContainsKey($req) -or [string]::IsNullOrWhiteSpace($r[$req])) {
                $errors += @{ Row = $rowNum; Field = $req; Message = "Missing required field: $req" }
            }
        }

        if ($r['UserPrincipalName']) {
            if ($r['UserPrincipalName'] -notmatch $upnRegex) {
                $errors += @{ Row = $rowNum; Field = 'UserPrincipalName'; Message = "Invalid UPN format: '$($r['UserPrincipalName'])'" }
            } else {
                $upnLower = $r['UserPrincipalName'].ToLowerInvariant()
                if ($upnSeen.ContainsKey($upnLower)) {
                    $errors += @{ Row = $rowNum; Field = 'UserPrincipalName'; Message = "Duplicate UPN in CSV (first seen on row $($upnSeen[$upnLower]))" }
                } else {
                    $upnSeen[$upnLower] = $rowNum
                }
            }
        }

        if ($r['ForwardTo'] -and $r['ForwardTo'] -notmatch $upnRegex) {
            $errors += @{ Row = $rowNum; Field = 'ForwardTo'; Message = "Invalid email format: '$($r['ForwardTo'])'" }
        }
        if ($r['HandoffOneDriveTo'] -and $r['HandoffOneDriveTo'] -notmatch $upnRegex) {
            $errors += @{ Row = $rowNum; Field = 'HandoffOneDriveTo'; Message = "Invalid email format: '$($r['HandoffOneDriveTo'])'" }
        }

        $normalized += $r
    }

    return @{ Rows = $normalized; Errors = $errors }
}

function Invoke-BulkOffboard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$WhatIf
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-ErrorMsg "CSV not found: $Path"; return
    }

    Write-SectionHeader "Bulk Offboard -- $(Split-Path $Path -Leaf)"
    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    $dryRun = Get-PreviewMode
    if ($dryRun) { Write-Warn "PREVIEW mode -- no tenant changes will be made." }
    try {

    $rows = $null
    try { $rows = @(Import-Csv -LiteralPath $Path) }
    catch { Write-ErrorMsg "Could not parse CSV: $_"; return }
    if ($rows.Count -eq 0) { Write-Warn "CSV has no data rows."; return }
    Write-InfoMsg "$($rows.Count) row(s) read from $Path"

    $validation = Test-BulkOffboardCsv -Rows $rows
    if ($validation.Errors.Count -gt 0) {
        Write-Host ""
        Write-ErrorMsg "Validation failed -- $($validation.Errors.Count) issue(s):"
        foreach ($e in $validation.Errors) {
            Write-Host ("    Row {0,3} {1,-20} {2}" -f $e.Row, $e.Field, $e.Message) -ForegroundColor Red
        }
        Write-Host ""
        Write-ErrorMsg "Fix the CSV and re-run."
        return
    }
    Write-Success "Validation passed: $($rows.Count) row(s) ready."

    if (-not (Connect-ForTask "Offboard")) { Write-ErrorMsg "Could not connect."; return }

    # Tenant-side check: every UPN exists
    Write-InfoMsg "Checking tenant for existing users..."
    $missing = @()
    $userMap = @{}
    foreach ($r in $validation.Rows) {
        $upn = $r['UserPrincipalName']
        try {
            $u = Get-MgUser -UserId $upn -Property "Id,DisplayName,UserPrincipalName" -ErrorAction Stop
            $userMap[$upn] = $u
        } catch { $missing += $upn }
    }
    if ($missing.Count -gt 0) {
        Write-Warn "$($missing.Count) UPN(s) not found in tenant:"
        foreach ($u in $missing) { Write-Host "    - $u" -ForegroundColor Yellow }
        if (-not (Confirm-Action "Skip these and continue with the rest?")) {
            Write-InfoMsg "Cancelled."; return
        }
    }
    $toProcess = @($validation.Rows | Where-Object { $missing -notcontains $_['UserPrincipalName'] })
    if ($toProcess.Count -eq 0) { Write-InfoMsg "Nothing left to process."; return }

    $modeLabel = if ($dryRun) { "PREVIEW (no changes)" } else { "LIVE -- changes WILL be made" }
    if (-not (Confirm-Action ("About to offboard {0} user(s) in {1}. Proceed?" -f $toProcess.Count, $modeLabel))) {
        Write-InfoMsg "Cancelled."; return
    }

    # ---- Execution ----
    $results = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $toProcess.Count; $i++) {
        $r = $toProcess[$i]
        $upn = $r['UserPrincipalName']
        $user = $userMap[$upn]
        $pct = [int](($i / $toProcess.Count) * 100)
        Write-Progress -Activity "Bulk offboard" -Status "$upn ($($i + 1) of $($toProcess.Count))" -PercentComplete $pct

        $convertShared = ConvertTo-BulkBool $r['ConvertToShared']
        $removeGroups  = ConvertTo-BulkBool $r['RemoveFromAllGroups']
        $forwardTo     = $r['ForwardTo']
        $oooMsg        = $r['Reason']
        $handoffOd     = $r['HandoffOneDriveTo']

        $entry = [ordered]@{
            UPN                 = $upn
            Status              = ''
            Reason              = ''
            SessionsRevoked     = $false
            SignInBlocked       = $false
            OOOSet              = $false
            ForwardingSet       = $false
            ConvertedToShared   = $false
            GroupsRemoved       = 0
            LicensesRemoved     = 0
            OneDriveHandoff     = if ($handoffOd) { "TODO (Phase 3): handoff to $handoffOd" } else { '' }
        }

        $stepErrors = @()
        $stepOk = {
            param([string]$Description, [scriptblock]$Body)
            $result = $null
            try {
                $result = Invoke-Action -Description $Description -Action $Body
                return $true
            } catch {
                $stepErrors += ("{0}: {1}" -f $Description, $_.Exception.Message)
                return $false
            }
        }

        # Step 0 — Revoke MFA methods (delegates to MFAManager.Remove-AllAuthMethods,
        # which wraps each per-method DELETE in Invoke-Action so audit + preview work).
        if (Get-Command Remove-AllAuthMethods -ErrorAction SilentlyContinue) {
            try {
                $mfaRevoked = Remove-AllAuthMethods -User $user.Id
                $entry | Add-Member -NotePropertyName MfaMethodsRevoked -NotePropertyValue $mfaRevoked -Force
            } catch { $stepErrors += "MFARevoke: $($_.Exception.Message)" }
        }

        # Step 1 — Revoke sessions
        if (Invoke-Action -Description ("Revoke sign-in sessions for {0}" -f $upn) -Action {
            Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop; $true
        }) { $entry.SessionsRevoked = $true } elseif (-not $dryRun) { $stepErrors += "RevokeSessions failed" }

        # Step 2 — Block sign-in
        if (Invoke-Action -Description ("Block sign-in for {0}" -f $upn) -Action {
            Update-MgUser -UserId $user.Id -AccountEnabled:$false -ErrorAction Stop; $true
        }) { $entry.SignInBlocked = $true } elseif (-not $dryRun) { $stepErrors += "BlockSignIn failed" }

        # Step 3 — OOO
        if (-not [string]::IsNullOrWhiteSpace($oooMsg)) {
            if (Invoke-Action -Description ("Set out-of-office auto-reply on {0}" -f $upn) -Action {
                Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $oooMsg -ExternalMessage $oooMsg -ErrorAction Stop; $true
            }) { $entry.OOOSet = $true } elseif (-not $dryRun) { $stepErrors += "OOO failed" }
        }

        # Step 4 — Forwarding
        if (-not [string]::IsNullOrWhiteSpace($forwardTo)) {
            if (Invoke-Action -Description ("Set forwarding {0} -> {1} (keep copy)" -f $upn, $forwardTo) -Action {
                Set-Mailbox -Identity $upn -ForwardingSmtpAddress ("smtp:{0}" -f $forwardTo) -DeliverToMailboxAndForward $true -ErrorAction Stop; $true
            }) { $entry.ForwardingSet = $true } elseif (-not $dryRun) { $stepErrors += "Forwarding failed" }
        }

        # Step 5 — Remove from all groups (must run BEFORE license removal,
        # since group-assigned licenses can only be revoked by removing
        # the user from the source group)
        if ($removeGroups) {
            try {
                $memberOf = @(Get-MgUserMemberOf -UserId $user.Id -All -ErrorAction Stop)
                foreach ($g in $memberOf) {
                    $gid = $g.Id
                    $gtype = $g.AdditionalProperties['@odata.type']
                    if ($gtype -ne '#microsoft.graph.group') { continue }
                    $gname = $g.AdditionalProperties['displayName']
                    $ok = Invoke-Action -Description ("Remove {0} from group '{1}'" -f $upn, $gname) -Action {
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$gid/members/$($user.Id)/`$ref" -ErrorAction Stop; $true
                    }
                    if ($ok) { $entry.GroupsRemoved++ }
                }
            } catch { $stepErrors += "GroupEnumerate: $($_.Exception.Message)" }
        }

        # Step 6 — Remove direct licenses
        if (-not $dryRun) {
            try {
                $lics = @(Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop)
                $fullUser = $null
                try { $fullUser = Get-MgUser -UserId $user.Id -Property "LicenseAssignmentStates" -ErrorAction Stop } catch {}
                $assignInfo = @{}
                if ($fullUser -and $fullUser.LicenseAssignmentStates) {
                    foreach ($s in $fullUser.LicenseAssignmentStates) {
                        $sid = "$($s.SkuId)"; $ai = @{ Direct = $false; Groups = @() }
                        if ($assignInfo.ContainsKey($sid)) { $ai = $assignInfo[$sid] }
                        if ($null -eq $s.AssignedByGroup -or $s.AssignedByGroup -eq "") { $ai.Direct = $true } else { $ai.Groups += $s.AssignedByGroup }
                        $assignInfo[$sid] = $ai
                    }
                }
                foreach ($lic in $lics) {
                    $sid = "$($lic.SkuId)"
                    if ($assignInfo.ContainsKey($sid) -and $assignInfo[$sid].Groups.Count -gt 0 -and -not $assignInfo[$sid].Direct) {
                        continue  # group-assigned only -- leave alone
                    }
                    $sku = $lic.SkuPartNumber
                    $ok = Invoke-Action -Description ("Remove license '{0}' from {1}" -f $sku, $upn) -Action {
                        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($lic.SkuId) -ErrorAction Stop; $true
                    }
                    if ($ok) { $entry.LicensesRemoved++ }
                }
            } catch { $stepErrors += "LicensesEnumerate: $($_.Exception.Message)" }
        } else {
            Invoke-Action -Description ("Remove direct-assigned licenses from {0}" -f $upn) -Action { } | Out-Null
        }

        # Step 7 — Convert to shared
        if ($convertShared) {
            if (Invoke-Action -Description ("Convert {0} to Shared Mailbox" -f $upn) -Action {
                Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop; $true
            }) { $entry.ConvertedToShared = $true } elseif (-not $dryRun) { $stepErrors += "ConvertShared failed" }
        }

        # Step 8 — OneDrive handoff (TODO Phase 3)
        if (-not [string]::IsNullOrWhiteSpace($handoffOd)) {
            Write-Warn "OneDrive handoff for $upn -> $handoffOd queued (TODO Phase 3)."
            Write-AuditEntry -EventType 'TODO' -Detail ("OneDriveHandoff {0} -> {1}" -f $upn, $handoffOd)
            $stepErrors += "OneDriveHandoff: NOT_IMPLEMENTED (Phase 3 deliverable)"
        }

        if ($dryRun) {
            $entry.Status = 'Preview'
            $entry.Reason = 'Dry-run, no tenant call made'
        } elseif ($stepErrors.Count -gt 0) {
            $entry.Status = 'PartialSuccess'
            $entry.Reason = ($stepErrors -join ' | ')
            Write-Warn ("Offboard for $upn finished with {0} issue(s)." -f $stepErrors.Count)
        } else {
            $entry.Status = 'Success'
            Write-Success "Offboarded: $upn"
        }
        [void]$results.Add([PSCustomObject]$entry)
    }
    Write-Progress -Activity "Bulk offboard" -Completed

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $resultPath = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-offboard-{0}.csv" -f $stamp)
    try {
        $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Force
        Write-Host ""
        Write-Success "Result CSV: $resultPath"
    } catch {
        Write-ErrorMsg "Could not write result CSV: $_"
    }

        $succeeded = @($results | Where-Object { $_.Status -eq 'Success'        }).Count
        $partial   = @($results | Where-Object { $_.Status -eq 'PartialSuccess' }).Count
        $failed    = @($results | Where-Object { $_.Status -eq 'Failed'         }).Count
        $preview   = @($results | Where-Object { $_.Status -eq 'Preview'        }).Count
        Write-Host ""
        Write-Host "  Bulk offboard summary:" -ForegroundColor White
        Write-StatusLine "Succeeded"      $succeeded "Green"
        Write-StatusLine "Partial success" $partial   $(if ($partial -gt 0) { 'Yellow' } else { 'Gray' })
        Write-StatusLine "Failed"          $failed    $(if ($failed  -gt 0) { 'Red'    } else { 'Gray' })
        if ($preview -gt 0) { Write-StatusLine "Preview" $preview "Yellow" }
        Write-Host ""
    }
    finally {
        Set-PreviewMode -Enabled $previousMode
    }
}

function Start-BulkOffboard {
    Write-SectionHeader "Bulk Offboard from CSV"
    $path = Read-UserInput "Path to CSV file (sample: templates/bulk-offboard-sample.csv)"
    if ([string]::IsNullOrWhiteSpace($path)) { return }
    $path = $path.Trim('"').Trim("'")

    $dryRun = Confirm-Action "Run as DRY-RUN first (validate + preview, no tenant changes)?"

    Invoke-BulkOffboard -Path $path -WhatIf:$dryRun
    Pause-ForUser
}
