# ============================================================
#  MFAManager.ps1 — MFA method inspection + management
#
#  Wraps Graph /users/{id}/authentication/methods endpoints.
#  Required scope: UserAuthenticationMethod.ReadWrite.All
#  (already in $script:MgScopes from Auth.ps1).
#
#  Operations:
#    Get-UserAuthMethods           list methods for one user
#    Remove-AuthMethod             delete one method by id+type
#    Remove-AllAuthMethods         delete every method on a user
#    Reset-UserMFA                 revoke all methods + sign-in
#                                  sessions + issue a fresh TAP
#                                  (so the user can re-register)
#    New-TemporaryAccessPass       single function used by both
#                                  Onboard / BulkOnboard and the
#                                  /Issue TAP menu
#
#  Compliance views:
#    Get-UsersWithOnlyPhoneMfa
#    Get-UsersWithNoMfa
#    Get-UsersWithActiveTap
#    Get-UsersWithFido2Keys
#  Each prompts for a scan-size cap (full-tenant enumeration is
#  slow). Output: table + CSV in the audit dir.
#
#  Mutations are wrapped in Invoke-Action so PREVIEW works and
#  every operation lands in the audit log; destructive operations
#  (revoke methods, reset MFA) carry NoUndoReason since the user
#  has to re-register the method anyway -- re-registering is
#  manual, not a recipe we can replay.
# ============================================================

$script:MfaMethodTypeMap = @{
    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'           = @{ Label='Microsoft Authenticator'; Segment='microsoftAuthenticatorMethods' }
    '#microsoft.graph.phoneAuthenticationMethod'                            = @{ Label='Phone (SMS/voice)';       Segment='phoneMethods' }
    '#microsoft.graph.emailAuthenticationMethod'                            = @{ Label='Email';                   Segment='emailMethods' }
    '#microsoft.graph.fido2AuthenticationMethod'                            = @{ Label='FIDO2 key';               Segment='fido2Methods' }
    '#microsoft.graph.temporaryAccessPassAuthenticationMethod'              = @{ Label='Temporary Access Pass';   Segment='temporaryAccessPassMethods' }
    '#microsoft.graph.softwareOathAuthenticationMethod'                     = @{ Label='Software OATH (TOTP)';    Segment='softwareOathMethods' }
    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod'          = @{ Label='Windows Hello';           Segment='windowsHelloForBusinessMethods' }
    '#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod' = @{ Label='Passwordless MS Auth';  Segment='passwordlessMicrosoftAuthenticatorMethods' }
    '#microsoft.graph.passwordAuthenticationMethod'                         = @{ Label='Password';                Segment='' }   # cannot be deleted via Graph
}

function Get-MfaMethodInfo {
    param($Method)
    $odata = [string]$Method.'@odata.type'
    if ($script:MfaMethodTypeMap.ContainsKey($odata)) {
        $info = $script:MfaMethodTypeMap[$odata]
        return [PSCustomObject]@{
            Type        = $odata
            Label       = $info.Label
            UrlSegment  = $info.Segment
            Id          = $Method.id
            Detail      = (Format-MfaMethodDetail $Method)
            Raw         = $Method
        }
    }
    return [PSCustomObject]@{
        Type       = $odata
        Label      = ($odata -replace '#microsoft.graph.','' -replace 'AuthenticationMethod','')
        UrlSegment = ''
        Id         = $Method.id
        Detail     = '(unknown method type)'
        Raw        = $Method
    }
}

function Format-MfaMethodDetail {
    param($Method)
    $odata = [string]$Method.'@odata.type'
    switch -Regex ($odata) {
        'phoneAuthenticationMethod'              { return "{0} ({1})" -f $Method.phoneNumber, $Method.phoneType }
        'microsoftAuthenticator'                 { return "device: $($Method.displayName)" }
        'emailAuthenticationMethod'              { return "address: $($Method.emailAddress)" }
        'fido2AuthenticationMethod'              { return "model: $($Method.model); aaguid: $($Method.aaGuid)" }
        'temporaryAccessPass'                    { return "lifetime: $($Method.lifetimeInMinutes)m; usable: $($Method.isUsableOnce); state: $($Method.methodUsabilityReason)" }
        'softwareOathAuthenticationMethod'       { return "device: $($Method.secretKey)" }
        'windowsHelloForBusiness'                { return "device: $($Method.displayName)" }
        default                                  { return '' }
    }
}

function Get-UserAuthMethods {
    <#
        List all authentication methods for one user (UPN or id).
        Returns an array of PSCustomObjects: Type, Label, Id,
        UrlSegment, Detail, Raw.
    #>
    param([Parameter(Mandatory)][string]$User)
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$User/authentication/methods" -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Could not enumerate methods for '$User': $($_.Exception.Message)"
        return @()
    }
    $out = @()
    foreach ($m in $resp.value) { $out += Get-MfaMethodInfo -Method $m }
    return $out
}

function Show-UserAuthMethods {
    param([Parameter(Mandatory)][string]$User)
    Write-SectionHeader "Authentication methods for $User"
    $methods = Get-UserAuthMethods -User $User
    if ($methods.Count -eq 0) {
        Write-Warn "No methods returned (user may not exist, or no methods registered, or insufficient scope)."
        return @()
    }
    Write-Host ""
    for ($i = 0; $i -lt $methods.Count; $i++) {
        $m = $methods[$i]
        $colour = if ($m.Label -eq 'Phone (SMS/voice)') { 'Yellow' } elseif ($m.Label -eq 'FIDO2 key' -or $m.Label -eq 'Microsoft Authenticator') { 'Green' } else { 'White' }
        Write-Host ("  [{0,2}] {1,-28} {2}" -f ($i + 1), $m.Label, $m.Detail) -ForegroundColor $colour
        Write-Host ("       id: {0}" -f $m.Id) -ForegroundColor DarkGray
    }
    Write-Host ""
    return $methods
}

function Remove-AuthMethod {
    <#
        Delete one method by its URL segment + id. Wrapped in
        Invoke-Action so the operation is auditable and PREVIEW-
        able.
    #>
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)]$MethodInfo
    )
    if (-not $MethodInfo.UrlSegment) {
        Write-Warn "Method type '$($MethodInfo.Label)' cannot be removed via Graph (likely the password method)."
        return $false
    }
    return [bool] (Invoke-Action `
        -Description ("Revoke MFA method '{0}' for {1}" -f $MethodInfo.Label, $User) `
        -ActionType 'RevokeAuthMethod' `
        -Target @{ userUpn = $User; methodType = $MethodInfo.Label; methodId = $MethodInfo.Id } `
        -NoUndoReason 'Auth method revocation cannot be undone via API; the user must re-register the method.' `
        -Action {
            $uri = "https://graph.microsoft.com/v1.0/users/$User/authentication/$($MethodInfo.UrlSegment)/$($MethodInfo.Id)"
            Invoke-MgGraphRequest -Method DELETE -Uri $uri -ErrorAction Stop | Out-Null
            $true
        })
}

function Remove-AllAuthMethods {
    param([Parameter(Mandatory)][string]$User)
    $methods = Get-UserAuthMethods -User $User
    if ($methods.Count -eq 0) { Write-InfoMsg "No methods to revoke."; return 0 }
    $count = 0
    foreach ($m in $methods) {
        if (-not $m.UrlSegment) { continue }   # skip Password
        if (Remove-AuthMethod -User $User -MethodInfo $m) { $count++ }
    }
    return $count
}

function New-TemporaryAccessPass {
    <#
        Create a Temporary Access Pass on a user. Single source of
        truth for TAP issuance -- BulkOnboard and the MFA menu
        both call here.

        -LifetimeMinutes (1..480)  -- default 60
        -IsUsableOnce              -- default $true (single use)
        Returns the JSON response object (so callers can read .temporaryAccessPass).
    #>
    param(
        [Parameter(Mandatory)][string]$User,
        [int]$LifetimeMinutes = 60,
        [bool]$IsUsableOnce = $true
    )
    $body = @{ lifetimeInMinutes = $LifetimeMinutes; isUsableOnce = $IsUsableOnce } | ConvertTo-Json -Compress
    return (Invoke-Action `
        -Description ("Issue Temporary Access Pass for {0} ({1} min, usable-once={2})" -f $User, $LifetimeMinutes, $IsUsableOnce) `
        -ActionType 'IssueTAP' `
        -Target @{ userUpn = $User; lifetimeMinutes = $LifetimeMinutes; isUsableOnce = $IsUsableOnce } `
        -NoUndoReason 'TAP issuance creates a transient credential; revocation is via Remove-AuthMethod on the new method id.' `
        -StubReturn ([PSCustomObject]@{ temporaryAccessPass = '<preview-tap-redacted>' }) `
        -Action {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$User/authentication/temporaryAccessPassMethods" -Body $body -ContentType 'application/json' -ErrorAction Stop
        })
}

function Reset-UserMFA {
    <#
        Force-reregistration sequence:
          1. revoke every existing method (except Password)
          2. revoke all sign-in sessions
          3. issue a new TAP (60 min, single use by default)
    #>
    param(
        [Parameter(Mandatory)][string]$User,
        [int]$TAPLifetime = 60,
        [bool]$TAPUsableOnce = $true
    )
    Write-SectionHeader "Reset MFA for $User"
    Write-Warn "This will lock $User out until they redeem the TAP and re-register at least one strong factor."
    if (-not (Confirm-Action "Proceed?")) { return $null }

    $removed = Remove-AllAuthMethods -User $User
    Write-InfoMsg "Methods revoked: $removed"

    $userObj = $null
    try { $userObj = Get-MgUser -UserId $User -ErrorAction Stop } catch { Write-Warn "Could not resolve UserId for session revocation: $_" }
    if ($userObj) {
        Invoke-Action `
            -Description ("Revoke sign-in sessions for {0}" -f $User) `
            -ActionType 'RevokeSignInSessions' `
            -Target @{ userId = [string]$userObj.Id; userUpn = $User } `
            -NoUndoReason 'Sign-in sessions cannot be un-revoked; user re-signs in fresh.' `
            -Action { Revoke-MgUserSignInSession -UserId $userObj.Id -ErrorAction Stop; $true } | Out-Null
    }

    $tap = New-TemporaryAccessPass -User $User -LifetimeMinutes $TAPLifetime -IsUsableOnce $TAPUsableOnce
    if ($tap -and $tap.temporaryAccessPass) {
        Write-Host ""
        Write-Host "  TAP: " -ForegroundColor Yellow -NoNewline
        if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
            try { Set-Clipboard -Value $tap.temporaryAccessPass } catch {}
            Write-Host "<copied to clipboard>" -ForegroundColor Yellow
            Write-Warn "Deliver to the user via a secure channel; this will not be shown again."
        } else {
            Write-Host $tap.temporaryAccessPass -ForegroundColor Yellow
            Write-Warn "Clipboard unavailable -- screen will clear shortly."
        }
        Write-Host ""
        $null = Read-UserInput "Press Enter when delivered"
        try { Set-Clipboard -Value ' ' -ErrorAction SilentlyContinue } catch {}
        Clear-Host
        Write-Success "MFA reset complete for $User."
    }
    return $tap
}

# ============================================================
#  Compliance views — slow when run against the whole tenant.
#  Each prompts the operator for a scan cap so they can iterate.
# ============================================================

function Invoke-MfaComplianceScan {
    <#
        Iterate over up to -Max users (Get-MgUser, sorted by UPN),
        enumerate their auth methods, and run a per-user predicate.
        Returns the matching users with their method summary.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Predicate,
        [int]$Max = 200
    )
    Write-InfoMsg "Enumerating up to $Max users..."
    $users = @()
    try { $users = @(Get-MgUser -Top $Max -Property "Id,UserPrincipalName,DisplayName,AccountEnabled" -ErrorAction Stop) }
    catch { Write-ErrorMsg "Could not enumerate users: $_"; return @() }
    Write-InfoMsg "Inspecting $($users.Count) user(s) for MFA state..."

    $hits = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($u in $users) {
        $i++
        Write-Progress -Activity "MFA compliance scan" -Status "$($u.UserPrincipalName)" -PercentComplete (($i / $users.Count) * 100)
        $methods = Get-UserAuthMethods -User $u.Id
        if (& $Predicate $methods) {
            [void]$hits.Add([PSCustomObject]@{
                UserPrincipalName = $u.UserPrincipalName
                DisplayName       = $u.DisplayName
                AccountEnabled    = $u.AccountEnabled
                MethodCount       = $methods.Count
                Methods           = (($methods | ForEach-Object { $_.Label }) -join ', ')
            })
        }
    }
    Write-Progress -Activity "MFA compliance scan" -Completed
    return @($hits)
}

function Get-UsersWithOnlyPhoneMfa {
    param([int]$Max = 200)
    Invoke-MfaComplianceScan -Max $Max -Predicate {
        param($methods)
        $real = @($methods | Where-Object { $_.Label -ne 'Password' })
        if ($real.Count -eq 0) { return $false }
        $allPhone = ($real | Where-Object { $_.Label -ne 'Phone (SMS/voice)' }).Count -eq 0
        return $allPhone
    }
}

function Get-UsersWithNoMfa {
    param([int]$Max = 200)
    Invoke-MfaComplianceScan -Max $Max -Predicate {
        param($methods)
        $real = @($methods | Where-Object { $_.Label -ne 'Password' })
        return $real.Count -eq 0
    }
}

function Get-UsersWithActiveTap {
    param([int]$Max = 200)
    Invoke-MfaComplianceScan -Max $Max -Predicate {
        param($methods)
        return ($methods | Where-Object { $_.Label -eq 'Temporary Access Pass' }).Count -gt 0
    }
}

function Get-UsersWithFido2Keys {
    param([int]$Max = 200)
    Invoke-MfaComplianceScan -Max $Max -Predicate {
        param($methods)
        return ($methods | Where-Object { $_.Label -eq 'FIDO2 key' }).Count -gt 0
    }
}

function Export-MfaComplianceCsv {
    param([Parameter(Mandatory)][array]$Records, [Parameter(Mandatory)][string]$Tag)
    if ($Records.Count -eq 0) { Write-InfoMsg "(no rows to export)"; return }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $path = Join-Path (Get-AuditLogDirectory) "mfa-$Tag-$stamp.csv"
    $Records | Export-Csv -LiteralPath $path -NoTypeInformation -Force
    Write-Success "CSV: $path"
}

# ============================================================
#  Bulk MFA from CSV
# ============================================================

function Invoke-BulkMfa {
    <#
        CSV columns: UPN, Action (Revoke | Reset | IssueTAP)
        Optional: TAPLifetime (int minutes), TAPUsableOnce (bool)
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [switch]$WhatIf)
    if (-not (Test-Path -LiteralPath $Path)) { Write-ErrorMsg "CSV not found: $Path"; return }
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { Write-Warn "Empty CSV."; return }

    $previousMode = Get-PreviewMode
    if ($WhatIf.IsPresent -and -not $previousMode) { Set-PreviewMode -Enabled $true }
    try {
        $results = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $r = $rows[$i]
            $upn = [string]$r.UPN
            $act = [string]$r.Action
            Write-Progress -Activity "Bulk MFA" -Status "$upn -- $act" -PercentComplete (($i / $rows.Count) * 100)
            $status = 'Pending'
            try {
                switch ($act.ToLower()) {
                    'revoke'    { $n = Remove-AllAuthMethods -User $upn; $status = "Revoked $n method(s)" }
                    'reset'     { $tap = Reset-UserMFA -User $upn; $status = if ($tap) { 'Reset + TAP issued' } else { 'Reset cancelled' } }
                    'issuetap'  {
                        $life  = if ($r.TAPLifetime) { [int]$r.TAPLifetime } else { 60 }
                        $usable= if ($r.TAPUsableOnce) { [bool]::Parse([string]$r.TAPUsableOnce) } else { $true }
                        $tap = New-TemporaryAccessPass -User $upn -LifetimeMinutes $life -IsUsableOnce $usable
                        $status = if ($tap) { 'TAP issued' } else { 'failed' }
                    }
                    default     { $status = "Unknown action: $act" }
                }
            } catch { $status = "ERROR: $($_.Exception.Message)" }
            [void]$results.Add([PSCustomObject]@{ UPN = $upn; Action = $act; Status = $status })
        }
        Write-Progress -Activity "Bulk MFA" -Completed
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $out = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-mfa-$stamp.csv")
        $results | Export-Csv -LiteralPath $out -NoTypeInformation -Force
        Write-Success "Result CSV: $out"
    } finally {
        Set-PreviewMode -Enabled $previousMode
    }
}

# ============================================================
#  Menu
# ============================================================

function Start-MFAMenu {
    while ($true) {
        $sel = Show-Menu -Title "MFA & Authentication" -Options @(
            "View user's methods",
            "Revoke specific method...",
            "Revoke all methods...",
            "Reset MFA (revoke all + new TAP)",
            "Issue Temporary Access Pass...",
            "Compliance: users with only SMS/voice",
            "Compliance: users with no MFA",
            "Compliance: users with active TAP",
            "Compliance: users with FIDO2 keys",
            "Bulk MFA from CSV..."
        ) -BackLabel "Back"

        switch ($sel) {
            0 { $u = Read-UserInput "User UPN"; if ($u) { Show-UserAuthMethods -User $u; Pause-ForUser } }
            1 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                $methods = Show-UserAuthMethods -User $u
                if ($methods.Count -eq 0) { Pause-ForUser; continue }
                $idxText = Read-UserInput "Row to revoke (1-$($methods.Count))"
                $idx = 0
                if ([int]::TryParse($idxText,[ref]$idx) -and $idx -ge 1 -and $idx -le $methods.Count) {
                    if (Confirm-Action "Revoke '$($methods[$idx-1].Label)' for $u?") {
                        Remove-AuthMethod -User $u -MethodInfo $methods[$idx-1] | Out-Null
                    }
                }
                Pause-ForUser
            }
            2 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                Write-Warn "Revoking ALL methods will lock the user out until re-registration."
                if (Confirm-Action "Proceed?") { $n = Remove-AllAuthMethods -User $u; Write-Success "Revoked $n method(s)." }
                Pause-ForUser
            }
            3 { $u = Read-UserInput "User UPN"; if ($u) { Reset-UserMFA -User $u | Out-Null }; Pause-ForUser }
            4 {
                $u = Read-UserInput "User UPN"; if (-not $u) { continue }
                $lifeSel = Show-Menu -Title "Lifetime" -Options @("1 hour","8 hours","24 hours") -BackLabel "Cancel"
                if ($lifeSel -eq -1) { continue }
                $life = @(60, 480, 1440)[$lifeSel]
                $usable = (Show-Menu -Title "Usable" -Options @("Once","Reusable until expiry") -BackLabel "Cancel") -eq 0
                $tap = New-TemporaryAccessPass -User $u -LifetimeMinutes $life -IsUsableOnce $usable
                if ($tap -and $tap.temporaryAccessPass -and -not (Get-PreviewMode)) {
                    if (Get-Command Set-Clipboard -ErrorAction SilentlyContinue) {
                        try { Set-Clipboard -Value $tap.temporaryAccessPass } catch {}
                        Write-Host "  TAP copied to clipboard." -ForegroundColor Yellow
                    } else {
                        Write-Host "  TAP: $($tap.temporaryAccessPass)" -ForegroundColor Yellow
                    }
                    $null = Read-UserInput "Press Enter when delivered"
                    try { Set-Clipboard -Value ' ' -ErrorAction SilentlyContinue } catch {}
                    Clear-Host
                }
                Pause-ForUser
            }
            5 {
                $maxT = Read-UserInput "Scan first N users (default 200)"
                $max = 200; [int]::TryParse($maxT,[ref]$max) | Out-Null
                $rows = Get-UsersWithOnlyPhoneMfa -Max $max
                Write-InfoMsg "Found $($rows.Count) user(s) with ONLY phone-based MFA."
                $rows | Format-Table -AutoSize
                if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { Export-MfaComplianceCsv -Records $rows -Tag 'only-phone' }
                Pause-ForUser
            }
            6 {
                $maxT = Read-UserInput "Scan first N users (default 200)"
                $max = 200; [int]::TryParse($maxT,[ref]$max) | Out-Null
                $rows = Get-UsersWithNoMfa -Max $max
                Write-InfoMsg "Found $($rows.Count) user(s) with NO MFA registered."
                $rows | Format-Table -AutoSize
                if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { Export-MfaComplianceCsv -Records $rows -Tag 'no-mfa' }
                Pause-ForUser
            }
            7 {
                $maxT = Read-UserInput "Scan first N users (default 200)"
                $max = 200; [int]::TryParse($maxT,[ref]$max) | Out-Null
                $rows = Get-UsersWithActiveTap -Max $max
                Write-InfoMsg "Found $($rows.Count) user(s) with an active TAP."
                $rows | Format-Table -AutoSize
                if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { Export-MfaComplianceCsv -Records $rows -Tag 'active-tap' }
                Pause-ForUser
            }
            8 {
                $maxT = Read-UserInput "Scan first N users (default 200)"
                $max = 200; [int]::TryParse($maxT,[ref]$max) | Out-Null
                $rows = Get-UsersWithFido2Keys -Max $max
                Write-InfoMsg "Found $($rows.Count) user(s) with FIDO2 keys."
                $rows | Format-Table -AutoSize
                if ($rows.Count -gt 0 -and (Confirm-Action "Export CSV?")) { Export-MfaComplianceCsv -Records $rows -Tag 'fido2' }
                Pause-ForUser
            }
            9 {
                $p = Read-UserInput "Path to CSV (columns: UPN, Action[, TAPLifetime, TAPUsableOnce])"
                if (-not $p) { continue }
                $dry = Confirm-Action "Run as DRY-RUN first?"
                Invoke-BulkMfa -Path $p.Trim('"').Trim("'") -WhatIf:$dry
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
