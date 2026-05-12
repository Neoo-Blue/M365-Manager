# ============================================================
#  BulkOnboard.ps1 — CSV-driven onboarding
#
#  Public surface:
#    Start-BulkOnboard            interactive wrapper (menu entry)
#    Invoke-BulkOnboard -Path ...  scriptable entry point
#
#  Reuses Apply-Template* from Onboard.ps1 for per-row template
#  application so single-user and bulk paths stay in sync.
#
#  Dry-run / preview wiring in this commit is local (an -WhatIf
#  switch flips a local $dryRun flag). Commit E centralizes this
#  through Preview.ps1 / Invoke-Action.
# ============================================================

$script:BulkOnboardRequiredColumns = @('FirstName','LastName','UserPrincipalName','UsageLocation')

function ConvertTo-BulkRow {
    <#
        Normalize a CSV row (PSCustomObject from Import-Csv) to a
        plain hashtable. Trims whitespace and accepts UPN as an
        alias for UserPrincipalName.
    #>
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

function Test-BulkOnboardCsv {
    <#
        Validate every row in a parsed CSV without making any tenant
        calls. Catches: missing required fields, malformed UPN,
        duplicate UPN within the CSV, unknown template name. Returns:
          @{
            Rows           = @(...)   normalized hashtables
            Errors         = @(...)   @{ Row=int; Field=str; Message=str }
            TemplateCache  = @{}      key -> resolved template hashtable
          }
    #>
    param(
        [array]$Rows,
        [string]$DefaultTemplate,
        [array]$TemplateList
    )

    $errors = @()
    $normalized = @()
    $upnSeen = @{}
    $templateCache = @{}

    $templateKeys = @{}
    foreach ($t in $TemplateList) { $templateKeys[$t.Key.ToLowerInvariant()] = $t }

    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $rowNum = $i + 2   # 1-based + header line
        $r = ConvertTo-BulkRow -Row $Rows[$i]

        foreach ($req in $script:BulkOnboardRequiredColumns) {
            if (-not $r.ContainsKey($req) -or [string]::IsNullOrWhiteSpace($r[$req])) {
                $errors += @{ Row = $rowNum; Field = $req; Message = "Missing required field: $req" }
            }
        }

        if ($r['UserPrincipalName']) {
            if ($r['UserPrincipalName'] -notmatch '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') {
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

        $templateName = if ($r['Template']) { $r['Template'] } else { $DefaultTemplate }
        if ($templateName) {
            $key = ($templateName.ToLowerInvariant() -replace '^role-', '')
            if (-not $templateKeys.ContainsKey($key)) {
                $available = if ($templateKeys.Count -gt 0) { $templateKeys.Keys -join ', ' } else { '(none)' }
                $errors += @{ Row = $rowNum; Field = 'Template'; Message = "Unknown template '$templateName' (available: $available)" }
            } else {
                if (-not $templateCache.ContainsKey($key)) {
                    try { $templateCache[$key] = Get-OnboardTemplate -Key $key }
                    catch { $errors += @{ Row = $rowNum; Field = 'Template'; Message = "Template '$templateName' failed to load: $_" } }
                }
                $r['__TemplateKey'] = $key
            }
        }

        $normalized += $r
    }

    return @{
        Rows          = $normalized
        Errors        = $errors
        TemplateCache = $templateCache
    }
}

function New-BulkOnboardPassword {
    return -join ((48..57) + (65..90) + (97..122) + (33,35,36,37,38) |
                  Get-Random -Count 16 |
                  ForEach-Object { [char]$_ })
}

function New-BulkOnboardBody {
    <#
        Build the New-MgUser body from a normalized CSV row.
        Template fills any field the operator left blank.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Row,
        [hashtable]$Template = $null
    )

    $tplDefaults = @{}
    if ($Template -and $Template['defaults'] -is [hashtable]) {
        $tplDefaults = $Template['defaults']
    }

    function _getField {
        param([string]$Field)
        $v = [string]$Row[$Field]
        if ([string]::IsNullOrWhiteSpace($v) -and $tplDefaults.ContainsKey($Field)) {
            $v = [string]$tplDefaults[$Field]
        }
        return $v
    }

    $upn      = $Row['UserPrincipalName']
    $mailNick = ($upn -split '@')[0]
    $display  = $Row['DisplayName']
    if ([string]::IsNullOrWhiteSpace($display)) {
        $display = ("{0} {1}" -f $Row['FirstName'], $Row['LastName']).Trim()
    }
    $usage = $Row['UsageLocation']
    if ([string]::IsNullOrWhiteSpace($usage) -and $Template) { $usage = [string]$Template['usageLocation'] }

    $body = @{
        AccountEnabled    = $true
        DisplayName       = $display
        GivenName         = $Row['FirstName']
        Surname           = $Row['LastName']
        UserPrincipalName = $upn
        MailNickname      = $mailNick
        UsageLocation     = $usage
    }
    foreach ($f in @('JobTitle','Department','OfficeLocation','CompanyName')) {
        $v = _getField $f
        if (-not [string]::IsNullOrWhiteSpace($v)) { $body[$f] = $v }
    }
    return $body
}

function Invoke-BulkOnboard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Template,
        [switch]$WhatIf
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-ErrorMsg "CSV not found: $Path"; return
    }

    $dryRun = $WhatIf.IsPresent
    Write-SectionHeader "Bulk Onboard -- $(Split-Path $Path -Leaf)"
    if ($dryRun) { Write-Warn "Dry-run / preview mode -- no tenant changes will be made." }

    # ---- Parse ----
    $rows = $null
    try { $rows = @(Import-Csv -LiteralPath $Path) }
    catch { Write-ErrorMsg "Could not parse CSV: $_"; return }
    if ($rows.Count -eq 0) { Write-Warn "CSV has no data rows."; return }
    Write-InfoMsg "$($rows.Count) row(s) read from $Path"

    # ---- Validate (no tenant calls yet) ----
    $templates = @(Get-OnboardTemplates)
    $validation = Test-BulkOnboardCsv -Rows $rows -DefaultTemplate $Template -TemplateList $templates

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

    if (-not (Connect-ForTask "Onboard")) { Write-ErrorMsg "Could not connect."; return }

    # ---- Tenant-side check: UPN already exists ----
    Write-InfoMsg "Checking tenant for existing UPNs..."
    $existingClashes = @()
    foreach ($r in $validation.Rows) {
        $upn = $r['UserPrincipalName']
        try {
            $existing = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
            if ($existing) { $existingClashes += $upn }
        } catch {}
    }
    if ($existingClashes.Count -gt 0) {
        Write-Warn "$($existingClashes.Count) UPN(s) already exist in tenant:"
        foreach ($u in $existingClashes) { Write-Host "    - $u" -ForegroundColor Yellow }
        if (-not (Confirm-Action "Skip these and continue with the rest?")) {
            Write-InfoMsg "Cancelled."; return
        }
    }
    $toProcess = @($validation.Rows | Where-Object { $existingClashes -notcontains $_['UserPrincipalName'] })
    if ($toProcess.Count -eq 0) { Write-InfoMsg "Nothing left to process."; return }

    $modeLabel = if ($dryRun) { "PREVIEW (no changes)" } else { "LIVE -- changes WILL be made" }
    if (-not (Confirm-Action ("About to onboard {0} user(s) in {1}. Proceed?" -f $toProcess.Count, $modeLabel))) {
        Write-InfoMsg "Cancelled."; return
    }

    # ---- Execution ----
    $results = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $toProcess.Count; $i++) {
        $r = $toProcess[$i]
        $upn = $r['UserPrincipalName']
        $pct = [int](($i / $toProcess.Count) * 100)
        Write-Progress -Activity "Bulk onboard" -Status "$upn ($($i + 1) of $($toProcess.Count))" -PercentComplete $pct

        $entry = [ordered]@{
            UPN               = $upn
            Status            = ''
            Reason            = ''
            GeneratedPassword = ''
            TAP               = ''
            TemplateKey       = if ($r['__TemplateKey']) { $r['__TemplateKey'] } else { '' }
        }

        $tpl = $null
        if ($r['__TemplateKey']) { $tpl = $validation.TemplateCache[$r['__TemplateKey']] }

        if ($dryRun) {
            Write-Host ("  [PREVIEW] Would create: {0,-40} display='{1}'" -f $upn, $r['DisplayName']) -ForegroundColor Yellow
            if ($tpl) {
                Write-Host ("  [PREVIEW]    Template: {0}" -f $tpl.name) -ForegroundColor Yellow
                if (@($tpl.licenseSKUs).Count -gt 0)       { Write-Host ("  [PREVIEW]    Licenses: {0}" -f ($tpl.licenseSKUs -join ', ')) -ForegroundColor Yellow }
                if (@($tpl.securityGroups).Count -gt 0)    { Write-Host ("  [PREVIEW]    SGs:      {0}" -f ($tpl.securityGroups -join ', ')) -ForegroundColor Yellow }
                if (@($tpl.distributionLists).Count -gt 0) { Write-Host ("  [PREVIEW]    DLs:      {0}" -f ($tpl.distributionLists -join ', ')) -ForegroundColor Yellow }
                if (@($tpl.sharedMailboxes).Count -gt 0)   { Write-Host ("  [PREVIEW]    SMs:      {0}" -f ((@($tpl.sharedMailboxes) | ForEach-Object { "$($_.identity)($($_.access))" }) -join ', ')) -ForegroundColor Yellow }
            }
            $entry.Status = 'Preview'
            $entry.Reason = 'Dry-run, no tenant call made'
            [void]$results.Add([PSCustomObject]$entry)
            continue
        }

        try {
            $body = New-BulkOnboardBody -Row $r -Template $tpl
            $password = if (-not [string]::IsNullOrWhiteSpace($r['Password'])) { $entry.GeneratedPassword = 'provided'; $r['Password'] } else { $entry.GeneratedPassword = 'generated'; New-BulkOnboardPassword }
            $body['PasswordProfile'] = @{ ForceChangePasswordNextSignIn = $true; Password = $password }

            $newUser = New-MgUser -BodyParameter $body -ErrorAction Stop
            Write-Success "Created: $upn"

            if ($tpl) {
                Start-Sleep -Seconds 3   # tenant propagation before applying licenses
                Apply-TemplateLicenses          -UserId $newUser.Id -Template $tpl
                Apply-TemplateSecurityGroups    -UserId $newUser.Id -Template $tpl
                if (@($tpl.distributionLists).Count -gt 0) { Apply-TemplateDistributionLists -Upn $upn -Template $tpl }
                if (@($tpl.sharedMailboxes).Count   -gt 0) { Apply-TemplateSharedMailboxes   -Upn $upn -Template $tpl }
                if ($tpl.contractorExpiryDays)             { Apply-TemplateContractorExpiry  -UserId $newUser.Id -Template $tpl }
            }

            if ($r['Manager']) {
                try {
                    $mgr = Get-MgUser -UserId $r['Manager'] -ErrorAction Stop
                    Set-MgUserManagerByRef -UserId $newUser.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } -ErrorAction Stop
                    Write-Success "Manager set: $($mgr.UserPrincipalName)"
                } catch {
                    Write-Warn "Manager '$($r['Manager'])' could not be set for $upn -- $($_.Exception.Message)"
                    $entry.Reason = ("ManagerWarn: {0}" -f $_.Exception.Message)
                }
            }

            if ($r['IssueTAP'] -and $r['IssueTAP'] -match '^(?i:true|yes|1|y)$') {
                try {
                    $tapBody = @{ lifetimeInMinutes = 60; isUsableOnce = $true } | ConvertTo-Json -Compress
                    $tapResp = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($newUser.Id)/authentication/temporaryAccessPassMethods" -Body $tapBody -ContentType 'application/json' -ErrorAction Stop
                    $entry.TAP = [string]$tapResp.temporaryAccessPass
                    Write-Success "TAP issued for $upn (one-time, 60 min)."
                } catch {
                    Write-Warn "TAP issue failed for $upn -- $($_.Exception.Message)"
                }
            }

            $entry.Status = 'Success'
        } catch {
            $entry.Status = 'Failed'
            $entry.Reason = $_.Exception.Message
            Write-ErrorMsg "Failed for $upn -- $_"
        }

        [void]$results.Add([PSCustomObject]$entry)
    }
    Write-Progress -Activity "Bulk onboard" -Completed

    # ---- Result CSV ----
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $resultPath = Join-Path (Split-Path -Parent (Resolve-Path $Path)) ("bulk-onboard-{0}.csv" -f $stamp)
    try {
        $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Force
        Write-Host ""
        Write-Success "Result CSV: $resultPath"
    } catch {
        Write-ErrorMsg "Could not write result CSV: $_"
    }

    $succeeded = @($results | Where-Object { $_.Status -eq 'Success' }).Count
    $failed    = @($results | Where-Object { $_.Status -eq 'Failed'  }).Count
    $preview   = @($results | Where-Object { $_.Status -eq 'Preview' }).Count
    Write-Host ""
    Write-Host "  Bulk onboard summary:" -ForegroundColor White
    Write-StatusLine "Succeeded" $succeeded "Green"
    Write-StatusLine "Failed"    $failed    $(if ($failed -gt 0) { 'Red' } else { 'Gray' })
    if ($preview -gt 0) { Write-StatusLine "Preview" $preview "Yellow" }
    Write-Host ""
}

function Start-BulkOnboard {
    Write-SectionHeader "Bulk Onboard from CSV"
    $path = Read-UserInput "Path to CSV file (sample: templates/bulk-onboard-sample.csv)"
    if ([string]::IsNullOrWhiteSpace($path)) { return }
    $path = $path.Trim('"').Trim("'")

    $templates = @(Get-OnboardTemplates)
    $defaultTemplate = $null
    if ($templates.Count -gt 0 -and (Confirm-Action "Apply a default template to rows without their own Template column?")) {
        $labels = $templates | ForEach-Object { "$($_.Name) -- $($_.Description)" }
        $sel = Show-Menu -Title "Default Template" -Options $labels -BackLabel "No default"
        if ($sel -ge 0) { $defaultTemplate = $templates[$sel].Key }
    }

    $dryRun = Confirm-Action "Run as DRY-RUN first (validate + preview, no tenant changes)?"

    Invoke-BulkOnboard -Path $path -Template $defaultTemplate -WhatIf:$dryRun
    Pause-ForUser
}
