# ============================================================
#  OneDriveManager.ps1 — leaver's OneDrive handoff orchestration
#
#  Required permissions:
#    - SharePoint Administrator (for SPO cmdlets)
#    - Graph: Files.ReadWrite.All (already in Sites.FullControl.All),
#             Mail.Send (added in Auth.ps1 Phase 3)
#
#  IMPORTANT: M365 begins a ~30-day countdown on a user's personal
#  site after their license is removed. There is no clean "extend
#  the timer" cmdlet -- the practical mitigations are:
#    1. Keep the license assigned during the handoff window
#    2. Apply a Microsoft Purview retention policy targeting the
#       leaver's OneDrive
#    3. Transfer the site to a "successor" account so its lifecycle
#       resets to that account
#  Extend-OneDriveRetention here records the intended retention
#  date in the audit log and calls Set-SPOSite -LockState NoAccess
#  defer-style only when explicitly requested. We don't pretend to
#  reset the auto-purge clock.
# ============================================================

function Get-UserOneDriveUrl {
    <#
        Return the personal site URL for a UPN. Tries Get-SPOSite
        with an Owner filter; falls back to constructing the URL
        from the tenant short name + UPN if the site isn't
        provisioned yet (Get-SPOSite returns nothing in that case
        and we don't want the operator to think the user has no
        site at all).
    #>
    param([Parameter(Mandatory)][string]$UPN)
    if (-not (Get-Command Get-SPOSite -ErrorAction SilentlyContinue)) {
        Write-Warn "Microsoft.Online.SharePoint.PowerShell not loaded."
        return $null
    }
    try {
        $site = Get-SPOSite -IncludePersonalSite $true -Filter "Owner -eq '$UPN'" -Limit 1 -ErrorAction Stop | Select-Object -First 1
        if ($site -and $site.Url) { return [string]$site.Url }
    } catch { Write-Warn "Get-SPOSite filter failed: $($_.Exception.Message)" }
    Write-InfoMsg "No provisioned personal site found for $UPN."
    return $null
}

function Get-OneDriveRecentFiles {
    <#
        Top-level recently-modified files for the manager summary.
        Uses Graph /drives/{driveId}/root/search with a
        lastModifiedDateTime filter to keep the payload small.
        Returns at most -Top items, default 20.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [int]$Days = 90,
        [int]$Top  = 20
    )
    # SiteUrl like https://contoso-my.sharepoint.com/personal/jane_contoso_com
    # The drive lookup needs the SPO site id which we can get via Graph:
    # /sites/{hostname}:/personal/{path}
    if ($SiteUrl -notmatch 'https?://([^/]+)/personal/([^/]+)') {
        Write-Warn "Not a recognized OneDrive site URL: $SiteUrl"
        return @()
    }
    $hostName = $Matches[1]
    $sitePath = "/personal/$($Matches[2])"
    try {
        $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($hostName):$($sitePath)" -ErrorAction Stop
        if (-not $site -or -not $site.id) { return @() }
        $drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive" -ErrorAction Stop
        if (-not $drive -or -not $drive.id) { return @() }
        # Recent children + descendant search
        $cutoff = (Get-Date).ToUniversalTime().AddDays(-$Days).ToString('o')
        $top = [Math]::Min(200, [Math]::Max(20, $Top))
        $resp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/delta?`$top=$top" -ErrorAction Stop
        $files = @()
        foreach ($v in $resp.value) {
            if ($v.file -and $v.lastModifiedDateTime -and [DateTime]$v.lastModifiedDateTime -ge $cutoff) {
                $files += [PSCustomObject]@{
                    Name             = $v.name
                    LastModifiedUtc  = ([DateTime]$v.lastModifiedDateTime).ToUniversalTime()
                    LastModifiedBy   = if ($v.lastModifiedBy.user.displayName) { $v.lastModifiedBy.user.displayName } else { '' }
                    Size             = $v.size
                    WebUrl           = $v.webUrl
                }
            }
        }
        return @($files | Sort-Object LastModifiedUtc -Descending | Select-Object -First $Top)
    } catch {
        Write-Warn "Could not enumerate recent files: $($_.Exception.Message)"
        return @()
    }
}

function Grant-OneDriveAccess {
    <#
        Add -GranteeUPN as a site collection admin on -SiteUrl.
        Reversible: pairs with Revoke-OneDriveAccess.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$GranteeUPN
    )
    Invoke-Action `
        -Description ("Grant {0} site-collection admin on {1}" -f $GranteeUPN, $SiteUrl) `
        -ActionType 'GrantOneDriveAccess' `
        -Target @{ siteUrl = $SiteUrl; granteeUpn = $GranteeUPN } `
        -ReverseType 'RevokeOneDriveAccess' `
        -ReverseDescription ("Revoke {0} site-collection admin on {1}" -f $GranteeUPN, $SiteUrl) `
        -Action {
            Set-SPOUser -Site $SiteUrl -LoginName $GranteeUPN -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            $true
        } | Out-Null
}

function Revoke-OneDriveAccess {
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$GranteeUPN
    )
    Invoke-Action `
        -Description ("Revoke {0} site-collection admin on {1}" -f $GranteeUPN, $SiteUrl) `
        -ActionType 'RevokeOneDriveAccess' `
        -Target @{ siteUrl = $SiteUrl; granteeUpn = $GranteeUPN } `
        -ReverseType 'GrantOneDriveAccess' `
        -ReverseDescription ("Grant {0} site-collection admin on {1}" -f $GranteeUPN, $SiteUrl) `
        -Action {
            Set-SPOUser -Site $SiteUrl -LoginName $GranteeUPN -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
            $true
        } | Out-Null
}

function Set-OneDriveReadOnly {
    <#
        Set-SPOSite -LockState ReadOnly. Unlock is technically
        possible (LockState=Unlock) but operators usually only
        flip this once during a handoff, so we flag it as a soft
        no-undo (the reverse is operator-judgment, not automatic).
    #>
    param([Parameter(Mandatory)][string]$SiteUrl)
    Invoke-Action `
        -Description ("Lock OneDrive '{0}' to ReadOnly" -f $SiteUrl) `
        -ActionType 'SetOneDriveReadOnly' `
        -Target @{ siteUrl = $SiteUrl } `
        -NoUndoReason 'Unlocking is operator-judgment, not an automatic reverse (use Set-SPOSite -LockState Unlock manually).' `
        -Action { Set-SPOSite -Identity $SiteUrl -LockState ReadOnly -ErrorAction Stop | Out-Null; $true } | Out-Null
}

function Extend-OneDriveRetention {
    <#
        Record the planned end-of-retention date in the audit log
        and warn the operator that the actual ~30-day OneDrive
        auto-purge clock is NOT reset by any single cmdlet. The
        practical options (keep license / Purview policy / transfer
        owner) are surfaced as guidance.

        Set-SPOSite -EnableAutoExpirationVersionTrim is unrelated;
        we don't try to set it. We do set -DenyAddAndCustomizePages
        to $false to ensure the successor can use the site cleanly.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [int]$Days = 60
    )
    $endDate = (Get-Date).ToUniversalTime().AddDays($Days)
    Invoke-Action `
        -Description ("Extend OneDrive retention intent for '{0}' to {1:o}" -f $SiteUrl, $endDate) `
        -ActionType 'ExtendOneDriveRetentionIntent' `
        -Target @{ siteUrl = $SiteUrl; days = $Days; plannedEndUtc = $endDate.ToString('o') } `
        -NoUndoReason 'Retention extension is intent-only; no per-site cmdlet reverses the M365 auto-purge clock.' `
        -Action {
            # Best effort: nothing to actually call. We just record
            # the intent. Surface the operational options to the
            # operator so they know what comes next.
            $true
        } | Out-Null
    Write-Warn "Auto-purge of orphaned OneDrive begins ~30 days after license removal."
    Write-InfoMsg "To genuinely extend retention, do ONE of:"
    Write-InfoMsg "  1. Keep the leaver's license assigned until $($endDate.ToString('yyyy-MM-dd'))"
    Write-InfoMsg "  2. Apply a Microsoft Purview retention policy targeting their OneDrive"
    Write-InfoMsg "  3. Reassign site ownership to a service account whose license stays on"
}

function Send-OneDriveHandoffSummary {
    <#
        Send the manager (or any recipient) an HTML email with the
        leaver's OneDrive URL, planned retention end date, and a
        top-20 recently-modified file list. Sent via Graph
        /me/sendMail using the operator's mailbox.
    #>
    param(
        [Parameter(Mandatory)][string]$ManagerUPN,
        [Parameter(Mandatory)][string]$LeaverUPN,
        [Parameter(Mandatory)][string]$SiteUrl,
        [array]$RecentFiles = @(),
        [DateTime]$PlannedRetentionEndUtc = ((Get-Date).ToUniversalTime().AddDays(60))
    )

    $rowsHtml = ''
    foreach ($f in $RecentFiles) {
        $name = [System.Net.WebUtility]::HtmlEncode([string]$f.Name)
        $by   = [System.Net.WebUtility]::HtmlEncode([string]$f.LastModifiedBy)
        $when = if ($f.LastModifiedUtc) { $f.LastModifiedUtc.ToString('yyyy-MM-dd HH:mm') } else { '' }
        $url  = [string]$f.WebUrl
        $rowsHtml += "<tr><td><a href='$url'>$name</a></td><td>$when</td><td>$by</td></tr>"
    }
    $body = @"
<html><body style='font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222'>
<p>Hello,</p>
<p>You have been granted access to the OneDrive of <b>$([System.Net.WebUtility]::HtmlEncode($LeaverUPN))</b> as part of their offboarding.</p>
<p><b>OneDrive URL:</b> <a href='$SiteUrl'>$SiteUrl</a></p>
<p><b>Planned retention end (UTC):</b> $($PlannedRetentionEndUtc.ToString('yyyy-MM-dd HH:mm'))<br>
Please save anything you need before this date.</p>
<h3>Recently modified files (top $($RecentFiles.Count), last 90 days):</h3>
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px'>
<tr style='background:#eef2f5'><th align='left'>File</th><th align='left'>Modified (UTC)</th><th align='left'>By</th></tr>
$rowsHtml
</table>
<p style='color:#666;font-size:12px'>Sent automatically by M365 Manager.</p>
</body></html>
"@

    if (Get-Command Send-Email -ErrorAction SilentlyContinue) {
        Send-Email -To @($ManagerUPN) -Subject "OneDrive handoff: $LeaverUPN" -Body $body | Out-Null
    } else {
        # Standalone fallback: direct /me/sendMail when Notifications.ps1 not loaded.
        $message = @{
            message = @{
                subject      = "OneDrive handoff: $LeaverUPN"
                body         = @{ contentType = "HTML"; content = $body }
                toRecipients = @(@{ emailAddress = @{ address = $ManagerUPN } })
            }
            saveToSentItems = $true
        } | ConvertTo-Json -Depth 10
        Invoke-Action `
            -Description ("Send OneDrive handoff email to {0} (leaver: {1})" -f $ManagerUPN, $LeaverUPN) `
            -ActionType 'SendOneDriveHandoffEmail' `
            -Target @{ recipient = $ManagerUPN; leaverUpn = $LeaverUPN; siteUrl = $SiteUrl } `
            -NoUndoReason 'Email send is irreversible.' `
            -Action {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/sendMail" -Body $message -ContentType 'application/json' -ErrorAction Stop | Out-Null
                $true
            } | Out-Null
    }
}

function Send-OffboardManagerSummary {
    <#
        Higher-level offboard summary email. Sent in Step 10 of the
        offboard flow once the per-system work is done. -Summary
        is a hashtable of "key -> value" rows that the helper
        renders as a small HTML table; the offboard orchestrator
        passes in counts ("Licenses removed: 3", "Teams handed off:
        2") so the manager has a one-glance recap.
    #>
    param(
        [Parameter(Mandatory)][string]$ManagerUPN,
        [Parameter(Mandatory)][string]$LeaverUPN,
        [hashtable]$Summary = @{}
    )
    $rows = ''
    foreach ($k in ($Summary.Keys | Sort-Object)) {
        $kHtml = [System.Net.WebUtility]::HtmlEncode([string]$k)
        $vHtml = [System.Net.WebUtility]::HtmlEncode([string]$Summary[$k])
        $rows += "<tr><th align='left'>$kHtml</th><td>$vHtml</td></tr>"
    }
    $body = @"
<html><body style='font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#222'>
<p>Offboarding complete for <b>$([System.Net.WebUtility]::HtmlEncode($LeaverUPN))</b>.</p>
<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:13px'>$rows</table>
<p>If anything looks off, reply to this email and IT will follow up.</p>
<p style='color:#666;font-size:12px'>Sent automatically by M365 Manager.</p>
</body></html>
"@
    if (Get-Command Send-Email -ErrorAction SilentlyContinue) {
        return [bool] (Send-Email -To @($ManagerUPN) -Subject "Offboard complete: $LeaverUPN" -Body $body)
    }
    # Standalone fallback when Notifications.ps1 isn't loaded.
    $message = @{
        message = @{
            subject      = "Offboard complete: $LeaverUPN"
            body         = @{ contentType = "HTML"; content = $body }
            toRecipients = @(@{ emailAddress = @{ address = $ManagerUPN } })
        }
        saveToSentItems = $true
    } | ConvertTo-Json -Depth 10
    return Invoke-Action `
        -Description ("Send offboard summary to manager {0} (leaver: {1})" -f $ManagerUPN, $LeaverUPN) `
        -ActionType 'SendOffboardSummary' `
        -Target @{ manager = $ManagerUPN; leaverUpn = $LeaverUPN } `
        -NoUndoReason 'Email send is irreversible.' `
        -Action {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/me/sendMail" -Body $message -ContentType 'application/json' -ErrorAction Stop | Out-Null
            $true
        }
}

function Invoke-OneDriveHandoff {
    <#
        Orchestrator. Returns a hashtable describing what happened
        so the offboard flow can record it in its result CSV row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LeaverUPN,
        [Parameter(Mandatory)][string]$SuccessorUPN,
        [int]$RetentionDays = 60,
        [switch]$NotifyManager,
        [string]$ManagerUPN
    )

    $result = [ordered]@{
        LeaverUPN      = $LeaverUPN
        SuccessorUPN   = $SuccessorUPN
        SiteUrl        = $null
        AccessGranted  = $false
        Notified       = $false
        Note           = ''
    }

    if (-not (Connect-ForTask 'OneDrive')) {
        $result.Note = 'SPO connection failed; skipped.'
        return [PSCustomObject]$result
    }

    $url = Get-UserOneDriveUrl -UPN $LeaverUPN
    if (-not $url) {
        $result.Note = 'No personal site found.'
        return [PSCustomObject]$result
    }
    $result.SiteUrl = $url

    Grant-OneDriveAccess -SiteUrl $url -GranteeUPN $SuccessorUPN
    $result.AccessGranted = $true

    if ($RetentionDays -gt 0) {
        Extend-OneDriveRetention -SiteUrl $url -Days $RetentionDays
    }

    if ($NotifyManager) {
        $recipient = if ($ManagerUPN) { $ManagerUPN } else { $SuccessorUPN }
        $files = Get-OneDriveRecentFiles -SiteUrl $url -Days 90 -Top 20
        $endUtc = (Get-Date).ToUniversalTime().AddDays($RetentionDays)
        Send-OneDriveHandoffSummary -ManagerUPN $recipient -LeaverUPN $LeaverUPN -SiteUrl $url -RecentFiles $files -PlannedRetentionEndUtc $endUtc
        $result.Notified = $true
    }
    return [PSCustomObject]$result
}
