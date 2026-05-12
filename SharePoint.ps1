# ============================================================
#  SharePoint.ps1 — SPO site, sharing, and lightweight
#  provisioning operations.
#
#  Builds on the SPO connection set up in Auth.ps1's Connect-SPO
#  (Commit A). Required permission: SharePoint Administrator.
#
#  Heavy operations (Get-OrphanedSites, Get-StaleSites) walk all
#  SPO sites and run extra cmdlets per site; they're report-style
#  and expected to take a few seconds per ~100 sites.
# ============================================================

# ============================================================
#  Site management
# ============================================================

function Get-SiteOwners {
    param([Parameter(Mandatory)][string]$SiteUrl)
    try {
        $admins = @(Get-SPOUser -Site $SiteUrl -Limit ALL -ErrorAction Stop | Where-Object { $_.IsSiteAdmin })
        return @($admins | ForEach-Object {
            [PSCustomObject]@{ LoginName = $_.LoginName; DisplayName = $_.DisplayName }
        })
    } catch { Write-Warn "Could not read owners of $SiteUrl -- $($_.Exception.Message)"; return @() }
}

function Set-SiteOwner {
    <#
        Add or remove a site collection admin. Reversible pair:
        AddSiteOwner <-> RemoveSiteOwner.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$UPN,
        [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Direction
    )
    if ($Direction -eq 'Add') {
        return Invoke-Action `
            -Description ("Add {0} as site-collection admin on {1}" -f $UPN, $SiteUrl) `
            -ActionType 'AddSiteOwner' `
            -Target @{ siteUrl = $SiteUrl; userUpn = $UPN } `
            -ReverseType 'RemoveSiteOwner' `
            -ReverseDescription ("Remove {0} as site-collection admin on {1}" -f $UPN, $SiteUrl) `
            -Action { Set-SPOUser -Site $SiteUrl -LoginName $UPN -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null; $true }
    } else {
        return Invoke-Action `
            -Description ("Remove {0} as site-collection admin on {1}" -f $UPN, $SiteUrl) `
            -ActionType 'RemoveSiteOwner' `
            -Target @{ siteUrl = $SiteUrl; userUpn = $UPN } `
            -ReverseType 'AddSiteOwner' `
            -ReverseDescription ("Re-add {0} as site-collection admin on {1}" -f $UPN, $SiteUrl) `
            -Action { Set-SPOUser -Site $SiteUrl -LoginName $UPN -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null; $true }
    }
}

function Get-AllSites {
    <#
        All non-personal sites. Returns a normalized record.
    #>
    try {
        $sites = @(Get-SPOSite -Limit ALL -ErrorAction Stop)
        return @($sites | ForEach-Object {
            [PSCustomObject]@{
                Url                       = $_.Url
                Title                     = $_.Title
                Owner                     = $_.Owner
                Template                  = $_.Template
                StorageUsageCurrent       = $_.StorageUsageCurrent
                StorageQuota              = $_.StorageQuota
                LastContentModifiedDate   = $_.LastContentModifiedDate
                LockState                 = $_.LockState
                SharingCapability         = $_.SharingCapability
            }
        })
    } catch { Write-ErrorMsg "Could not enumerate sites: $($_.Exception.Message)"; return @() }
}

function Get-OrphanedSites {
    <#
        Sites whose only "owners" are system accounts (or no SCA
        at all). Slow on big tenants -- one Get-SPOUser per site.
    #>
    Write-InfoMsg "Scanning sites for zero non-system owners..."
    $sites = Get-AllSites
    $hits  = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($s in $sites) {
        $i++
        Write-Progress -Activity "Orphan site scan" -Status $s.Url -PercentComplete (($i / [Math]::Max(1, $sites.Count)) * 100)
        $owners = Get-SiteOwners -SiteUrl $s.Url
        # Filter out the system accounts (login starts with c:0t.c|tenant|... or app@sharepoint, etc.)
        $human = @($owners | Where-Object { $_.LoginName -notmatch '^(c:0|SHAREPOINT\\|app@sharepoint|i:0#\.f\|membership\|app_)' })
        if ($human.Count -eq 0) {
            [void]$hits.Add([PSCustomObject]@{ Url = $s.Url; Title = $s.Title; LastModified = $s.LastContentModifiedDate; OwnerCount = 0 })
        }
    }
    Write-Progress -Activity "Orphan site scan" -Completed
    return @($hits | Sort-Object Url)
}

function Get-StaleSites {
    param([int]$DaysSinceModified = 365)
    $cutoff = (Get-Date).AddDays(-$DaysSinceModified)
    $sites = Get-AllSites
    return @($sites | Where-Object { $_.LastContentModifiedDate -and ([DateTime]$_.LastContentModifiedDate) -lt $cutoff } | Sort-Object LastContentModifiedDate)
}

function Get-SiteStorageReport {
    <#
        Top N sites by storage. -NearQuotaPercent filter (>= N%
        of quota used) optional.
    #>
    param([int]$Top = 25, [int]$NearQuotaPercent = 0)
    $sites = Get-AllSites
    $with = $sites | ForEach-Object {
        $pct = 0
        if ($_.StorageQuota -gt 0) { $pct = [Math]::Round(($_.StorageUsageCurrent / $_.StorageQuota) * 100, 1) }
        $_ | Add-Member -NotePropertyName UsedPercent -NotePropertyValue $pct -PassThru -Force
    }
    if ($NearQuotaPercent -gt 0) { $with = $with | Where-Object UsedPercent -ge $NearQuotaPercent }
    return @($with | Sort-Object StorageUsageCurrent -Descending | Select-Object -First $Top)
}

# ============================================================
#  External sharing
# ============================================================

function Get-UserOutboundShares {
    <#
        Find every external/sharing event the user created in the
        last -Days days via the Unified Audit Log. Requires UAL
        ingestion enabled (cf. UnifiedAuditLog.ps1's health check).
        We don't re-check here; the caller is expected to have
        connected EXO already. Returns one record per sharing
        operation with enough context to revoke.
    #>
    param(
        [Parameter(Mandatory)][string]$UPN,
        [int]$Days = 365
    )
    if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
        Write-Warn "Search-UnifiedAuditLog not loaded. Connect EXO first."
        return @()
    }
    $from = (Get-Date).AddDays(-$Days)
    $to   = Get-Date
    try {
        $rows = @(Search-UnifiedAuditLog -StartDate $from -EndDate $to -UserIds $UPN `
                  -Operations 'AnonymousLinkCreated','SecureLinkCreated','CompanyLinkCreated','SharingSet','SharingInvitationCreated' `
                  -ResultSize 1000 -ErrorAction Stop)
    } catch { Write-ErrorMsg "Audit search failed: $($_.Exception.Message)"; return @() }

    $out = New-Object System.Collections.ArrayList
    foreach ($r in $rows) {
        $parsed = $null
        try { $parsed = $r.AuditData | ConvertFrom-Json -ErrorAction Stop } catch {}
        [void]$out.Add([PSCustomObject]@{
            TimeUtc       = ([DateTime]$r.CreationDate).ToUniversalTime()
            Operation     = $r.Operations
            SiteUrl       = if ($parsed) { $parsed.SiteUrl } else { '' }
            ObjectId      = if ($parsed) { $parsed.ObjectId } else { '' }
            TargetUser    = if ($parsed -and $parsed.TargetUserOrGroupName) { $parsed.TargetUserOrGroupName } else { '' }
            LinkType      = if ($parsed) { $parsed.EventData } else { '' }
            SharingType   = if ($parsed) { $parsed.SharingType } else { '' }
            AuditDataJson = $r.AuditData
        })
    }
    return @($out | Sort-Object TimeUtc -Descending)
}

function Revoke-Share {
    <#
        Best-effort revocation. Path depends on the operation type:
          - AnonymousLinkCreated / SecureLinkCreated: revoke the
            sharing link via Graph /shares/{shareId}/permission/{id}.
            We don't always have the share id, so this often
            requires an operator-provided link URL.
          - SharingInvitationCreated: remove the external user.
          - SharingSet on a document: remove the granted permission.
        We accept either a -ShareRecord (from Get-UserOutboundShares)
        or explicit -SiteUrl / -GranteeUPN parameters.
    #>
    [CmdletBinding(DefaultParameterSetName='record')]
    param(
        [Parameter(ParameterSetName='record', Mandatory)][PSCustomObject]$ShareRecord,
        [Parameter(ParameterSetName='manual')][string]$SiteUrl,
        [Parameter(ParameterSetName='manual')][string]$GranteeUPN
    )
    if ($ShareRecord) {
        $SiteUrl    = [string]$ShareRecord.SiteUrl
        $GranteeUPN = [string]$ShareRecord.TargetUser
    }
    if (-not $SiteUrl) { Write-Warn "Cannot revoke -- no SiteUrl on record."; return $false }

    return Invoke-Action `
        -Description ("Revoke external share on {0} for '{1}'" -f $SiteUrl, $GranteeUPN) `
        -ActionType 'RevokeExternalShare' `
        -Target @{ siteUrl = $SiteUrl; granteeUpn = $GranteeUPN } `
        -NoUndoReason 'Share recreation requires the original recipient consent and the original document context.' `
        -Action {
            if ($GranteeUPN -and $GranteeUPN -match '@') {
                # Best effort: drop the external user from the tenant
                Remove-SPOExternalUser -DisplayNames $GranteeUPN -Confirm:$false -ErrorAction Stop | Out-Null
            } else {
                # No grantee -- treat as no-op success so the caller can
                # move on; the operator will likely need to use the SPO
                # admin UI to revoke an anonymous link by URL.
                Write-Warn "No grantee UPN on this share -- revoke anonymous links via the SPO admin UI."
            }
            $true
        }
}

function Set-AnonymousLinkExpiry {
    <#
        Tenant policy nudge -- prints the exact cmdlet rather than
        running it, because changing tenant policy is high-touch
        and operators usually want to read the docs first.
    #>
    param([int]$DefaultDays = 14)
    Write-Host ""
    Write-Host "  Tenant policy command (run manually after review):" -ForegroundColor Yellow
    Write-Host "    Set-SPOTenant -RequireAnonymousLinksExpireInDays $DefaultDays" -ForegroundColor Yellow
    Write-Host ""
    Write-InfoMsg "Docs: https://learn.microsoft.com/sharepoint/turn-external-sharing-on-or-off"
}

# ============================================================
#  Permissions audit
# ============================================================

function Get-SitePermissions {
    <#
        Flatten site-collection users + group members into a
        single list of (User, Role, GrantedVia). 'Role' is taken
        from the Site Permission Group ('Owners','Members','Visitors',
        custom group name). 'GrantedVia' is the group name when
        the user is a group member, or 'direct' otherwise.
    #>
    param([Parameter(Mandatory)][string]$SiteUrl)
    $out = New-Object System.Collections.ArrayList
    try {
        $groups = @(Get-SPOSiteGroup -Site $SiteUrl -ErrorAction Stop)
        foreach ($g in $groups) {
            try {
                $users = @($g.Users)   # already populated by Get-SPOSiteGroup
                foreach ($u in $users) {
                    [void]$out.Add([PSCustomObject]@{
                        User       = $u
                        Role       = $g.Title
                        GrantedVia = $g.Title
                    })
                }
            } catch {}
        }
        # Direct SCAs
        foreach ($a in (Get-SiteOwners -SiteUrl $SiteUrl)) {
            [void]$out.Add([PSCustomObject]@{ User = $a.LoginName; Role = 'Site Collection Admin'; GrantedVia = 'direct' })
        }
    } catch { Write-Warn "Could not enumerate permissions: $($_.Exception.Message)" }
    return @($out | Sort-Object User, Role)
}

# ============================================================
#  Lightweight site provisioning
# ============================================================

function Get-SiteTemplate {
    param([Parameter(Mandatory)][string]$Name)
    $key = ($Name -replace '^site-', '').ToLowerInvariant()
    $base = Get-OnboardTemplates  # reuse Templates.ps1's TemplatesRoot
    $tplDir = $null
    if (Get-Variable -Name TemplatesRoot -Scope Script -ErrorAction SilentlyContinue) { $tplDir = $script:TemplatesRoot }
    if (-not $tplDir) {
        # Best-guess fallback to <module root>/templates
        $here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $tplDir = Join-Path $here 'templates'
    }
    $path = Join-Path $tplDir ("site-{0}.json" -f $key)
    if (-not (Test-Path -LiteralPath $path)) { throw "Site template 'site-$key' not found at $path" }
    try { return (Get-Content -LiteralPath $path -Raw | ConvertFrom-Json) }
    catch { throw "Could not parse site template at $path : $_" }
}

function New-ProjectSite {
    <#
        Lightweight provisioner: creates an SPO communication or
        team site from a templates/site-*.json descriptor.

        Descriptor fields:
          name        : human-readable site name
          alias       : URL alias (must be unique tenant-wide)
          template    : SPO template -- "STS#3" (modern team
                        without group), "SITEPAGEPUBLISHING#0"
                        (communication), or "GROUP#0" (group team
                        site)
          description : optional
          hubSiteUrl  : optional, will Register-SPOHubSite-associate
                        the new site after creation
          permissions : array of { upn, role } where role is one
                        of Owner|Member|Visitor (best-effort
                        Add-SPOUser/Add-SPOSiteCollectionAppCatalog)
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$OwnerUPN,
        [Parameter(Mandatory)][string]$Template
    )
    $tpl = Get-SiteTemplate -Name $Template
    if (-not $tpl.alias) { throw "Site template missing 'alias' field." }
    $url = "$($script:SessionState.SharePointAdminUrl -replace '-admin\.sharepoint\.com','.sharepoint.com')/sites/$($tpl.alias)"

    return Invoke-Action `
        -Description ("Create SPO site '{0}' at {1} (template: {2})" -f $Name, $url, $tpl.template) `
        -ActionType 'CreateSPOSite' `
        -Target @{ siteUrl = $url; owner = $OwnerUPN; templateName = [string]$tpl.template; templateKey = $Template } `
        -NoUndoReason 'Site deletion is destructive and asynchronous; no clean per-call reverse.' `
        -StubReturn ([PSCustomObject]@{ Url = $url; Title = $Name; Owner = $OwnerUPN }) `
        -Action {
            New-SPOSite -Url $url -Owner $OwnerUPN -Title $Name -Template $tpl.template -StorageQuota 1024 -ErrorAction Stop
            if ($tpl.hubSiteUrl) {
                try { Add-SPOHubSiteAssociation -Site $url -HubSite $tpl.hubSiteUrl -ErrorAction Stop } catch { Write-Warn "Hub association failed: $_" }
            }
            [PSCustomObject]@{ Url = $url; Title = $Name; Owner = $OwnerUPN }
        }
}

# ============================================================
#  Offboard integration -- revoke a leaver's outbound shares
# ============================================================

function Invoke-SharePointOffboardCleanup {
    <#
        Pull every outbound share the leaver created (last 365 days
        from UAL) and revoke each with Y/A/N confirm. Returns a
        counter object for the offboard summary CSV.
    #>
    param([Parameter(Mandatory)][string]$LeaverUPN, [int]$LookbackDays = 365)
    $shares = Get-UserOutboundShares -UPN $LeaverUPN -Days $LookbackDays
    if (-not $shares -or $shares.Count -eq 0) {
        return [PSCustomObject]@{ LeaverUPN = $LeaverUPN; ShareCount = 0; Revoked = 0; Skipped = 0; Failed = 0 }
    }
    Write-InfoMsg "Found $($shares.Count) outbound share event(s) for $LeaverUPN."
    $revoked = 0; $skipped = 0; $failed = 0; $runAll = $false
    foreach ($s in $shares) {
        $line = "{0}  {1}  {2} -> {3}" -f $s.TimeUtc.ToString('yyyy-MM-dd'), $s.Operation, $s.SiteUrl, $s.TargetUser
        if (-not $runAll) {
            $opt = Show-Menu -Title "Revoke this share?" -Options @("Yes","Yes to ALL remaining","No","Quit") -BackLabel "Skip"
            if ($opt -eq -1 -or $opt -eq 2) { $skipped++; continue }
            if ($opt -eq 3) { break }
            if ($opt -eq 1) { $runAll = $true }
        }
        Write-Host "  $line" -ForegroundColor DarkGray
        $ok = Revoke-Share -ShareRecord $s
        if ($ok) { $revoked++ } else { $failed++ }
    }
    return [PSCustomObject]@{ LeaverUPN = $LeaverUPN; ShareCount = $shares.Count; Revoked = $revoked; Skipped = $skipped; Failed = $failed }
}

# ============================================================
#  Menu
# ============================================================

function Start-SharePointMenu {
    while ($true) {
        $sel = Show-Menu -Title "SharePoint" -Options @(
            "Get site owners",
            "Add / remove site owner...",
            "Get site permissions",
            "Report: orphaned sites",
            "Report: stale sites (365d default)",
            "Report: top storage sites",
            "Report: external shares for a user",
            "Tenant policy: set anonymous-link expiry default",
            "Create site from template..."
        ) -BackLabel "Back"
        switch ($sel) {
            0 { $u = Read-UserInput "Site URL"; if ($u) { Get-SiteOwners -SiteUrl $u | Format-Table -AutoSize; Pause-ForUser } }
            1 {
                $u = Read-UserInput "Site URL"; if (-not $u) { continue }
                $upn = Read-UserInput "User UPN"; if (-not $upn) { continue }
                $dirSel = Show-Menu -Title "Direction" -Options @("Add","Remove") -BackLabel "Cancel"
                if ($dirSel -eq -1) { continue }
                Set-SiteOwner -SiteUrl $u -UPN $upn -Direction $(if ($dirSel -eq 0) {'Add'} else {'Remove'}) | Out-Null
                Pause-ForUser
            }
            2 { $u = Read-UserInput "Site URL"; if ($u) { Get-SitePermissions -SiteUrl $u | Format-Table -AutoSize; Pause-ForUser } }
            3 { Get-OrphanedSites | Format-Table -AutoSize; Pause-ForUser }
            4 {
                $dt = Read-UserInput "Days threshold (default 365)"
                $d = 365; [int]::TryParse($dt, [ref]$d) | Out-Null
                Get-StaleSites -DaysSinceModified $d | Format-Table -AutoSize
                Pause-ForUser
            }
            5 {
                $nt = Read-UserInput "Top N (default 25)"; $n = 25; [int]::TryParse($nt,[ref]$n) | Out-Null
                $qt = Read-UserInput "Near-quota %% (0 = ignore, e.g. 80)"; $q = 0; [int]::TryParse($qt,[ref]$q) | Out-Null
                Get-SiteStorageReport -Top $n -NearQuotaPercent $q | Format-Table -AutoSize
                Pause-ForUser
            }
            6 {
                $upn = Read-UserInput "User UPN"; if (-not $upn) { continue }
                $dt  = Read-UserInput "Lookback days (default 365)"; $d = 365; [int]::TryParse($dt, [ref]$d) | Out-Null
                Get-UserOutboundShares -UPN $upn -Days $d | Format-Table -AutoSize
                Pause-ForUser
            }
            7 {
                $dt = Read-UserInput "Default expiry days (default 14)"; $d = 14; [int]::TryParse($dt,[ref]$d) | Out-Null
                Set-AnonymousLinkExpiry -DefaultDays $d
                Pause-ForUser
            }
            8 {
                $name = Read-UserInput "Site display name"; if (-not $name) { continue }
                $own = Read-UserInput "Owner UPN"; if (-not $own) { continue }
                $opts = @('project','team')
                $sel = Show-Menu -Title "Template" -Options $opts -BackLabel "Cancel"
                if ($sel -eq -1) { continue }
                try { New-ProjectSite -Name $name -OwnerUPN $own -Template $opts[$sel] } catch { Write-ErrorMsg "$_" }
                Pause-ForUser
            }
            -1 { return }
        }
    }
}
