# ============================================================
#  TemplateGenerator.ps1 - Build onboarding role templates by
#  scraping the connected tenant.
#
#  Workflow:
#    1. Enumerate active member users from Microsoft Graph.
#    2. Cluster them by (OfficeLocation, JobTitle).
#    3. For each operator-selected cluster, sample a few users,
#       intersect their license SKUs / group memberships / DL
#       memberships, and write the result as templates/role-*.json.
#    4. The standard Start-Onboard "Apply a role-based template?"
#       prompt picks the new files up immediately on next run.
#
#  Required Graph scopes (already in $script:MgScopes):
#    User.Read.All, Directory.Read.All, Organization.Read.All
# ============================================================

function ConvertTo-TemplateSlug {
    <#
        Normalize a "Location | JobTitle" key into a filesystem-safe
        slug used for the role-*.json filename. Idempotent.
    #>
    param([Parameter(Mandatory)][string]$Text)
    $s = $Text.ToLowerInvariant()
    $s = $s -replace '[^a-z0-9]+', '-'
    $s = $s.Trim('-')
    if (-not $s) { $s = 'role' }
    return $s
}

function Get-TenantUserProfile {
    <#
        Pull every active member user with the properties we need to
        cluster. Returns an array of PSCustomObjects. Uses -All so
        large tenants are paged for us; -Top caps the absolute pull
        size as a safety brake for >50k tenants.
    #>
    param([int]$Limit = 10000)
    if (-not (Get-Command Get-MgUser -ErrorAction SilentlyContinue)) {
        Write-Warn "Microsoft.Graph.Users not loaded."
        return @()
    }
    if (Get-Command Test-GraphConnected -ErrorAction SilentlyContinue) {
        if (-not (Test-GraphConnected)) {
            Write-Warn "Microsoft Graph is not authenticated."
            return @()
        }
    }
    $props = "Id,DisplayName,UserPrincipalName,JobTitle,Department,OfficeLocation,UsageLocation,City,State,Country,AccountEnabled,UserType"
    Write-InfoMsg "Enumerating tenant users (this may take a minute on larger tenants)..."
    try {
        $all = @(Get-MgUser -All -Property $props -PageSize 999 -ErrorAction Stop |
            Where-Object { $_.UserType -eq 'Member' -and $_.AccountEnabled } |
            Select-Object -First $Limit)
        Write-Success "Pulled $($all.Count) active member user(s)."
        return $all
    } catch {
        $msg = if (Get-Command Resolve-GraphError -ErrorAction SilentlyContinue) { Resolve-GraphError -ErrorRecord $_ } else { "$_" }
        Write-ErrorMsg "User enumeration failed: $msg"
        return @()
    }
}

function Group-UsersByLocationAndJob {
    <#
        Bucket users by a (Location, JobTitle) pair. Skips users with
        no JobTitle (templates without a job title are useless for
        onboarding). Location falls back: OfficeLocation -> "City,
        State" -> City -> "" (job-only cluster).
        Returns an array of cluster PSCustomObjects sorted by Users
        count desc.
    #>
    param([Parameter(Mandatory)][array]$Users)
    $clusters = @{}
    foreach ($u in $Users) {
        $job = if ($u.JobTitle) { $u.JobTitle.Trim() } else { '' }
        if (-not $job) { continue }
        $loc = ''
        if ($u.OfficeLocation) { $loc = $u.OfficeLocation.Trim() }
        elseif ($u.City -and $u.State) { $loc = "$($u.City), $($u.State)".Trim() }
        elseif ($u.City) { $loc = $u.City.Trim() }
        $key = if ($loc) { "$loc | $job" } else { $job }

        if (-not $clusters.ContainsKey($key)) {
            $clusters[$key] = [PSCustomObject]@{
                Key            = $key
                Location       = $loc
                JobTitle       = $job
                Department     = $u.Department
                UsageLocation  = $u.UsageLocation
                Users          = New-Object System.Collections.ArrayList
            }
        }
        [void]$clusters[$key].Users.Add($u)
        # Use first non-null department / usage location we see for the cluster.
        if (-not $clusters[$key].Department -and $u.Department)       { $clusters[$key].Department = $u.Department }
        if (-not $clusters[$key].UsageLocation -and $u.UsageLocation) { $clusters[$key].UsageLocation = $u.UsageLocation }
    }
    return @($clusters.Values | Sort-Object { -$_.Users.Count })
}

function Get-CommonAssignments {
    <#
        Sample up to -MaxSamples random users from the cluster, look
        up each one's licenses (Get-MgUserLicenseDetail) and group
        memberships (Get-MgUserMemberOf), and return the items that
        appear in >= -Threshold fraction of the samples.

        Group memberships are split into security groups vs
        distribution lists using mailEnabled / securityEnabled
        heuristics (mailEnabled+not-securityEnabled = DL).
    #>
    param(
        [Parameter(Mandatory)]$Cluster,
        [int]$MaxSamples = 5,
        [double]$Threshold = 0.5
    )
    $userCount = $Cluster.Users.Count
    $n = [Math]::Min($MaxSamples, $userCount)
    $samples = @($Cluster.Users | Get-Random -Count $n)
    Write-InfoMsg ("  Sampling {0} of {1} users in cluster '{2}'..." -f $n, $userCount, $Cluster.Key)

    $skuCounts = @{}
    $sgCounts  = @{}
    $dlCounts  = @{}

    foreach ($u in $samples) {
        try {
            $lics = @(Get-MgUserLicenseDetail -UserId $u.Id -ErrorAction Stop)
            foreach ($l in $lics) {
                if ($l.SkuPartNumber) {
                    if (-not $skuCounts.ContainsKey($l.SkuPartNumber)) { $skuCounts[$l.SkuPartNumber] = 0 }
                    $skuCounts[$l.SkuPartNumber]++
                }
            }
        } catch {
            Write-Warn ("    license lookup failed for {0}: {1}" -f $u.UserPrincipalName, $_.Exception.Message)
        }
        try {
            $groups = @(Get-MgUserMemberOf -UserId $u.Id -All -ErrorAction Stop)
            foreach ($g in $groups) {
                $ap = $g.AdditionalProperties
                $name = [string]$ap['displayName']
                if (-not $name) { continue }
                $odataType = [string]$ap['@odata.type']
                if ($odataType -notmatch 'group$') { continue }   # skip roles / orgs
                $mail = [bool]$ap['mailEnabled']
                $sec  = [bool]$ap['securityEnabled']
                if ($mail -and -not $sec) {
                    if (-not $dlCounts.ContainsKey($name)) { $dlCounts[$name] = 0 }
                    $dlCounts[$name]++
                } else {
                    if (-not $sgCounts.ContainsKey($name)) { $sgCounts[$name] = 0 }
                    $sgCounts[$name]++
                }
            }
        } catch {
            Write-Warn ("    group lookup failed for {0}: {1}" -f $u.UserPrincipalName, $_.Exception.Message)
        }
    }

    $cut = [Math]::Max(1, [Math]::Ceiling($n * $Threshold))
    $skus = @($skuCounts.GetEnumerator() | Where-Object { $_.Value -ge $cut } | Sort-Object Value -Descending | ForEach-Object { $_.Key })
    $sgs  = @($sgCounts.GetEnumerator()  | Where-Object { $_.Value -ge $cut } | Sort-Object Value -Descending | ForEach-Object { $_.Key })
    $dls  = @($dlCounts.GetEnumerator()  | Where-Object { $_.Value -ge $cut } | Sort-Object Value -Descending | ForEach-Object { $_.Key })

    return [PSCustomObject]@{
        SampleCount       = $n
        LicenseSKUs       = $skus
        SecurityGroups    = $sgs
        DistributionLists = $dls
    }
}

function New-TemplateFromCluster {
    <#
        Build a template hashtable in the schema Onboard.ps1 expects.
        Calls Get-CommonAssignments for license / group / DL discovery.
    #>
    param(
        [Parameter(Mandatory)]$Cluster,
        [int]$MaxSamples = 5,
        [double]$Threshold = 0.5
    )
    $assign = Get-CommonAssignments -Cluster $Cluster -MaxSamples $MaxSamples -Threshold $Threshold

    $usageLoc = if ($Cluster.UsageLocation) { $Cluster.UsageLocation } else { 'US' }
    $desc = if ($Cluster.Location) {
        "Auto-generated from {0} sampled user(s): {1} at {2}" -f $assign.SampleCount, $Cluster.JobTitle, $Cluster.Location
    } else {
        "Auto-generated from {0} sampled user(s): {1}" -f $assign.SampleCount, $Cluster.JobTitle
    }

    [ordered]@{
        _comment              = "Auto-generated by Template Generator on $(Get-Date -Format 'yyyy-MM-dd HH:mm') from cluster '$($Cluster.Key)'. Review/edit before relying on it."
        name                  = $Cluster.Key
        description           = $desc
        usageLocation         = $usageLoc
        licenseSKUs           = $assign.LicenseSKUs
        securityGroups        = $assign.SecurityGroups
        distributionLists     = $assign.DistributionLists
        sharedMailboxes       = @()
        teams                 = @()
        oneDrive              = $null
        defaults              = @{
            Department     = $Cluster.Department
            OfficeLocation = $Cluster.Location
        }
        contractorExpiryDays  = $null
    }
}

function Save-GeneratedTemplate {
    <#
        Write a template object to templates/[<tenant-slug>/]role-<slug>.json.
        When a tenant context is active (Get-TenantTemplatesRoot returns
        non-null) the file goes into that per-tenant folder so MSPs
        running multiple Direct tenants don't pollute each other's
        templates. Falls back to the global templates/ root otherwise.

        If the file exists, append "-N" until we find a free name
        (we never silently overwrite operator-curated templates).
        Returns the path written, or $null on failure.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Template,
        [Parameter(Mandatory)][string]$Slug,
        [switch]$Overwrite
    )
    $dir = $null
    if (Get-Command Get-TenantTemplatesRoot -ErrorAction SilentlyContinue) {
        $dir = Get-TenantTemplatesRoot
    }
    if (-not $dir) {
        if (Get-Variable -Name TemplatesRoot -Scope Script -ErrorAction SilentlyContinue) {
            $dir = $script:TemplatesRoot
        }
    }
    if (-not $dir) {
        $here = if ($Global:M365RepoRoot) { $Global:M365RepoRoot } `
                elseif ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot } `
                else { (Get-Location).Path }
        $dir = Join-Path $here 'templates'
    }
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        catch { Write-ErrorMsg "Could not create templates directory: $_"; return $null }
    }

    $path = Join-Path $dir ("role-{0}.json" -f $Slug)
    if ((Test-Path -LiteralPath $path) -and -not $Overwrite) {
        # Find a free suffix
        $i = 2
        while (Test-Path -LiteralPath (Join-Path $dir ("role-{0}-{1}.json" -f $Slug, $i))) { $i++ }
        $path = Join-Path $dir ("role-{0}-{1}.json" -f $Slug, $i)
    }
    try {
        ($Template | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $path -Encoding UTF8 -Force
        return $path
    } catch {
        Write-ErrorMsg "Failed to write $path : $_"
        return $null
    }
}

function Start-TemplateGeneratorMenu {
    Write-SectionHeader "Onboarding Template Generator"

    if (-not (Connect-ForTask 'Onboard')) {
        Write-Warn "Graph not connected; cannot generate."
        Pause-ForUser; return
    }

    $sel = Show-Menu -Title "Template Generator" -Options @(
        "Scan tenant and generate role templates",
        "List existing templates on disk"
    ) -BackLabel "Back"
    if ($sel -eq -1) { return }

    if ($sel -eq 1) {
        if (Get-Command Get-OnboardTemplates -ErrorAction SilentlyContinue) {
            $existing = @(Get-OnboardTemplates)
            if ($existing.Count -eq 0) { Write-Warn "No templates found." }
            else { $existing | Select-Object Key,Name,Description,Path | Format-Table -AutoSize }
        } else { Write-Warn "Templates module not loaded." }
        Pause-ForUser
        return
    }

    # ---- Scan + cluster ----
    $cap = Read-UserInput "Cap on users to enumerate (default 5000)"
    $limit = 5000; [int]::TryParse($cap, [ref]$limit) | Out-Null
    $users = Get-TenantUserProfile -Limit $limit
    if ($users.Count -eq 0) { Pause-ForUser; return }

    $clusters = Group-UsersByLocationAndJob -Users $users
    if ($clusters.Count -eq 0) {
        Write-Warn "No (location, job title) clusters found. Users may be missing JobTitle on their profiles."
        Pause-ForUser; return
    }

    # ---- Show top clusters ----
    $minSize = Read-UserInput "Skip clusters smaller than N users (default 2)"
    $mn = 2; [int]::TryParse($minSize, [ref]$mn) | Out-Null
    $eligible = @($clusters | Where-Object { $_.Users.Count -ge $mn })
    if ($eligible.Count -eq 0) {
        Write-Warn "No clusters meet the minimum size of $mn. Lower the threshold or onboard more representative users."
        Pause-ForUser; return
    }

    Write-Host ""
    Write-Host ("  Found {0} cluster(s) with >= {1} user(s):" -f $eligible.Count, $mn) -ForegroundColor $script:Colors.Info
    $eligible | Select-Object @{N='Users';E={$_.Users.Count}}, JobTitle, Location, Department, UsageLocation | Format-Table -AutoSize

    # ---- Pick which clusters to template ----
    $labels = $eligible | ForEach-Object { "{0,4} users - {1}" -f $_.Users.Count, $_.Key }
    $idx = Show-MultiSelect -Title "Select clusters to generate templates for" -Options $labels -Prompt "Enter cluster #s (e.g. 1,3,5), or just '1' for one"
    if (-not $idx -or $idx.Count -eq 0) { Pause-ForUser; return }

    $sCap = Read-UserInput "Max users to sample per cluster (default 5)"
    $sN = 5; [int]::TryParse($sCap, [ref]$sN) | Out-Null
    $thrIn = Read-UserInput "Common-assignment threshold percent (default 50)"
    $thrPct = 50; [int]::TryParse($thrIn, [ref]$thrPct) | Out-Null
    if ($thrPct -lt 1 -or $thrPct -gt 100) { $thrPct = 50 }
    $threshold = $thrPct / 100.0

    $overwrite = (Read-UserInput "Overwrite existing role-<slug>.json if present? (y/N)") -match '^[Yy]'

    $written = New-Object System.Collections.ArrayList
    foreach ($i in $idx) {
        $c = $eligible[$i]
        Write-Host ""
        Write-InfoMsg ("Building template for cluster '{0}'..." -f $c.Key)
        $tpl = New-TemplateFromCluster -Cluster $c -MaxSamples $sN -Threshold $threshold
        $slug = ConvertTo-TemplateSlug -Text $c.Key
        $path = Save-GeneratedTemplate -Template $tpl -Slug $slug -Overwrite:$overwrite
        if ($path) {
            Write-Success ("  Wrote {0}" -f $path)
            [void]$written.Add([PSCustomObject]@{ Cluster = $c.Key; Path = $path; Licenses = $tpl.licenseSKUs.Count; SGs = $tpl.securityGroups.Count; DLs = $tpl.distributionLists.Count })
        }
    }

    Write-Host ""
    if ($written.Count -gt 0) {
        Write-Success ("Generated {0} template(s):" -f $written.Count)
        $written | Format-Table -AutoSize
        Write-InfoMsg "Edit before relying on these. Onboard New User will offer them on next launch."
    } else {
        Write-Warn "No templates were written."
    }
    Pause-ForUser
}
