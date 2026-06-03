# ============================================================
#  Templates.ps1 — Onboarding role-template loader
#
#  Scans templates/role-*.json, validates the schema, and merges
#  template defaults into operator-supplied user data. Tenant-side
#  validation (does this SKU exist, does this group exist) lives in
#  Onboard.ps1 / BulkOnboard.ps1 where a tenant connection is already
#  available.
# ============================================================

# templates/ now lives one level up from app/ (where the .ps1 files
# are). Prefer $Global:M365RepoRoot when Main.ps1 has set it;
# otherwise walk one level up from $PSScriptRoot and prefer that
# location when it contains a templates/ folder. Falls back to the
# legacy in-folder lookup so dot-sourcing this in isolation (tests)
# still resolves.
function Get-DefaultTemplatesRoot {
    if ($Global:M365RepoRoot) {
        $cand = Join-Path $Global:M365RepoRoot 'templates'
        if (Test-Path -LiteralPath $cand) { return $cand }
    }
    if ($PSScriptRoot) {
        $parent = Split-Path -Parent $PSScriptRoot
        if ($parent) {
            $cand = Join-Path $parent 'templates'
            if (Test-Path -LiteralPath $cand) { return $cand }
        }
        return (Join-Path $PSScriptRoot 'templates')
    }
    if ($env:M365ADMIN_ROOT) {
        $parent = Split-Path -Parent $env:M365ADMIN_ROOT
        if ($parent) {
            $cand = Join-Path $parent 'templates'
            if (Test-Path -LiteralPath $cand) { return $cand }
        }
        return (Join-Path $env:M365ADMIN_ROOT 'templates')
    }
    return (Join-Path (Get-Location).Path 'templates')
}
$script:TemplatesRoot = Get-DefaultTemplatesRoot

$script:TemplateRequiredFields = @('name', 'description', 'usageLocation')
$script:TemplateValidAccess    = @('Full', 'SendAs', 'FullSendAs')

# ---- Internal: PSCustomObject (from ConvertFrom-Json) -> nested hashtable ----
function ConvertTo-TemplateHashtable {
    param($Object)
    if ($null -eq $Object) { return $null }
    if ($Object -is [string]) { return $Object }
    if ($Object.GetType().IsValueType) { return $Object }
    if ($Object -is [System.Collections.IDictionary]) {
        $ht = @{}
        foreach ($k in $Object.Keys) { $ht[$k] = ConvertTo-TemplateHashtable -Object $Object[$k] }
        return $ht
    }
    if ($Object -is [System.Collections.IList] -and -not ($Object -is [string])) {
        return @($Object | ForEach-Object { ConvertTo-TemplateHashtable -Object $_ })
    }
    if ($Object.PSObject -and $Object.PSObject.Properties) {
        $ht = @{}
        foreach ($p in $Object.PSObject.Properties) {
            if ($p.Name -like '_comment*') { continue }
            $ht[$p.Name] = ConvertTo-TemplateHashtable -Object $p.Value
        }
        return $ht
    }
    return $Object
}

function ConvertTo-TenantFolderSlug {
    <#
        Filesystem-safe slug for a tenant key. Used to build the
        per-tenant template folder name. Idempotent.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $s = $Text.ToLowerInvariant()
    $s = $s -replace '[^a-z0-9]+', '-'
    return $s.Trim('-')
}

function Get-TenantTemplatesRoot {
    <#
        Returns the per-tenant template folder path for the active
        SessionState, or $null when there's no tenant context (e.g.
        a fresh Direct session that never picked a registered
        profile). Folder is templates/<tenant-slug>/.
        Tenant key preference: TenantDomain -> TenantName -> TenantId.
    #>
    if (-not (Get-Variable -Name SessionState -Scope Script -ErrorAction SilentlyContinue)) { return $null }
    $key = $null
    if ($script:SessionState.TenantDomain) { $key = $script:SessionState.TenantDomain }
    elseif ($script:SessionState.TenantName -and $script:SessionState.TenantName -ne 'Own Tenant') { $key = $script:SessionState.TenantName }
    elseif ($script:SessionState.TenantId) { $key = $script:SessionState.TenantId }
    if (-not $key) { return $null }
    $slug = ConvertTo-TenantFolderSlug -Text $key
    if (-not $slug) { return $null }
    return Join-Path $script:TemplatesRoot $slug
}

function Get-TemplateSearchRoots {
    <#
        Ordered list of directories to look at for role-*.json. The
        per-tenant folder wins on name collisions (looked at first),
        and templates/ acts as the shared fallback.
    #>
    $roots = @()
    $tenantDir = Get-TenantTemplatesRoot
    if ($tenantDir -and (Test-Path -LiteralPath $tenantDir)) { $roots += $tenantDir }
    if ($script:TemplatesRoot -and (Test-Path -LiteralPath $script:TemplatesRoot)) { $roots += $script:TemplatesRoot }
    return ,$roots
}

function Get-OnboardTemplates {
    <#
        Returns an array of PSCustomObjects describing every valid
        role-*.json template found under tenant-scoped + global
        roots. Templates in the tenant-scoped folder override
        same-name templates in the global folder. Malformed files
        are warned about and skipped, not throw.
    #>
    $roots = @(Get-TemplateSearchRoots)
    if ($roots.Count -eq 0) { return @() }
    $seen = @{}    # key -> $true so first-seen wins (tenant root is first)
    $list = @()
    $files = @()
    foreach ($r in $roots) {
        $files += @(Get-ChildItem -LiteralPath $r -Filter 'role-*.json' -File -ErrorAction SilentlyContinue)
    }
    foreach ($f in ($files | Sort-Object Name)) {
        try {
            $raw = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $missing = @()
            foreach ($req in $script:TemplateRequiredFields) {
                if (-not ($raw.PSObject.Properties.Name -contains $req)) { $missing += $req }
            }
            if ($missing.Count -gt 0) {
                Write-Warn "Skipping template '$($f.Name)' — missing required field(s): $($missing -join ', ')"
                continue
            }
            $key = [System.IO.Path]::GetFileNameWithoutExtension($f.Name).Substring(5)  # strip "role-"
            # Tenant-scoped templates appear first in $files; honor the
            # first-seen-wins rule so a per-tenant role-engineer.json
            # overrides the global one without showing both.
            $dedupKey = $key.ToLowerInvariant()
            if ($seen.ContainsKey($dedupKey)) { continue }
            $seen[$dedupKey] = $true
            $list += [PSCustomObject]@{
                Key         = $key
                Name        = $raw.name
                Description = $raw.description
                Path        = $f.FullName
            }
        } catch {
            Write-Warn "Skipping unreadable template '$($f.Name)': $_"
        }
    }
    return $list
}

function Get-OnboardTemplate {
    <#
        Load and schema-validate a single template by key (file basename
        without the "role-" prefix and ".json" extension). Looks in the
        tenant-scoped folder first, then the global one.
        Returns a hashtable with all template fields populated to
        safe defaults.
    #>
    param([Parameter(Mandatory)][string]$Key)

    $k = ($Key -replace '\.json$', '').ToLowerInvariant() -replace '^role-', ''
    $path = $null
    foreach ($r in @(Get-TemplateSearchRoots)) {
        $candidate = Join-Path $r ("role-{0}.json" -f $k)
        if (Test-Path -LiteralPath $candidate) { $path = $candidate; break }
    }
    if (-not $path) {
        throw "Template 'role-$k' not found in tenant or global templates folder."
    }
    $raw = $null
    try { $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Could not parse '$path': $_" }
    if ($null -eq $raw) { throw "Empty template file: $path" }

    $tpl = ConvertTo-TemplateHashtable -Object $raw

    foreach ($req in $script:TemplateRequiredFields) {
        if (-not $tpl.ContainsKey($req) -or [string]::IsNullOrWhiteSpace([string]$tpl[$req])) {
            throw "Template 'role-$k' missing required field: $req"
        }
    }

    # Defaults for optional fields
    foreach ($pair in @(
        @{ Key='licenseSKUs';          Default = @() },
        @{ Key='securityGroups';       Default = @() },
        @{ Key='distributionLists';    Default = @() },
        @{ Key='sharedMailboxes';      Default = @() },
        @{ Key='teams';                Default = @() },
        @{ Key='oneDrive';             Default = $null },
        @{ Key='defaults';             Default = @{} },
        @{ Key='contractorExpiryDays'; Default = $null }
    )) {
        if (-not $tpl.ContainsKey($pair.Key) -or $null -eq $tpl[$pair.Key]) {
            $tpl[$pair.Key] = $pair.Default
        }
    }

    # Shared mailbox access validation
    foreach ($sm in @($tpl['sharedMailboxes'])) {
        if (-not $sm) { continue }
        $access = [string]$sm['access']
        if ($script:TemplateValidAccess -notcontains $access) {
            throw "Template 'role-$k' has invalid sharedMailbox access '$access' (must be one of: $($script:TemplateValidAccess -join ', '))"
        }
        if ([string]::IsNullOrWhiteSpace([string]$sm['identity'])) {
            throw "Template 'role-$k' has a sharedMailbox entry with no identity"
        }
    }

    # contractorExpiryDays must be a positive int when set
    if ($null -ne $tpl['contractorExpiryDays']) {
        $n = 0
        if (-not [int]::TryParse([string]$tpl['contractorExpiryDays'], [ref]$n) -or $n -le 0) {
            throw "Template 'role-$k' contractorExpiryDays must be a positive integer or null"
        }
        $tpl['contractorExpiryDays'] = $n
    }

    $tpl['__key']  = $k
    $tpl['__path'] = $path
    return $tpl
}

function Resolve-OnboardTemplate {
    <#
        Merge a template into operator-supplied $UserData. Operator
        values always win — template only fills gaps. Returns a new
        hashtable; neither input is mutated. The returned hashtable
        carries the template under the __Template key so downstream
        onboarding steps (license/group/SM assignment) can read it.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Template,
        [hashtable]$UserData = @{}
    )

    $merged = @{}
    foreach ($k in $UserData.Keys) { $merged[$k] = $UserData[$k] }

    if (-not $merged.ContainsKey('UsageLocation') -or [string]::IsNullOrWhiteSpace([string]$merged['UsageLocation'])) {
        $merged['UsageLocation'] = $Template['usageLocation']
    }

    $defaults = $Template['defaults']
    if ($defaults) {
        if ($defaults -isnot [hashtable]) {
            $tmp = @{}; foreach ($p in $defaults.PSObject.Properties) { $tmp[$p.Name] = $p.Value }; $defaults = $tmp
        }
        foreach ($k in $defaults.Keys) {
            if (-not $merged.ContainsKey($k) -or [string]::IsNullOrWhiteSpace([string]$merged[$k])) {
                $merged[$k] = $defaults[$k]
            }
        }
    }

    $merged['__Template'] = $Template
    return $merged
}
