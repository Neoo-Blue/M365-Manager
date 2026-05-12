# ============================================================
#  Templates.ps1 — Onboarding role-template loader
#
#  Scans templates/role-*.json, validates the schema, and merges
#  template defaults into operator-supplied user data. Tenant-side
#  validation (does this SKU exist, does this group exist) lives in
#  Onboard.ps1 / BulkOnboard.ps1 where a tenant connection is already
#  available.
# ============================================================

$script:TemplatesRoot = $null
if ($PSScriptRoot) {
    $script:TemplatesRoot = Join-Path $PSScriptRoot 'templates'
} elseif ($env:M365ADMIN_ROOT) {
    $script:TemplatesRoot = Join-Path $env:M365ADMIN_ROOT 'templates'
} elseif ($MyInvocation.MyCommand.Path) {
    $script:TemplatesRoot = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'templates'
} else {
    $script:TemplatesRoot = Join-Path (Get-Location).Path 'templates'
}

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

function Get-OnboardTemplates {
    <#
        Returns an array of PSCustomObjects describing every valid
        role-*.json template found under templates/. Malformed files
        are warned about and skipped, not throw.
    #>
    if (-not (Test-Path -LiteralPath $script:TemplatesRoot)) { return @() }
    $files = @(Get-ChildItem -LiteralPath $script:TemplatesRoot -Filter 'role-*.json' -File -ErrorAction SilentlyContinue)
    $list = @()
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
        without the "role-" prefix and ".json" extension). Returns a
        hashtable with all template fields populated to safe defaults.
    #>
    param([Parameter(Mandatory)][string]$Key)

    $k = ($Key -replace '\.json$', '').ToLowerInvariant() -replace '^role-', ''
    $path = Join-Path $script:TemplatesRoot ("role-{0}.json" -f $k)
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Template 'role-$k' not found at $path"
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
