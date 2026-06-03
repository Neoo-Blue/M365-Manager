# ============================================================
#  TenantOverrides.ps1 -- per-tenant configuration overrides
#  (Phase 6 commit E)
#
#  Some config keys make sense to override per-tenant rather
#  than globally:
#    - notification recipients (different IT contact per customer)
#    - stale-guest threshold (one customer's "stale" is 30 days,
#      another's is 180)
#    - OneDrive retention default
#    - monthly AI budget cap
#    - default role templates
#    - license prices (different SKU contracts per customer)
#
#  An operator can drop tenant-overrides/<tenant-name>.json
#  alongside the global config to override any subset of keys.
#  Resolution order (last wins):
#    1. global ai_config.json / Notifications config
#    2. <stateDir>/tenant-overrides/<name>.json
#    3. environment variable (M365MGR_<KEY>)
#    4. CLI flag (caller-supplied)
# ============================================================

function Get-TenantOverridesDir {
    $base = if (Get-Command Get-StateDirectory -ErrorAction SilentlyContinue) { Get-StateDirectory } else { (Get-Location).Path }
    $d = Join-Path $base 'tenant-overrides'
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    return $d
}

function Get-TenantOverrideFilePath {
    param([Parameter(Mandatory)][string]$Name)
    $slug = ($Name.ToLower() -replace '[^a-z0-9]+','-').Trim('-')
    return (Join-Path (Get-TenantOverridesDir) ("$slug.json"))
}

function Get-TenantOverrides {
    <#
        Read the override file for one tenant. Returns an empty
        hashtable when the file doesn't exist (so callers can
        treat absence as "no overrides", not as an error).
    #>
    param([Parameter(Mandatory)][string]$Name)
    $p = Get-TenantOverrideFilePath -Name $Name
    if (-not (Test-Path -LiteralPath $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) { $h[$prop.Name] = $prop.Value }
        return $h
    } catch {
        Write-Warn "tenant-overrides/$Name.json malformed: $($_.Exception.Message)"
        return @{}
    }
}

function Save-TenantOverrides {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][hashtable]$Overrides)
    $p = Get-TenantOverrideFilePath -Name $Name
    Set-Content -LiteralPath $p -Value ($Overrides | ConvertTo-Json -Depth 8) -Encoding UTF8 -Force
}

function Get-EffectiveConfig {
    <#
        Compute the effective config value for one key.
        Resolution order, last wins:
          1. -GlobalConfig hashtable value (caller supplies)
          2. tenant override file (current tenant if -Tenant not set)
          3. env var M365MGR_<KEY uppercased, '.' -> '_'>
          4. -CliValue (caller-supplied; usually the explicit CLI flag)

        Returns $null when no source supplies the key. Treats
        $null and '' as "not present" so empty CLI values don't
        clobber legitimate overrides.

        Example:
          Get-EffectiveConfig -Key 'StaleGuestDays' `
              -GlobalConfig $Config -CliValue $PSBoundParameters.Days
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Key,
        [hashtable]$GlobalConfig,
        [string]$Tenant,
        $CliValue
    )
    $resolved = $null

    if ($GlobalConfig -and $GlobalConfig.ContainsKey($Key)) {
        $v = $GlobalConfig[$Key]
        if ($null -ne $v -and -not [string]::IsNullOrEmpty([string]$v)) { $resolved = $v }
    }

    $tenantName = $Tenant
    if (-not $tenantName -and (Get-Command Get-CurrentTenant -ErrorAction SilentlyContinue)) {
        $t = Get-CurrentTenant
        if ($t) { $tenantName = $t.name }
    }
    if ($tenantName) {
        $over = Get-TenantOverrides -Name $tenantName
        if ($over.ContainsKey($Key)) {
            $v = $over[$Key]
            if ($null -ne $v -and -not [string]::IsNullOrEmpty([string]$v)) { $resolved = $v }
        }
    }

    $envKey = 'M365MGR_' + ($Key.ToUpper() -replace '\.','_')
    $envVal = [System.Environment]::GetEnvironmentVariable($envKey)
    if ($envVal) { $resolved = $envVal }

    if ($null -ne $CliValue -and -not [string]::IsNullOrEmpty([string]$CliValue)) {
        $resolved = $CliValue
    }
    return $resolved
}

function Show-TenantOverrides {
    <#
        Pretty-print one tenant's overrides. Empty file or no
        file prints a friendly message rather than an error.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $p = Get-TenantOverrideFilePath -Name $Name
    $o = Get-TenantOverrides -Name $Name
    Write-Host ""
    Write-Host ("  TENANT OVERRIDES -- {0}" -f $Name) -ForegroundColor $script:Colors.Title
    Write-Host ("  File: {0}" -f $p) -ForegroundColor DarkGray
    if ($o.Count -eq 0) { Write-Host "  (no overrides set)" -ForegroundColor DarkGray; Write-Host ""; return }
    foreach ($k in @($o.Keys | Sort-Object)) {
        Write-StatusLine $k ([string]$o[$k]) 'White'
    }
    Write-Host ""
}

function Edit-TenantOverrides {
    <#
        Drop the override JSON into $EDITOR (or notepad / nano)
        so the operator can hand-edit. Re-validates as JSON on
        save; leaves the prior file in place if the edit produced
        invalid JSON.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $p = Get-TenantOverrideFilePath -Name $Name
    if (-not (Test-Path -LiteralPath $p)) {
        # Seed an empty {} so the operator has something to edit.
        Set-Content -LiteralPath $p -Value '{}' -Encoding UTF8 -Force
    }
    $editor = $env:EDITOR
    if (-not $editor) { $editor = if ($env:LOCALAPPDATA) { 'notepad' } else { 'nano' } }
    Write-InfoMsg "Opening $editor on $p ..."
    & $editor $p
    try {
        $null = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        Write-Success "Overrides saved."
        if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
            Write-AuditEntry -EventType 'TenantOverrideEdit' -Detail ("Edited overrides for '{0}'" -f $Name) -ActionType 'TenantOverrideEdit' -Target @{ name = $Name } -Result 'ok' | Out-Null
        }
    } catch {
        Write-ErrorMsg "Edited file no longer parses as JSON. Fix manually: $p"
    }
}

# List of keys that are PER-TENANT overridable. Module callers
# should consult this rather than reading config directly when
# the value is one that makes sense to vary by customer.
$script:TenantOverridableKeys = @(
    'StaleGuestDays',
    'OneDriveRetentionDays',
    'OneDriveRetentionPolicy',
    'Notifications.Recipients',
    'Notifications.TeamsWebhook',
    'Notifications.SmtpFrom',
    'AI.MonthlyBudgetUsd',
    'AI.AlertAtPct',
    'AI.AutoPlanThreshold',
    'LicensePrices',
    'DefaultRoleTemplate',
    'BreakGlassReminderDays',
    'AuditRetentionDays'
)

function Get-TenantOverridableKeys { return $script:TenantOverridableKeys }
