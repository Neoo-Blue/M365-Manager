# ============================================================
#  TenantRegistry.ps1 -- per-tenant profile registry (Phase 6)
#
#  Replaces the ephemeral in-memory tenant context with a
#  persistent on-disk registry. Each profile records:
#    - human name
#    - tenantId (Azure AD GUID)
#    - primaryDomain (acme.onmicrosoft.com)
#    - spoAdminUrl (https://acme-admin.sharepoint.com)
#    - credentialRef (key into the encrypted secret store)
#    - tags / notes / lastUsed
#
#  The metadata file <stateDir>/tenants.json is plaintext --
#  no secrets inside. Per-tenant credentials live in
#  <stateDir>/secrets/tenant-<name>.dat encrypted via
#  Protect-Secret (DPAPI on Windows, falls back to B64 on POSIX
#  with a warning so operators can opt into ephemeral mode
#  instead of leaving creds at rest).
# ============================================================

if ($null -eq (Get-Variable -Name CurrentTenantProfile -Scope Script -ErrorAction SilentlyContinue)) {
    $script:CurrentTenantProfile = $null
}

function Get-TenantRegistryPath { return (Join-Path (Get-StateDirectory) 'tenants.json') }
function Get-TenantSecretsDir {
    $d = Join-Path (Get-StateDirectory) 'secrets'
    if (-not (Test-Path -LiteralPath $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        if (-not $env:LOCALAPPDATA -and (Get-Command chmod -ErrorAction SilentlyContinue)) { try { & chmod 700 $d 2>$null } catch {} }
    }
    return $d
}

function Get-Tenants {
    <#
        Read tenants.json. Returns an array of hashtables; empty
        if the file doesn't exist yet. Always returns an array
        (never $null) so callers can foreach safely.
    #>
    $p = Get-TenantRegistryPath
    if (-not (Test-Path -LiteralPath $p)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json -ErrorAction Stop
        $out = @()
        foreach ($t in @($raw)) {
            $h = @{}
            foreach ($pp in $t.PSObject.Properties) { $h[$pp.Name] = $pp.Value }
            if (-not $h.tags) { $h.tags = @() } else { $h.tags = @($h.tags) }
            $out += ,$h
        }
        return $out
    } catch {
        Write-Warn "tenants.json malformed: $($_.Exception.Message)"
        return @()
    }
}

function Save-TenantRegistry {
    <#
        Atomically write the registry. Sort by name so diffs are
        clean if the operator checks the file into a config-mgmt
        system.
    #>
    # AllowEmptyCollection so Remove-Tenant on the last remaining
    # tenant can pass @() through without a binding error.
    param([Parameter(Mandatory)][AllowEmptyCollection()][array]$Tenants)
    $sorted = @($Tenants | Sort-Object { [string]$_.name })
    $p = Get-TenantRegistryPath
    $tmp = "$p.tmp"
    # -AsArray keeps a single-tenant registry serialized as [...] so
    # the read path's @(ConvertFrom-Json) wrap stays consistent.
    Set-Content -LiteralPath $tmp -Value ($sorted | ConvertTo-Json -Depth 8 -AsArray) -Encoding UTF8 -Force
    Move-Item -LiteralPath $tmp -Destination $p -Force
}

function Get-Tenant {
    <#
        Lookup by exact name (case-insensitive). Returns the
        hashtable or $null. Use Get-Tenants to enumerate.
    #>
    param([Parameter(Mandatory)][string]$Name)
    foreach ($t in (Get-Tenants)) { if ([string]::Equals([string]$t.name, $Name, [StringComparison]::OrdinalIgnoreCase)) { return $t } }
    return $null
}

function Get-CurrentTenant { return $script:CurrentTenantProfile }

function Set-CurrentTenant {
    <#
        Make a tenant profile the current context. Stamps lastUsed
        and writes to the registry. Does NOT reconnect -- that's
        Switch-Tenant's job (commit B). Returns the profile.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $t = Get-Tenant -Name $Name
    if (-not $t) { throw "Tenant '$Name' not registered. Use Register-Tenant first." }
    $t.lastUsed = (Get-Date).ToUniversalTime().ToString('o')
    $all = @(Get-Tenants | Where-Object { $_.name -ne $t.name }) + ,$t
    Save-TenantRegistry -Tenants $all
    $script:CurrentTenantProfile = $t
    # Mirror into legacy SessionState so the rest of the code keeps working.
    if ($script:SessionState) {
        $script:SessionState.TenantId     = $t.tenantId
        $script:SessionState.TenantName   = $t.name
        $script:SessionState.TenantDomain = $t.primaryDomain
        $script:SessionState.TenantMode   = if ($t.credentialRef) { 'Profile' } else { 'Direct' }
    }
    return $t
}

function Register-Tenant {
    <#
        Add a new tenant profile. The credential payload is
        optional; if supplied it's encrypted and stored as
        secrets/tenant-<name>.dat with a sidecar manifest that
        records the auth mode + ClientId + (for cert) thumbprint.
        Storing a client-secret at rest is supported but the
        operator gets a warning that cert+thumbprint is safer.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$TenantId,
        [string]$PrimaryDomain,
        [string]$SpoAdminUrl,
        [string]$ClientId,
        [ValidateSet('CertThumbprint','ClientSecret','Interactive')][string]$AuthMode = 'Interactive',
        [string]$CertThumbprint,
        [string]$ClientSecret,
        [string[]]$Tags = @(),
        [string]$Notes
    )
    if (Get-Tenant -Name $Name) { throw "Tenant '$Name' is already registered. Use Update-Tenant." }

    $credRef = $null
    if ($AuthMode -ne 'Interactive') {
        if ($AuthMode -eq 'ClientSecret') {
            Write-Warn "Storing a client_secret on disk -- cert+thumbprint is safer. Rotate this secret regularly."
        }
        $credRef = "tenant-" + ($Name.ToLower() -replace '[^a-z0-9]+','-')
        $manifest = [ordered]@{
            schemaVersion = 1
            authMode      = $AuthMode
            clientId      = $ClientId
            thumbprint    = $CertThumbprint
            secret        = $null
        }
        if ($AuthMode -eq 'ClientSecret' -and $ClientSecret) {
            $manifest.secret = Protect-Secret -PlainText $ClientSecret
        }
        $secretFile = Join-Path (Get-TenantSecretsDir) ($credRef + '.dat')
        Set-Content -LiteralPath $secretFile -Value ($manifest | ConvertTo-Json -Depth 6) -Encoding UTF8 -Force
    }

    $t = @{
        name          = $Name
        tenantId      = $TenantId
        primaryDomain = $PrimaryDomain
        spoAdminUrl   = $SpoAdminUrl
        credentialRef = $credRef
        tags          = $Tags
        notes         = $Notes
        lastUsed      = $null
        createdUtc    = (Get-Date).ToUniversalTime().ToString('o')
    }
    Save-TenantRegistry -Tenants (@(Get-Tenants) + ,$t)
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'TenantRegister' -Detail ("Registered tenant '{0}' ({1})" -f $Name, $TenantId) -ActionType 'TenantRegister' -Target @{ name = $Name; tenantId = $TenantId; authMode = $AuthMode } -Result 'ok' | Out-Null
    }
    return $t
}

function Update-Tenant {
    <#
        Partial update -- only supplied params are touched.
        Credential rotation: supply -ClientSecret to re-encrypt
        the secret manifest; supply -CertThumbprint to update
        the manifest's thumbprint. Use Remove-Tenant + Register-
        Tenant if the AuthMode itself changes.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$TenantId,
        [string]$PrimaryDomain,
        [string]$SpoAdminUrl,
        [string]$ClientId,
        [string]$CertThumbprint,
        [string]$ClientSecret,
        [string[]]$Tags,
        [string]$Notes
    )
    $t = Get-Tenant -Name $Name
    if (-not $t) { throw "Tenant '$Name' not registered." }
    if ($PSBoundParameters.ContainsKey('TenantId'))       { $t.tenantId      = $TenantId }
    if ($PSBoundParameters.ContainsKey('PrimaryDomain'))  { $t.primaryDomain = $PrimaryDomain }
    if ($PSBoundParameters.ContainsKey('SpoAdminUrl'))    { $t.spoAdminUrl   = $SpoAdminUrl }
    if ($PSBoundParameters.ContainsKey('Tags'))           { $t.tags          = $Tags }
    if ($PSBoundParameters.ContainsKey('Notes'))          { $t.notes         = $Notes }

    if ($t.credentialRef) {
        $secretFile = Join-Path (Get-TenantSecretsDir) ($t.credentialRef + '.dat')
        if (Test-Path -LiteralPath $secretFile) {
            $manifest = Get-Content -LiteralPath $secretFile -Raw | ConvertFrom-Json
            $h = @{}
            foreach ($p in $manifest.PSObject.Properties) { $h[$p.Name] = $p.Value }
            if ($PSBoundParameters.ContainsKey('ClientId'))       { $h.clientId   = $ClientId }
            if ($PSBoundParameters.ContainsKey('CertThumbprint')) { $h.thumbprint = $CertThumbprint }
            if ($PSBoundParameters.ContainsKey('ClientSecret'))   { $h.secret     = if ($ClientSecret) { Protect-Secret -PlainText $ClientSecret } else { $null } }
            Set-Content -LiteralPath $secretFile -Value ($h | ConvertTo-Json -Depth 6) -Encoding UTF8 -Force
        }
    }

    $all = @(Get-Tenants | Where-Object { $_.name -ne $t.name }) + ,$t
    Save-TenantRegistry -Tenants $all
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'TenantUpdate' -Detail ("Updated tenant '{0}'" -f $Name) -ActionType 'TenantUpdate' -Target @{ name = $Name } -Result 'ok' | Out-Null
    }
    return $t
}

function Remove-Tenant {
    <#
        Drop the profile + its secret manifest. Audit-only --
        does NOT touch the actual Azure AD tenant. Operator can
        re-Register with the same name later.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $t = Get-Tenant -Name $Name
    if (-not $t) { return $false }
    if ($t.credentialRef) {
        $f = Join-Path (Get-TenantSecretsDir) ($t.credentialRef + '.dat')
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
    }
    Save-TenantRegistry -Tenants @(Get-Tenants | Where-Object { $_.name -ne $t.name })
    if ($script:CurrentTenantProfile -and $script:CurrentTenantProfile.name -eq $Name) { $script:CurrentTenantProfile = $null }
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'TenantRemove' -Detail ("Removed tenant '{0}'" -f $Name) -ActionType 'TenantRemove' -Target @{ name = $Name } -Result 'ok' | Out-Null
    }
    return $true
}

function Get-TenantCredentialManifest {
    <#
        Return the decrypted credential manifest as a hashtable,
        or $null if the tenant has no credentialRef. Used by
        Switch-Tenant to call Connect-Graph / Connect-EXO with
        the right principal.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $t = Get-Tenant -Name $Name
    if (-not $t -or -not $t.credentialRef) { return $null }
    $f = Join-Path (Get-TenantSecretsDir) ($t.credentialRef + '.dat')
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json -ErrorAction Stop
        $h = @{}
        foreach ($p in $raw.PSObject.Properties) { $h[$p.Name] = $p.Value }
        if ($h.secret) { $h.secret = Unprotect-Secret -Stored ([string]$h.secret) }
        return $h
    } catch {
        Write-Warn "Failed to read credential manifest for '$Name': $($_.Exception.Message)"
        return $null
    }
}

function Show-TenantRegistry {
    <#
        Pretty-print the registry. Hides the credential payload --
        only the auth mode + clientId / thumbprint hint surface.
    #>
    $rows = Get-Tenants
    if ($rows.Count -eq 0) {
        Write-Host "  (no tenants registered -- use Register-Tenant or 'Tenants' menu)" -ForegroundColor DarkGray
        return
    }
    Write-Host ""
    Write-Host "  TENANT REGISTRY" -ForegroundColor $script:Colors.Title
    foreach ($t in $rows) {
        $cur = if ($script:CurrentTenantProfile -and $script:CurrentTenantProfile.name -eq $t.name) { '*' } else { ' ' }
        $auth = '(interactive)'
        if ($t.credentialRef) {
            $m = Get-TenantCredentialManifest -Name $t.name
            if ($m) { $auth = "($($m.authMode) clientId=$([string]$m.clientId))" }
        }
        $last = if ($t.lastUsed) { ([datetime]$t.lastUsed).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { 'never' }
        $tagStr = if ($t.tags) { '[' + (@($t.tags) -join ',') + ']' } else { '' }
        Write-Host ("  {0} {1,-24} {2,-40} {3} {4} last={5}" -f $cur, $t.name, $t.tenantId, $tagStr, $auth, $last) -ForegroundColor White
        if ($t.notes) { Write-Host ("        notes: {0}" -f $t.notes) -ForegroundColor DarkGray }
    }
    Write-Host ""
}

function Test-FirstRunMigration {
    <#
        If the registry is empty but the legacy SessionState has
        a tenant context (operator has connected interactively),
        offer to register the current tenant as a profile so
        future runs can /tenant switch back to it without the
        partner-center / manual-id flow.
    #>
    if ((Get-Tenants).Count -gt 0) { return }
    if (-not $script:SessionState -or -not $script:SessionState.TenantId) { return }
    if (Get-NonInteractiveMode) { return }
    Write-Host ""
    Write-Host "  This is the first run with the new tenant registry." -ForegroundColor Yellow
    Write-Host "  Register the currently-connected tenant for future fast switches?" -ForegroundColor White
    if (Confirm-Action ("Register '{0}' ({1})?" -f $script:SessionState.TenantName, $script:SessionState.TenantId)) {
        Register-Tenant -Name $script:SessionState.TenantName -TenantId $script:SessionState.TenantId -PrimaryDomain $script:SessionState.TenantDomain | Out-Null
        Set-CurrentTenant -Name $script:SessionState.TenantName | Out-Null
        Write-InfoMsg "Registered. Use Switch-Tenant or /tenant to swap later."
    }
}
