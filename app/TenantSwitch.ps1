# ============================================================
#  TenantSwitch.ps1 -- fast tenant switching + colored banner +
#  audit fingerprint (Phase 6 commit B)
#
#  Switch-Tenant disconnects the current Graph / EXO / SCC / SPO
#  sessions, swaps in a registered tenant profile, and (when the
#  profile carries an app-only credential manifest) re-connects
#  immediately. Interactive profiles defer the reconnect to the
#  next service call so the operator only sees the browser flow
#  if they actually need it.
#
#  The colored top banner makes it visually obvious which tenant
#  you're acting on -- the same set of menus rendered in red vs
#  green text is much harder to mix up than the old plain-white
#  banner.
# ============================================================

$script:TenantBannerPalette = @('Cyan','Green','Yellow','Magenta','Blue','Red','DarkCyan','DarkGreen','DarkYellow','DarkMagenta')

function Get-TenantBannerColor {
    <#
        Stable per-tenant color picked by hashing the tenant name.
        Cycles through TenantBannerPalette so two tenants get
        different colors (collisions only beyond ~10 tenants).
    #>
    param([string]$Name)
    if (-not $Name) { return 'Cyan' }
    $sum = 0
    foreach ($c in $Name.ToCharArray()) { $sum = ($sum + [int]$c) }
    return $script:TenantBannerPalette[$sum % $script:TenantBannerPalette.Count]
}

function Get-CurrentTenantLabel {
    <#
        Returns "[<name>]" with brackets when a tenant profile is
        active, or "" when in legacy direct/own-tenant mode. Used
        by every menu / banner / audit filename.
    #>
    $t = if (Get-Command Get-CurrentTenant -ErrorAction SilentlyContinue) { Get-CurrentTenant } else { $null }
    if ($t -and $t.name) { return "[$($t.name)] " }
    return ''
}

function Write-TenantBanner {
    <#
        Top-of-screen colored bar showing the current tenant.
        Renders nothing in legacy mode so Phase 1-5 callers don't
        see a visual regression.
    #>
    $t = if (Get-Command Get-CurrentTenant -ErrorAction SilentlyContinue) { Get-CurrentTenant } else { $null }
    if (-not $t) { return }
    $color = Get-TenantBannerColor -Name $t.name
    $bar   = "  +=== TENANT: $($t.name) ($($t.tenantId)) ===+"
    Write-Host $bar -ForegroundColor $color
}

function Switch-Tenant {
    <#
        Orchestrator. Steps:
          1. Audit the switch (so the prior-tenant audit log
             still records the intent before context flips).
          2. Disconnect all sessions for the old tenant.
          3. Set-CurrentTenant -- mirrors profile into
             SessionState so subsequent Connect-* calls target
             the new tenant.
          4. If the profile carries app-only creds, eagerly
             reconnect Graph + EXO. Interactive profiles defer.
          5. Audit the post-switch state.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$NoReconnect
    )
    $target = Get-Tenant -Name $Name
    if (-not $target) { Write-ErrorMsg "Tenant '$Name' is not registered. Use Register-Tenant first."; return $false }

    $priorName = if ($script:SessionState -and $script:SessionState.TenantName) { $script:SessionState.TenantName } else { '(none)' }
    if (Get-Command Write-AuditEntry -ErrorAction SilentlyContinue) {
        Write-AuditEntry -EventType 'TenantSwitch' -Detail ("Switching from '{0}' to '{1}'" -f $priorName, $target.name) -ActionType 'TenantSwitch' -Target @{ from = $priorName; to = $target.name; tenantId = $target.tenantId } -Result 'info' | Out-Null
    }

    if (Get-Command Reset-AllSessions -ErrorAction SilentlyContinue) { Reset-AllSessions }
    if (Get-Command Reset-AuditLogPath -ErrorAction SilentlyContinue) { Reset-AuditLogPath }
    Set-CurrentTenant -Name $Name | Out-Null
    # Reset-AllSessions clears TenantId/Name/Domain, so re-apply after.
    $script:SessionState.TenantId     = $target.tenantId
    $script:SessionState.TenantName   = $target.name
    $script:SessionState.TenantDomain = $target.primaryDomain
    $script:SessionState.TenantMode   = if ($target.credentialRef) { 'Profile' } else { 'Direct' }

    if (-not $NoReconnect -and $target.credentialRef) {
        $m = Get-TenantCredentialManifest -Name $Name
        if ($m -and $m.authMode -in 'CertThumbprint','ClientSecret') {
            Connect-TenantAppOnly -Profile $target -Manifest $m | Out-Null
        }
    }

    Write-TenantBanner
    Write-InfoMsg ("Switched to tenant '{0}'. Mode: {1}." -f $target.name, $script:SessionState.TenantMode)
    return $true
}

function Connect-TenantAppOnly {
    <#
        App-only reconnect via Microsoft Graph + Exchange Online
        using the credential manifest stored alongside the
        registered profile. This is the path that makes MSP
        tenant-hopping practical -- no browser flow per tenant.
        Falls back gracefully when the SDK isn't installed or the
        cert isn't in the local store; the operator can still
        type interactive creds at the next prompt.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Profile,
        [Parameter(Mandatory)][hashtable]$Manifest
    )
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-Warn "Microsoft.Graph SDK not installed; app-only reconnect skipped."
        return $false
    }
    try {
        if ($Manifest.authMode -eq 'CertThumbprint') {
            Connect-MgGraph -TenantId $Profile.tenantId -ClientId $Manifest.clientId -CertificateThumbprint $Manifest.thumbprint -NoWelcome -ErrorAction Stop
        } elseif ($Manifest.authMode -eq 'ClientSecret') {
            $sec = ConvertTo-SecureString $Manifest.secret -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($Manifest.clientId, $sec)
            Connect-MgGraph -TenantId $Profile.tenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
        }
        $script:SessionState.MgGraph = $true
        Write-Success ("App-only Graph connected to {0}." -f $Profile.name)
    } catch {
        Write-Warn ("App-only Graph connect failed: {0}. Interactive flow will run on next service call." -f $_.Exception.Message)
        return $false
    }
    if (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
        try {
            if ($Manifest.authMode -eq 'CertThumbprint') {
                Connect-ExchangeOnline -AppId $Manifest.clientId -CertificateThumbprint $Manifest.thumbprint -Organization $Profile.primaryDomain -ShowBanner:$false -ErrorAction Stop
                $script:SessionState.ExchangeOnline = $true
                Write-Success "App-only Exchange Online connected."
            }
        } catch {
            Write-Warn ("App-only EXO connect failed: {0}. Interactive flow will run on next mailbox call." -f $_.Exception.Message)
        }
    }
    return $true
}

function Start-TenantMenu {
    <#
        Top-level "Tenants..." submenu (slot 21). Operators land
        here from the main menu; no chat involvement required.
    #>
    while ($true) {
        Initialize-UI; Write-Banner; Write-TenantBanner
        Show-TenantRegistry
        $sel = Show-Menu -Title "Tenants" -Options @(
            "Switch to tenant...",
            "Register a new tenant",
            "Edit a tenant",
            "Remove a tenant",
            "MSP portfolio dashboard"
        ) -BackLabel "Back to main menu"
        switch ($sel) {
            0 {
                $tenants = Get-Tenants
                if ($tenants.Count -eq 0) { Write-Warn "No tenants registered."; Pause-ForUser; continue }
                $i = Show-Menu -Title "Switch to" -Options @($tenants | ForEach-Object { "$($_.name)  ($($_.tenantId))" })
                if ($i -ge 0) { Switch-Tenant -Name ($tenants[$i].name) | Out-Null; Pause-ForUser }
            }
            1 { Invoke-RegisterTenantWizard; Pause-ForUser }
            2 {
                $tenants = Get-Tenants
                if ($tenants.Count -eq 0) { Write-Warn "No tenants registered."; Pause-ForUser; continue }
                $i = Show-Menu -Title "Edit" -Options @($tenants | ForEach-Object { "$($_.name)" })
                if ($i -ge 0) { Invoke-EditTenantWizard -Name ($tenants[$i].name); Pause-ForUser }
            }
            3 {
                $tenants = Get-Tenants
                if ($tenants.Count -eq 0) { Write-Warn "No tenants registered."; Pause-ForUser; continue }
                $i = Show-Menu -Title "Remove" -Options @($tenants | ForEach-Object { "$($_.name)" })
                if ($i -ge 0) {
                    if (Confirm-Action "Remove '$($tenants[$i].name)' from the registry?") {
                        Remove-Tenant -Name $tenants[$i].name | Out-Null
                    }
                    Pause-ForUser
                }
            }
            4 {
                if (Get-Command Update-MSPDashboard -ErrorAction SilentlyContinue) { Update-MSPDashboard }
                else { Write-Warn "MSP dashboard module not loaded (Phase 6 commit D)." }
                Pause-ForUser
            }
            -1 { return }
        }
    }
}

function Invoke-RegisterTenantWizard {
    <#
        Interactive wrapper around Register-Tenant. Walks the
        operator through the three auth modes with sensible
        defaults and pops the right follow-up prompts.
    #>
    Write-SectionHeader "Register a new tenant"
    $name     = Read-RequiredInput "Friendly name (e.g. Contoso)"
    $tenantId = Read-RequiredInput "Tenant ID (GUID)"
    $domain   = Read-Host "Primary domain (e.g. contoso.onmicrosoft.com) [optional]"
    $spo      = Read-Host "SPO admin URL [optional]"

    $authIdx = Show-Menu -Title "Auth mode" -Options @(
        "Interactive       (browser; no creds at rest)",
        "App + Certificate (recommended for app-only)",
        "App + ClientSecret(stored encrypted, warned)"
    )
    if ($authIdx -lt 0) { return }
    $authMode = @('Interactive','CertThumbprint','ClientSecret')[$authIdx]

    $clientId = $null; $thumb = $null; $secret = $null
    if ($authMode -ne 'Interactive') {
        $clientId = Read-RequiredInput "App (Client) ID"
        if ($authMode -eq 'CertThumbprint') {
            $thumb = Read-RequiredInput "Certificate thumbprint"
        } else {
            $secret = Read-RequiredInput "Client secret (will be encrypted)"
        }
    }
    Register-Tenant -Name $name -TenantId $tenantId -PrimaryDomain $domain -SpoAdminUrl $spo `
        -AuthMode $authMode -ClientId $clientId -CertThumbprint $thumb -ClientSecret $secret | Out-Null
    Write-Success "Tenant '$name' registered."
}

function Invoke-EditTenantWizard {
    param([Parameter(Mandatory)][string]$Name)
    $t = Get-Tenant -Name $Name
    if (-not $t) { Write-Warn "'$Name' not found."; return }
    Write-SectionHeader "Editing $Name"
    $newDomain = Read-Host "Primary domain [current: $($t.primaryDomain)]"
    $newSpo    = Read-Host "SPO admin URL [current: $($t.spoAdminUrl)]"
    $newTags   = Read-Host "Tags (comma-separated) [current: $(@($t.tags) -join ',')]"
    $newNotes  = Read-Host "Notes [current: $($t.notes)]"
    $params = @{ Name = $Name }
    if ($newDomain) { $params.PrimaryDomain = $newDomain }
    if ($newSpo)    { $params.SpoAdminUrl   = $newSpo }
    if ($newTags)   { $params.Tags          = @($newTags -split '\s*,\s*' | Where-Object { $_ }) }
    if ($newNotes)  { $params.Notes         = $newNotes }
    Update-Tenant @params | Out-Null
    Write-Success "Updated."
}

function Read-RequiredInput {
    param([string]$Prompt)
    while ($true) {
        $v = Read-Host $Prompt
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        Write-Warn "Required."
    }
}
