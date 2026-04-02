# ============================================================
#  License.ps1 - Add / Remove License Management
# ============================================================

function Start-LicenseManagement {
    Write-SectionHeader "License Management"

    if (-not (Connect-ForTask "License")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ---- Identify user ----
    $user = Resolve-UserIdentity -PromptText "Enter user name or email"
    if ($null -eq $user) { Pause-ForUser; return }

    $upn = $user.UserPrincipalName

    try {
        $msolUser = Get-MsolUser -UserPrincipalName $upn -ErrorAction Stop
    } catch {
        Write-ErrorMsg "Could not retrieve MSOnline user: $_"
        Pause-ForUser; return
    }

    # ---- Show current licenses ----
    Write-SectionHeader "Current Licenses for $($user.DisplayName)"
    $currentLicenses = $msolUser.Licenses
    if ($currentLicenses.Count -eq 0) {
        Write-InfoMsg "This user has no licenses assigned."
    } else {
        for ($i = 0; $i -lt $currentLicenses.Count; $i++) {
            Write-Host "    [$($i+1)] $($currentLicenses[$i].AccountSkuId)" -ForegroundColor White
        }
    }
    Write-Host ""

    # ---- Add or Remove ----
    $action = Show-Menu -Title "What would you like to do?" -Options @(
        "Add license(s)",
        "Remove license(s)"
    ) -BackLabel "Cancel"

    if ($action -eq -1) { return }

    if ($action -eq 0) {
        # ---- ADD ----
        Write-SectionHeader "Available Tenant Licenses"
        try {
            $allSkus = Get-MsolAccountSku -ErrorAction Stop
            $available = $allSkus | ForEach-Object {
                $free = $_.ActiveUnits - $_.ConsumedUnits
                [PSCustomObject]@{
                    SkuPartNumber = $_.SkuPartNumber
                    AccountSkuId  = $_.AccountSkuId
                    Total         = $_.ActiveUnits
                    Used          = $_.ConsumedUnits
                    Free          = $free
                }
            }

            $labels = $available | ForEach-Object {
                "$($_.SkuPartNumber)  [Total: $($_.Total) | Used: $($_.Used) | Free: $($_.Free)]"
            }

            $selected = Show-MultiSelect -Title "Select license(s) to add" -Options $labels `
                -Prompt "Enter license number(s) (e.g. 1,3,5)"

            foreach ($idx in $selected) {
                $sku = $available[$idx]
                if ($sku.Free -le 0) {
                    Write-Warn "$($sku.SkuPartNumber) has no available seats. Skipping."
                    continue
                }
                if (Confirm-Action "Assign '$($sku.SkuPartNumber)' to $upn?") {
                    try {
                        Set-MsolUserLicense -UserPrincipalName $upn `
                            -AddLicenses $sku.AccountSkuId -ErrorAction Stop
                        Write-Success "License '$($sku.SkuPartNumber)' assigned."
                    } catch {
                        Write-ErrorMsg "Failed to assign $($sku.SkuPartNumber): $_"
                    }
                }
            }
        } catch {
            Write-ErrorMsg "Could not retrieve tenant licenses: $_"
        }
    }
    else {
        # ---- REMOVE ----
        if ($currentLicenses.Count -eq 0) {
            Write-Warn "No licenses to remove."
            Pause-ForUser; return
        }

        $licLabels = $currentLicenses | ForEach-Object { $_.AccountSkuId }
        $selected = Show-MultiSelect -Title "Select license(s) to remove" -Options $licLabels `
            -Prompt "Enter license number(s) (e.g. 1,3)"

        foreach ($idx in $selected) {
            $lic = $currentLicenses[$idx]
            if (Confirm-Action "Remove '$($lic.AccountSkuId)' from $upn?") {
                try {
                    Set-MsolUserLicense -UserPrincipalName $upn `
                        -RemoveLicenses $lic.AccountSkuId -ErrorAction Stop
                    Write-Success "Removed: $($lic.AccountSkuId)"
                } catch {
                    Write-ErrorMsg "Failed to remove $($lic.AccountSkuId): $_"
                }
            }
        }
    }

    Write-Success "License management complete."
    Pause-ForUser
}
