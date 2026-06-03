# ============================================================
#  Archive.ps1 - Mailbox Archiving Management
# ============================================================

function Start-ArchiveManagement {
    Write-SectionHeader "Mailbox Archiving"

    if (-not (Connect-ForTask "Archive")) {
        Write-ErrorMsg "Could not connect to required services."
        Pause-ForUser; return
    }

    # ---- Identify user ----
    $user = Resolve-UserIdentity -PromptText "Enter user name or email"
    if ($null -eq $user) { Pause-ForUser; return }

    $upn = $user.UserPrincipalName

    # ---- Check current archive status ----
    Write-SectionHeader "Archive Status for $($user.DisplayName)"

    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop

        $archiveEnabled = $mailbox.ArchiveStatus -eq "Active"
        $archiveName    = $mailbox.ArchiveName
        $archiveGuid    = $mailbox.ArchiveGuid

        Write-StatusLine "Archive Enabled" $(if ($archiveEnabled) { "Yes" } else { "No" }) `
            $(if ($archiveEnabled) { "Green" } else { "Red" })

        if ($archiveEnabled) {
            Write-StatusLine "Archive Name" $archiveName "White"
            Write-StatusLine "Archive GUID" $archiveGuid "White"
        }

        # Check retention policy
        $retentionPolicy = $mailbox.RetentionPolicy
        if ($retentionPolicy) {
            Write-StatusLine "Retention Policy" $retentionPolicy "Cyan"
        } else {
            Write-StatusLine "Retention Policy" "(none)" "Gray"
        }
    }
    catch {
        Write-ErrorMsg "Could not retrieve mailbox info: $_"
        Pause-ForUser; return
    }

    Write-Host ""

    if ($archiveEnabled) {
        Write-InfoMsg "Archive is already enabled for this user."

        $changePolicy = Show-Menu -Title "Options" -Options @(
            "Change retention policy",
            "View archive details only"
        ) -BackLabel "Done"

        if ($changePolicy -eq 0) {
            Set-ArchiveRetentionPolicy -UPN $upn
        }
    }
    else {
        # ---- Enable archive ----
        if (Confirm-Action "Enable archive mailbox for $upn and start archiving immediately?") {
            $ok1 = Invoke-Action -Description ("Enable archive mailbox for {0}" -f $upn) -Action {
                Enable-Mailbox -Identity $upn -Archive -ErrorAction Stop; $true
            }
            if ($ok1 -and -not (Get-PreviewMode)) { Write-Success "Archive mailbox enabled for $upn." }
            if (-not $ok1 -and -not (Get-PreviewMode)) { Pause-ForUser; return }

            Write-InfoMsg "Starting Managed Folder Assistant to initiate archiving..."
            $ok2 = Invoke-Action -Description ("Start Managed Folder Assistant on {0}" -f $upn) -Action {
                Start-ManagedFolderAssistant -Identity $upn -ErrorAction Stop; $true
            }
            if ($ok2 -and -not (Get-PreviewMode)) { Write-Success "Managed Folder Assistant started. Archiving will begin processing." }

            # Offer to set retention policy
            $setPolicy = Show-Menu -Title "Set a retention policy?" -Options @(
                "Yes, choose a retention policy",
                "No, use default"
            ) -BackLabel "Skip"

            if ($setPolicy -eq 0) {
                Set-ArchiveRetentionPolicy -UPN $upn
            }
        }
    }

    Write-Success "Archive management complete."
    Pause-ForUser
}

function Set-ArchiveRetentionPolicy {
    param([string]$UPN)

    Write-SectionHeader "Available Retention Policies"

    try {
        $policies = Get-RetentionPolicy -ErrorAction Stop
        if ($policies.Count -eq 0) {
            Write-Warn "No retention policies found in the tenant."
            return
        }

        $policyLabels = $policies | ForEach-Object {
            $tags = ($_.RetentionPolicyTagLinks | ForEach-Object { $_.Name }) -join ", "
            if ([string]::IsNullOrWhiteSpace($tags)) { $tags = "(no tags)" }
            "$($_.Name)  [$tags]"
        }

        $sel = Show-Menu -Title "Select a retention policy" -Options $policyLabels -BackLabel "Cancel"
        if ($sel -eq -1) { return }

        $chosenPolicy = $policies[$sel]

        if (Confirm-Action "Apply retention policy '$($chosenPolicy.Name)' to $UPN?") {
            $ok1 = Invoke-Action -Description ("Apply retention policy '{0}' to {1}" -f $chosenPolicy.Name, $UPN) -Action {
                Set-Mailbox -Identity $UPN -RetentionPolicy $chosenPolicy.Name -ErrorAction Stop; $true
            }
            if ($ok1 -and -not (Get-PreviewMode)) { Write-Success "Retention policy '$($chosenPolicy.Name)' applied." }

            Write-InfoMsg "Starting Managed Folder Assistant to process immediately..."
            $ok2 = Invoke-Action -Description ("Start Managed Folder Assistant on {0}" -f $UPN) -Action {
                Start-ManagedFolderAssistant -Identity $UPN -ErrorAction Stop; $true
            }
            if ($ok2 -and -not (Get-PreviewMode)) { Write-Success "Managed Folder Assistant started." }
        }
    }
    catch {
        Write-ErrorMsg "Retention policy error: $_"
    }
}
