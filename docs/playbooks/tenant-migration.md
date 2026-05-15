# Tenant migration

Moving users between M365 tenants. Most commonly: M&A merger, divestiture, consolidating multiple legacy tenants. The actual mailbox / OneDrive content migration uses Microsoft's tooling — this playbook covers the operational pieces the tool helps with.

## When to use this playbook

- M&A: company A acquired company B, consolidating tenants.
- Divestiture: subset of users moving out to a spun-off tenant.
- IT consolidation: consolidating multiple legacy tenants into one.
- Cloud-to-cloud: migrating from one M365 tenant to another for compliance / region reasons.

## What this tool does + doesn't do

| Phase | Tool can help | Out of scope (use Microsoft / 3rd-party) |
|---|---|---|
| Pre-migration assessment | License inventory, role inventory, group / DL inventory | n/a |
| Coexistence setup | Tenant profile registration | Hybrid mail flow, GAL sync |
| User cutover prep | Disable source, onboard target | Mailbox content migration |
| Mailbox + OneDrive content | n/a | Microsoft's cross-tenant migration |
| Post-migration cleanup | Source tenant deprovisioning | n/a |
| Audit + verification | Full audit log per tenant | n/a |

## Pre-migration assessment (1-2 weeks)

### Register both tenants

```powershell
# Source (the tenant being migrated away from):
Register-Tenant -Name "Source-ABC" -TenantId <source-guid> -AuthMode Interactive

# Target (the tenant being migrated to):
Register-Tenant -Name "Target-XYZ" -TenantId <target-guid> -AuthMode Interactive
```

For unattended migrations, use `CertThumbprint` mode. See [`../getting-started/tenant-setup.md`](../getting-started/tenant-setup.md).

### Inventory the source

```powershell
Switch-Tenant -Name "Source-ABC"

# What licenses are in use?
Get-MgSubscribedSku | Format-Table SkuPartNumber, ConsumedUnits, PrepaidUnits

# What users will need migration?
Get-MgUser -All -Property UserPrincipalName, UserType, AccountEnabled |
    Where-Object { $_.UserType -eq 'Member' -and $_.AccountEnabled -eq $true } |
    Export-Csv -Path .\source-users.csv -NoTypeInformation

# What groups / DLs need recreating?
Get-MgGroup -All -Property DisplayName, GroupTypes, Description |
    Export-Csv -Path .\source-groups.csv -NoTypeInformation
```

Identify users / groups to migrate vs decommission.

### Inventory the target

```powershell
Switch-Tenant -Name "Target-XYZ"

# What licenses are available for the migrating users?
Get-MgSubscribedSku | Format-Table SkuPartNumber, ConsumedUnits, PrepaidUnits

# What's the role-template inventory?
Get-ChildItem .\templates\role-*.json | ForEach-Object { (Get-Content $_.FullName | ConvertFrom-Json).name }
```

If the target lacks SKUs the source uses, the cutover will partially fail unless you procure first.

## Coexistence (1-4 weeks)

Microsoft has tools for this — most commonly **B2B Direct Connect** or **Hybrid Mail Flow** for cross-tenant identity sharing during the transition. M365 Manager doesn't directly enable coexistence; it lives alongside.

During coexistence:

- Run the tool against both tenants via `Switch-Tenant`.
- Each tenant's audit log is independent (per-tenant slug in filename).
- AI sessions are per-tenant — don't move them across.

## Cutover (per-user)

For each user, the cutover is:

1. **Target tenant: onboard the user** with the right SKUs / groups / DLs.
2. **Microsoft tooling: migrate mailbox + OneDrive content** (out of scope here).
3. **Source tenant: offboard the user** (12-step flow with forwarding to the target UPN).

### Bulk-cutover approach

CSV at `migration-batch-1.csv` (target-tenant onboard):

```csv
FirstName,LastName,UserPrincipalName,Manager,UsageLocation,Template
Alice,Smith,alice@target-xyz.com,bob@target-xyz.com,US,sales-rep
Charlie,Davis,charlie@target-xyz.com,dave@target-xyz.com,US,engineer
```

```powershell
# Target tenant:
Switch-Tenant -Name "Target-XYZ"
Invoke-BulkOnboard -Path .\migration-batch-1.csv
```

Then `migration-batch-1-offboard.csv` (source-tenant offboard):

```csv
UserPrincipalName,ForwardTo,ConvertToShared,RemoveFromAllGroups,Reason
alice@source-abc.com,alice@target-xyz.com,yes,yes,M&A migration to Target-XYZ 2026-Q2
charlie@source-abc.com,charlie@target-xyz.com,yes,yes,M&A migration to Target-XYZ 2026-Q2
```

```powershell
# Source tenant:
Switch-Tenant -Name "Source-ABC"
Invoke-BulkOffboard -Path .\migration-batch-1-offboard.csv
```

### Forwarding-window strategy

Set source-tenant mailbox forwarding to the target UPN for 30-90 days after cutover. Any external party still emailing the old address gets routed to the new one. Most tenants do this for 90 days; legal-sensitive industries may extend to 1 year.

After the forwarding window:

```powershell
# Source tenant: remove forwarding + delete users
Switch-Tenant -Name "Source-ABC"

foreach ($upn in (Get-Content .\migrated-users.txt)) {
    # Remove the forwarding (was set 90 days ago):
    Set-Mailbox -Identity $upn -ForwardingAddress $null -ForwardingSmtpAddress $null
    # Delete the user (lands in /directory/deletedItems for 30 days):
    Remove-MgUser -UserId $upn
}
```

## Verification

After cutover, verify per-user:

```powershell
# Target tenant:
Switch-Tenant -Name "Target-XYZ"
Get-MgUser -UserId alice@target-xyz.com -Property AccountEnabled, AssignedLicenses, SignInActivity
Get-MgUserMemberOf -UserId alice@target-xyz.com | Select-Object DisplayName

# Source tenant:
Switch-Tenant -Name "Source-ABC"
Get-MgUser -UserId alice@source-abc.com -Property AccountEnabled, AssignedLicenses
# Expected: AccountEnabled=False, no licenses
```

The dual-tenant audit log (one file per tenant) shows the full migration trail per user.

## Audit handoff

When the migration is complete:

```powershell
# Both tenants -- per-tenant audit bundles
Switch-Tenant -Name "Source-ABC"
$migAuditSource = Read-AuditEntries | Where-Object { $_.description -like 'M&A migration*' }
$migAuditSource | Export-AuditEntriesHtml -Path .\migration-source-audit.html

Switch-Tenant -Name "Target-XYZ"
$migAuditTarget = Read-AuditEntries | Where-Object { $_.description -like 'M&A*' }
$migAuditTarget | Export-AuditEntriesHtml -Path .\migration-target-audit.html
```

Bundle both with the per-tenant inventories. This is what your compliance team will want.

## Decommissioning the source tenant

If the source tenant is going away entirely (full consolidation):

1. **Verify no users remain active** in the source. Anyone still active means a migration was missed.
2. **Cancel source licenses** at the M365 admin center (out of tool's scope).
3. **Remove the tenant profile** from M365 Manager:
   ```powershell
   Remove-Tenant -Name "Source-ABC"
   ```
4. **Archive the source audit logs**. Don't delete — retention obligations apply for the post-migration window.

## Common pitfalls

- **License-SKU mismatches.** Source had Visio Plan 2; target only has Visio Plan 1. Users lose Visio features after cutover. Procure first OR accept the feature loss.
- **DL membership rot.** Bulk-onboarding target-tenant users by template doesn't replicate every DL membership. Audit + manually re-add legacy DLs.
- **External-shared-with relationships.** If `alice@source-abc.com` had Guest access to a 3rd-party SharePoint, that doesn't auto-transfer. The 3rd-party has to re-invite `alice@target-xyz.com`.
- **OneDrive content size.** Microsoft's migration is fast for small drives, slow for many-GB. Plan downtime per user accordingly.
- **MFA re-enrollment.** Target tenant has its own auth. Users will re-enroll their MFA method on first sign-in. TAP issuance speeds this up.

## See also

- [`../concepts/multi-tenant.md`](../concepts/multi-tenant.md) — multi-tenant architecture.
- [`../guides/onboarding.md`](../guides/onboarding.md) — target-tenant onboarding.
- [`../guides/offboarding.md`](../guides/offboarding.md) — source-tenant offboarding.
- [`audit-prep.md`](audit-prep.md) — bundling migration audit evidence for compliance.
