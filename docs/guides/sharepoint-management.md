# SharePoint management

Site provisioning, sharing controls, and outbound-share audits.

## Menu

**Slot 18 → SharePoint**

```
  1. List sites
  2. Create a site (from template)
  3. Add / remove a site owner
  4. List a user's outbound shares (last N days)
  5. Revoke a single share
  6. Revoke ALL of a user's outbound shares (offboarding cleanup)
  7. Stale-site report
```

## Prereqs

- `SharePoint Administrator` role.
- The tool needs your tenant's SPO admin URL (e.g. `https://contoso-admin.sharepoint.com`). On first connection, the tool prompts for it and caches at `<stateDir>\spo-tenants.json` keyed by tenant ID.

## Creating a site

**Slot 18 → option 2.** Pick a template from `templates/site-*.json`. Two ship:

- `templates/site-project.json` — collaboration with external sharing disabled.
- `templates/site-team.json` — Microsoft Team-connected, full M365 group.

Schema:

```jsonc
{
  "name":          "project",
  "description":   "Project collaboration site with external sharing disabled.",
  "siteUrl":       "/sites/proj-{slug}",
  "title":         "Project: {project}",
  "owner":         "{owner}",
  "template":      "STS#3",
  "sharing":       "ExistingExternalUserSharingOnly",
  "storageQuotaMb": 1024,
  "groupConnected": false
}
```

Placeholders (`{slug}`, `{project}`, `{owner}`) get filled at create time from interactive prompts.

## Managing site owners

Owners can be added or removed individually. Both write audit entries with reverse recipes:

```
[+] Added 'alice@contoso.com' as owner of /sites/marketing
    actionType: AddSiteOwner
    reverseType: RemoveSiteOwner
```

The reverse handler (`Undo.ps1`) re-runs the inverse if you `Invoke-Undo` the entry.

## Outbound shares audit

The most operationally important feature of this module. Lists shares one user has created in the last N days.

```powershell
Get-UserOutboundShares -UPN alice@contoso.com -LookbackDays 7 | Format-Table

# Output:
# SharedAtUtc           TargetUserOrEmail              ItemName           SiteUrl                  Permission
# 2026-05-14T17:41:08Z  external@protonmail.example    Q1-Customers.xlsx  /sites/sales             edit
# 2026-05-12T09:22:55Z  bob@contoso.com                proposal.pdf       /sites/sales             read
```

Used by:

- Phase 3 offboarding (clean up shares the leaver created).
- Phase 7 incident response (step 10 — `AuditShares`).
- Phase 7 `Detect-MassExternalShare` (auto-detection).

The function reads UAL operations `SharingSet`, `AnonymousLinkCreated`, `SecureLinkCreated` filtered by user + time window.

## Revoking shares

Per-share or bulk:

**Slot 18 → option 5** prompts for the share id; `Revoke-Share -ShareId <id>` runs the inverse.

**Slot 18 → option 6** runs `Invoke-SharePointOffboardCleanup -LeaverUPN <upn>` — walks every outbound share the leaver created in the last `LookbackDays` (default 365) and revokes each one. Used by the offboarding 12-step flow.

```
[Cleaning up SharePoint shares created by alice@contoso.com]
[+] Revoked share to external@protonmail.example (Q1-Customers.xlsx)
[+] Revoked share to bob@contoso.com            (proposal.pdf)
...
[Done] 12 share(s) revoked.
```

Each revoke is a standalone `Invoke-Action` with `actionType=RevokeShare`. Reversible via `Invoke-Undo` per entry.

## Stale-site report

**Slot 18 → option 7.** Lists sites whose `LastContentModifiedDate` is older than a threshold (default 365 days). Output sorted by oldest first.

Useful for cleanup before a migration or audit. The tool does NOT auto-delete stale sites — that's a manual decision.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Connect-SPOService: The remote server returned an error: (404)` | Wrong SPO admin URL. Re-run `Set-SPOAdminUrl` or check the cache at `<stateDir>\spo-tenants.json`. |
| `Get-UserOutboundShares returns empty when you know there are shares` | UAL not enabled for this tenant, or the user is outside the search retention window. Verify UAL is on (Compliance Center → Audit). |
| `Add-PnP module not found` | Some flows need the PnP SharePoint module. Install: `Install-Module PnP.PowerShell -Scope CurrentUser`. |

Full troubleshooting at [`../operations/troubleshooting.md`](../operations/troubleshooting.md).

## See also

- [`offboarding.md`](offboarding.md) — step 7 cleans up outbound shares automatically.
- [`incident-response.md`](incident-response.md) — step 10 audits shares + the `Detect-MassExternalShare` auto-trigger.
- [`../reference/template-schema.md`](../reference/template-schema.md) — site template JSON shape.
