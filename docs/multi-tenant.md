# Multi-tenant / MSP mode (Phase 6)

Phase 6 turns the tool from "single-tenant operator console" into
"MSP cockpit". An operator can register many tenants, switch
between them with a single command, run reports across the whole
portfolio in one pass, and view a single HTML dashboard summarizing
everyone's posture.

## The trust model

There is no central server. Every tenant profile, every credential,
every audit log lives on the operator's local box under `<stateDir>`
(`%LOCALAPPDATA%\M365Manager` on Windows, `~/.m365manager` on POSIX,
chmod 700).

- `tenants.json` (metadata) is **plaintext**. Safe to check into a
  config-management repo if you scrub `tenantId` first.
- `secrets/tenant-<name>.dat` (credential manifest) is encrypted via
  the same `Protect-Secret` / DPAPI path the AI API key uses
  (Windows-only true crypto; POSIX falls back to base64 with a
  warning -- combine with full-disk encryption for confidentiality).
- The cert thumbprint mode stores only the thumbprint, not the
  private key. The private key lives in the OS cert store and the
  Connect-MgGraph SDK pulls it from there at connect time. This is
  the recommended mode for production MSPs.
- Client-secret mode is supported but warned -- rotate frequently
  and prefer cert+thumbprint for app-only auth.

The audit log is per-tenant from Phase 6 onward: every entry now
carries a structured `tenant: { name; id; domain; mode }` field, and
the log filename includes a tenant-name slug. Switching tenants
opens a fresh log so the operator can grep one customer cleanly.

## Profile schema

```jsonc
{
  "name":          "Contoso",
  "tenantId":      "abcd1234-...",
  "primaryDomain": "contoso.onmicrosoft.com",
  "spoAdminUrl":   "https://contoso-admin.sharepoint.com",
  "credentialRef": "tenant-contoso",   // -> secrets/<credentialRef>.dat
  "tags":          ["customer","prod"],
  "lastUsed":      "2026-05-12T18:33:00Z",
  "notes":         "Primary IT contact: bob@contoso.com",
  "createdUtc":    "2026-05-12T18:30:00Z"
}
```

The credential manifest sidecar:

```jsonc
{
  "schemaVersion": 1,
  "authMode":      "CertThumbprint",   // or "ClientSecret" / "Interactive"
  "clientId":      "5e7d...",
  "thumbprint":    "9F8E...",
  "secret":        null                 // populated only for ClientSecret mode
}
```

## CRUD via PowerShell

```powershell
Register-Tenant -Name 'Contoso' -TenantId 'abcd1234-...' `
    -PrimaryDomain 'contoso.onmicrosoft.com' `
    -AuthMode 'CertThumbprint' -ClientId '...' -CertThumbprint '...' `
    -Tags @('customer','prod')

Update-Tenant -Name 'Contoso' -CertThumbprint '<rotated>'

Remove-Tenant -Name 'Contoso'

Get-Tenants | Format-Table name, tenantId, lastUsed
Get-Tenant  -Name 'Contoso'
Show-TenantRegistry
```

## Switching

```powershell
Switch-Tenant -Name 'Contoso'              # PowerShell
/tenant Contoso                            # from the AI chat
/tenants                                   # list from the AI chat
```

Switch-Tenant:
1. Audits the intent under the *prior* tenant's log.
2. Disconnects Graph / EXO / SCC / SPO.
3. Calls `Reset-AuditLogPath` so subsequent entries land in a new
   per-tenant file.
4. Sets `$script:CurrentTenantProfile` and mirrors fields into
   `$script:SessionState` for backward compatibility.
5. For app-only profiles, eagerly reconnects via cert thumbprint or
   client secret. Interactive profiles defer to the next service
   call so the browser opens only when needed.
6. Paints the per-tenant colored banner so the next operation is
   visually unambiguous.

## Cross-tenant operations

`Invoke-AcrossTenants -Tenants @('Acme','Contoso') -Script { ... }`
runs a scriptblock per tenant, returning a uniform shape:

```
@(
  @{ Tenant='Acme';    Success=$true;  Result=...; Error=$null; DurationMs=420 }
  @{ Tenant='Contoso'; Success=$true;  Result=...; Error=$null; DurationMs=380 }
)
```

Pre-built rollups:
- `Get-CrossTenantLicenseUtilization`
- `Get-CrossTenantMFAGaps`
- `Get-CrossTenantStaleGuests -DaysSinceSignIn 90`
- `Get-CrossTenantOrphanedTeams`
- `Get-CrossTenantBreakGlassPosture`

Each returns `@{ Tenants=<rows>; Summary=<rollup> }` so they all
feed the dashboard uniformly.

**Sequential by design.** The Graph / EXO SDK keeps connection state
in process-global singletons. Two PowerShell runspaces hitting
different tenants in parallel would race against the same singleton
and produce data attributed to the wrong tenant. `-Parallel` is
accepted for API symmetry but emits a warning and falls through to
sequential execution.

## First-run migration

If you upgrade from a Phase 5 build, the registry starts empty.
On first run with an active interactive tenant context, the tool
offers to register that tenant as a profile so future `/tenant`
switches don't re-walk the partner-center picker. Skip the prompt
to defer; you can always `Register-Tenant` manually later.

## Token refresh + long-running sessions

The Graph SDK auto-refreshes tokens. Before Phase 6 a long-running
session that had switched tenants risked the refresh hitting the
*wrong* tenant. Now Switch-Tenant calls Reset-AllSessions before
Set-CurrentTenant, which forces a fresh consent on the new tenant
rather than refreshing the previous tenant's token. Operators in a
multi-hour session should still `/tenants` and Show-TenantRegistry
before any destructive operation -- visual verification beats trust.

## See also

- [tenant-overrides.md](tenant-overrides.md) -- per-tenant config keys.
- [msp-dashboard.md](msp-dashboard.md) -- HTML portfolio overview.
- [audit-format.md](audit-format.md) -- structured tenant field details.
