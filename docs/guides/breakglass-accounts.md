# Break-glass account best practices

Adapted from Microsoft's [emergency access account guidance](https://learn.microsoft.com/azure/active-directory/roles/security-emergency-access). Each line maps to a check that `BreakGlass.ps1`'s `Test-BreakGlassPosture` runs.

| Microsoft recommendation | Module check | Default |
|---|---|---|
| Account is cloud-only (not synced from on-prem). | (manual) | n/a |
| Permanent Global Administrator (PIM-eligible isn't enough). | (manual) | n/a |
| Strong, unique credential — FIDO2 security key recommended. | `noFido2` warning when no FIDO2 method is registered. | warn |
| MFA registered. | `noMfaRegistered` warning when no strong methods are registered. | warn |
| Excluded from every conditional-access policy that could block sign-in (MFA, legacy auth, device compliance, etc.). | `caRiskyInclude` warning when an enabled CA policy with MFA/block grant controls includes this account (or "All users") without excluding it. | warn |
| Account is enabled. | `accountDisabled` warning when `accountEnabled = false`. | warn |
| Account is rarely used — sign-ins should be near-zero outside genuine outages. | `recentSignIn` warning when last sign-in is within 30 days. | warn |
| Password rotated at least annually (often quarterly). | `passwordAge` warning when `lastPasswordChangeDateTime` is older than 180 days (`$script:BGPasswordAgeWarnDays`). | warn |

## Recommended cadence

| Cadence | Action |
|---|---|
| Daily | `health-breakglass-signins.ps1` (alerts on any sign-in within 24h). |
| Monthly | `Test-AllBreakGlassPosture` to surface any drift. |
| Quarterly | `Invoke-QuarterlyBreakGlassAttestation` — emails each account's attestation contact, records `LastAttestedAt`. Sign-off is captured manually (reply YES). |
| Annually | Rotate the password / replace the FIDO2 key, even if posture is clean. |

## How to register an account

```powershell
# Interactive (Audit & Reporting -> Break-glass accounts -> Register)
Register-BreakGlassAccount -UPN 'breakglass-01@contoso.com' -AttestationEmail 'security-team@contoso.com'
```

The registry lives at `<stateDir>/breakglass-accounts.json`. The state file is plain JSON — there's nothing secret in it (just UPNs and metadata), but the `<stateDir>` itself is mode `0700` on POSIX and inherits user-only NTFS ACLs from `%LOCALAPPDATA%` on Windows.

## Quarterly attestation

`Invoke-QuarterlyBreakGlassAttestation` walks the registry, runs the posture check on each account, builds an HTML attestation email with any warnings inline, and emails the configured `AttestationEmail`. It stamps `LastAttestedAt` with the send timestamp and `LastAttestedBy = 'pending-reply'`. When the recipient replies "ATTESTED", the operator updates `LastAttestedBy` to that contact's UPN manually — this commit deliberately doesn't wire up a mailbox-rule callback because the security review surface should stay simple.
