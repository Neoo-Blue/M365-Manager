# Sign-in lookup

Search Entra sign-in logs with operator-friendly filters. Built on Graph `/auditLogs/signIns`.

## Menu

**Slot 14 → Audit & Reporting → Sign-in lookup**

Or directly:

```powershell
Search-SignIns -User alice@contoso.com -From (Get-Date).AddDays(-7) -OnlyFailures
```

## Filter parameters

| Parameter | Notes |
|---|---|
| `-User` | UPN. Empty = all users. |
| `-From` / `-To` | `[DateTime]` UTC. Default is last 24h. |
| `-AppName` | Wildcard match against application display name (`Office 365 Exchange Online`, `Microsoft Graph`, etc.). |
| `-IP` | Source IP address (exact or CIDR — exact only today). |
| `-Country` | ISO country code (`US`, `GB`) or full name (`United States`). |
| `-RiskLevel` | `none` / `low` / `medium` / `high`. From Identity Protection. |
| `-OnlyFailures` | Switch — restrict to non-success status codes. |
| `-MaxResults` | Default 200. Increase for wide-window searches. |

## Common queries

### Recent failures for a user

```powershell
Search-SignIns -User alice@contoso.com -OnlyFailures -From (Get-Date).AddDays(-2)
```

### All high-risk sign-ins in the tenant

```powershell
Search-SignIns -RiskLevel high -From (Get-Date).AddDays(-7) -MaxResults 500
```

### Sign-ins from an unexpected country

```powershell
Search-SignIns -Country Nigeria -From (Get-Date).AddDays(-30)
```

### A specific app's failures

```powershell
Search-SignIns -AppName "Microsoft Authenticator" -OnlyFailures
```

## Output shape

Each row is a `PSCustomObject` with:

| Property | Meaning |
|---|---|
| `CreatedDateTime` | UTC timestamp. |
| `UserDisplayName` / `UserPrincipalName` | The user. |
| `AppDisplayName` | Application that initiated the auth. |
| `IpAddress` | Source. |
| `Location` | Hashtable `{city, state, countryOrRegion}` or string fallback. |
| `Status` | Human-readable status (`Success`, `Authentication cancelled by the user`, etc.). |
| `StatusCode` | Numeric Entra code. |
| `ConditionalAccessStatus` | If a CA policy fired. |
| `RiskLevel` | `none` / `low` / `medium` / `high`. From Identity Protection. |
| `MfaDetail` | MFA method actually used (if any). |

## Status code reference

A few codes worth knowing:

| Code | Meaning |
|---|---|
| `0` | Success. |
| `50053` | Account locked due to lockout policy. |
| `50074` | Strong auth required but not satisfied. |
| `50076` | MFA required by CA policy. |
| `50126` | Bad password. |
| `50158` | MFA challenge cancelled by user. |
| `500121` | MFA challenge timeout. |
| `530002` | Conditional Access blocked the sign-in. |

The full list: https://learn.microsoft.com/en-us/azure/active-directory/develop/reference-error-codes.

## Used by

- [`incident-response.md`](incident-response.md) step 1 (snapshot recentSignIns), step 8 (Audit24h).
- [`../playbooks/incident-triggers.md`](../playbooks/incident-triggers.md) — three detectors (`AnomalousLocationSignIn`, `ImpossibleTravel`, `HighRiskSignIn`) all read from `Search-SignIns`.
- `Detect-MFAFatigue` — looks for rejected MFA codes via specific failure status patterns.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Returns 0 rows for a user who you know signed in` | `signInActivity` in Graph is updated with a lag (sometimes hours). For real-time data, use Identity Protection events instead. |
| `Insufficient privileges` | Account needs `AuditLog.Read.All` Graph scope. Re-consent or add the `Security Reader` role. |
| `Identity Protection not configured for this tenant` | Risk-level filtering returns null for `RiskLevel` on every row. The tenant needs Entra P1 / P2 for risk events. |
| Wide-window queries time out | Reduce window or paginate. Graph caps results per page; the tool pages internally but very wide windows still hit timeouts. |

## See also

- [`unified-audit-log.md`](unified-audit-log.md) — for resource-level events (file access, mail sends, etc.) not just sign-ins.
- [`incident-response.md`](incident-response.md) — the playbook's snapshot + audit steps consume this.
- [`../operations/permissions.md`](../operations/permissions.md) — required Graph scopes.
