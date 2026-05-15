# MFA management

Audit and manage Multi-Factor Authentication methods across users. List who's registered for what, revoke specific methods or all of them, issue Temporary Access Passes for re-enrollment, and run compliance views.

## Menu

**Slot 15 → MFA & Authentication**

```
  1. List a user's MFA methods
  2. Revoke a specific MFA method
  3. Revoke ALL MFA methods (use during incident response)
  4. Issue Temporary Access Pass (TAP)
  5. Compliance views (no-MFA / phone-only / TAP-active / FIDO2 users)
  6. CSV export
  7. Bulk MFA reset (CSV-driven)
```

## Prereqs

- Graph scopes: `UserAuthenticationMethod.ReadWrite.All`, `UserAuthenticationMethod.Read.All`.
- Role: `Authentication Administrator` or higher.

## Method types

The tool recognizes seven authentication method types:

| Type | Graph URL segment | Notes |
|---|---|---|
| Microsoft Authenticator | `microsoftAuthenticatorMethods` | The app. Most common. |
| Phone (SMS / voice) | `phoneMethods` | Phishable. Surface in compliance views. |
| FIDO2 key | `fido2Methods` | Phishing-resistant. Prefer for sensitive accounts. |
| Windows Hello for Business | `windowsHelloForBusinessMethods` | Device-bound. |
| TAP | `temporaryAccessPassMethods` | Single-use; for re-enrollment. |
| Email (OTP) | `emailMethods` | Self-service password reset backup. |
| Password | (no URL segment — cannot be revoked via Graph) | Surfaced for completeness. |

## Listing methods

```powershell
Get-UserAuthMethods -User alice@contoso.com | Format-Table

# Output:
# Label                       Id          UrlSegment
# -----                       --          ----------
# Microsoft Authenticator     auth-001    microsoftAuthenticatorMethods
# Phone (SMS/voice)           phone-002   phoneMethods
# FIDO2 key                   fido-003    fido2Methods
```

## Revoking a specific method

**Slot 15 → option 2.** Pick the user, pick the method by index.

```
Revoke MFA method 'Microsoft Authenticator' (auth-001) for alice@contoso.com? [Y/N]: Y

[Audit] Action: RevokeAuthMethod  result: success
[Audit] noUndoReason: Auth method revocation cannot be undone via API; the user must re-register the method.
```

The audit entry's `noUndoReason` is explicit: there's no curated reverse recipe. Recovery is operator-driven re-enrollment.

## Revoking ALL methods

**Slot 15 → option 3.** Used during incident response (see [`incident-response.md`](incident-response.md) step 4). Walks every registered method and revokes each individually. The snapshot taken at the start of the incident playbook preserves the original method list so the operator can verify what was revoked.

This step is irreversible. The user CANNOT sign in until they re-enroll — either at password reset (which prompts for MFA setup) or via a Temporary Access Pass.

## Issuing a Temporary Access Pass

**Slot 15 → option 4.** A TAP is a single-use one-hour code that satisfies MFA for the next sign-in. Used for:

- New-hire first sign-in (paired with the onboarding flow).
- Compromised-account recovery (paired with the incident playbook).
- "I lost my phone and my hardware key is in the office" scenarios.

```
> UPN              : alice@contoso.com
> Lifetime minutes : 60                  (default 60; max 480)
> Single-use       : Y                   (default Y; N for "valid for any sign-in in the window")

[+] TAP issued
    Code: XXXX-XXXX-XXXX
    Lifetime: 60 minutes
    Single-use: True

  Deliver via phone or in-person -- not email -- and remind the user to register a permanent method
  immediately after signing in.
```

The TAP code is shown ONCE. The tool does NOT log it (audit entry confirms TAP was issued but not the value).

## Compliance views

**Slot 15 → option 5.**

```
  1. Users with NO MFA registered
  2. Users with ONLY phone-based MFA (phishable)
  3. Users with an active TAP
  4. Users registered for FIDO2 (the gold standard)
  5. Users with self-service password reset backup configured
```

Each option prints a table + offers CSV export. Pattern: enumerate users via Graph, fetch each user's methods, filter, output.

For very large tenants the enumeration is paginated (500 users per page, configurable via `Get-UsersWithNoMfa -Max <n>`).

## CSV export

**Slot 15 → option 6** writes `mfa-methods-<ts>.csv` with one row per (user, method) pair:

```csv
UPN,DisplayName,MethodType,MethodId,Registered,IsDefault
alice@contoso.com,Alice Smith,Microsoft Authenticator,auth-001,2025-01-15,true
alice@contoso.com,Alice Smith,Phone (SMS/voice),phone-002,2025-01-15,false
```

Useful for tracking re-enrollment campaigns or for compliance audit evidence.

## Bulk reset from CSV

**Slot 15 → option 7.** For "we got phished as a team" scenarios where you need to reset MFA on multiple users without doing a full incident response per user.

CSV columns: `UPN, IssueTAP, TAPLifetimeMinutes, Reason`.

```csv
UPN,IssueTAP,TAPLifetimeMinutes,Reason
alice@contoso.com,true,60,"Team-wide phishing 2026-05-14"
bob@contoso.com,true,60,"Team-wide phishing 2026-05-14"
```

The flow:

1. Validate the CSV.
2. PREVIEW pass: shows what would happen.
3. Confirm.
4. LIVE pass: per-user, revoke all methods, optionally issue TAP, audit each step.

For a full compromised-account response (block sign-in, revoke sessions, force password change, audit activity), prefer [`incident-response.md`](incident-response.md)'s `Invoke-BulkIncidentResponse` instead.

## Common failures

| Symptom | Cause + fix |
|---|---|
| `Insufficient privileges to read this user's authentication methods` | Account lacks `UserAuthenticationMethod.Read.All` consent. Re-consent the Graph app or use an account with `Authentication Administrator`. |
| `Cannot revoke the Password method` | Passwords aren't authentication methods — they're set via `passwordProfile`. Use the password reset flow instead. |
| `User has no methods to revoke` | The user has only a password registered, OR you're looking at a Guest user whose auth lives in their home tenant. |
| `TAP issuance failed: User is excluded from policy` | Your tenant's TAP policy excludes this user (often: Global Admins). Adjust the policy in Entra → Authentication methods → TAP. |

Full troubleshooting at [`../operations/troubleshooting.md`](../operations/troubleshooting.md).

## See also

- [`incident-response.md`](incident-response.md) — the playbook that uses these primitives at scale.
- [`audit-and-undo.md`](audit-and-undo.md) — viewing MFA revoke audit entries.
- [`../operations/permissions.md`](../operations/permissions.md) — Graph scopes per feature.
