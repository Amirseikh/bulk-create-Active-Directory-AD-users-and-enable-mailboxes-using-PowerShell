# AD & Exchange Bulk User Provisioning Toolkit

A PowerShell/WPF GUI tool for bulk-creating Active Directory users from an
Excel/CSV sheet and enabling their Exchange mailboxes — built for large,
multi-domain-controller environments where an admin team hands off a
spreadsheet and expects consistent, auditable provisioning.

## Why this exists

Most bulk-AD-creation scripts assume a single simple domain and skip mailbox
provisioning entirely. This tool was built for a real multi-domain-controller
environment with a Global Catalog, mailbox database capacity limits, and an
Exchange environment with restricted RBAC — and handles the edge cases that
only show up at that scale (a few hundred users, multiple DCs, multiple
mailbox databases).

## What's included

| File | Purpose |
|---|---|
| `ADUserProvisioning.ps1` | The tool — WPF GUI: configure once, browse to a sheet, dry-run, create |

## Features

- **Configurable mapping** from a spreadsheet to AD attributes (first/second/third/last
  name composition into `givenName`, `sn`, `cn`, `displayName`, `name`)
- **Dry run mode** — validate everything (duplicate checks, OU reachability,
  mailbox database placement) with zero writes to AD or Exchange
- **Duplicate detection** across multiple search DNs and multiple LDAP filter
  templates, with automatic fallback (Global Catalog → plain DC → auto-locate)
  and a fail-safe: if duplicate-checking can't be verified, the tool refuses
  to create rather than risk a silent duplicate
- **Random 12-character passwords** (guaranteed upper/lower/digit/special),
  `ChangePasswordAtLogon` set automatically
- **Mailbox database capacity + distribution** — cap mailboxes per database,
  choose Fill-sequentially or Round-robin. By default the cap applies only
  to mailboxes created *this run* — it is not reduced by whatever already
  exists in the database, so requesting 100 always creates 100. Check
  "Count pre-existing mailboxes..." in the UI to instead treat the cap as a
  true total-capacity ceiling (existing + new)
- **Replication-safe mailbox enablement** — pins Exchange's AD read/write to
  the same domain controller the account was just created on, with a short
  backoff-retry. Without this, a fast bulk batch can hit "object couldn't be
  found on &lt;DC&gt;" when Exchange picks a DC that hasn't replicated the
  brand-new account yet
- **Exchange mailbox enablement via implicit remoting** (`Import-PSSession`),
  not `Invoke-Command` — Exchange's remote runspace runs in a restricted
  "no-language mode" that rejects scriptblocks with `param()`/`if`/`$using:`,
  a subtlety that breaks most naive implementations
- **RBAC pre-flight check** — verifies the Exchange account can actually run
  `Enable-Mailbox` at connection time, not silently mid-batch
- **DPAPI-encrypted credential storage** (Windows-user + machine bound, never
  plaintext) for the AD provisioning account
- **Per-run transcript logging** and CSV export of results

## Requirements

- Windows PowerShell 5.1, .NET Framework 4.7.2+
- RSAT `ActiveDirectory` module (`Install-WindowsFeature RSAT-AD-PowerShell`
  on Server, or `Add-WindowsCapability` on Windows 10/11)
- Domain-joined workstation, network path to a DC on **TCP 9389** (ADWS —
  required by every AD cmdlet, not just LDAP 389) and **TCP 3268** (Global
  Catalog, if used)
- `ImportExcel` PowerShell module for `.xlsx` input (auto-installs; falls
  back to Excel COM, or just use CSV)
- Exchange: reachable on TCP 80 (or 443), and the connecting account needs
  the **Recipient Management** / **Mail Recipients** RBAC role to run
  `Enable-Mailbox`

## Quick start

```powershell
powershell -ExecutionPolicy Bypass -File .\ADUserProvisioning.ps1
```

In the GUI: fill in Forest/Domain/Exchange/Search-scope configuration once
(it's saved to a local, gitignored `ADProvisioning.config.json`), browse to
your import file, **run a dry run first**, then go live.

## Import file format

| Excel Column Header | AD Attribute |
|---|---|
| `employeeID` | employeeID |
| `employeeNumber` | employeeNumber |
| `First name` | givenName |
| `Second name` + `Third name` + `Last name` | sn |
| `First name` + `Second name` + `Third name` + `Last name` | cn, displayName, name |
| `mobile` | mobile |
| `sAMAccountName` | sAMAccountName |
| `mail` | mail, UserPrincipalName |
| `title` | title |
| `occupation`, `cluster`, `clusterCode`, `facilty`, `faciltyCode`, `employeeType` | schema-extended custom attributes (adjust to your own schema) |

Column order doesn't matter — matching is by exact header name.

## Security notes

- **Never commit** `ADProvisioning.config.json`, anything under `Logs/`, or
  any `Results_*.csv` — they contain generated passwords and/or infrastructure
  details. The included `.gitignore` excludes them by default.
- The stored AD provisioning password is DPAPI-encrypted and only decryptable
  by the same Windows user on the same machine that saved it.
- Exported result CSVs contain initial plaintext passwords for handover —
  treat them as sensitive and delete after distribution.

## License

MIT — see `LICENSE`.
