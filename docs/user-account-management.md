# Manage-ADUserAccounts.ps1 Manual

`Manage-ADUserAccounts.ps1` provides Active Directory user account reports,
single-user audit lookup, locked-out account detail, and explicit account reset
actions.

Reporting modes are read-only. Reset mode supports PowerShell `ShouldProcess`,
so you can use `-WhatIf` before making account changes.

## Setup

Run this script from Windows PowerShell 5.1 or later on a domain-joined admin
workstation or management server.

Required components:

- Windows PowerShell 5.1+
- RSAT Active Directory module
- Permission to query AD user accounts
- Permission to read Domain Controller Security logs when event reports are requested
- Delegated rights for unlock, enable, password reset, or password-at-logon actions

Recommended first-time check:

```powershell
powershell
Get-Module -ListAvailable ActiveDirectory
Get-ADUser -Filter * -ResultSetSize 1
```

## Quick Start

Password age report:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType PasswordAge `
  -ExportFormat Csv,Html
```

Locked-out account report:

```powershell
.\Manage-ADUserAccounts.ps1 -Mode LockedOut -ExportFormat Csv,Html
```

Single-user audit summary:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode UserAudit `
  -Identity jsmith `
  -ExportFormat Csv,Json
```

Preview an unlock action:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -Unlock `
  -WhatIf
```

## High-Level Flow

1. Read command-line parameters.
2. Validate identities, OU scope, Domain Controller names, output paths, and safety switches.
3. Load the trusted RSAT `ActiveDirectory` module.
4. Discover domain metadata and Domain Controllers as needed.
5. Build scoped user queries from `SearchBase`, `SearchBaseList`, `ExcludeOU`, and `UserListPath`.
6. Run the selected mode: report, user audit, locked-out detail, or reset.
7. Query Security logs only when event details are requested.
8. Export CSV, JSON, and HTML reports.
9. For reset mode, write an action report that records what was attempted without storing generated passwords.

## Modes

- `Report`: generate one or more account reports.
- `UserAudit`: export a summary for one user and, with `-IncludeEvents`, related Security log events.
- `LockedOut`: export currently locked-out users and, with `-IncludeEvents`, lockout event details.
- `Reset`: unlock, enable, reset password, or require password change at next logon for one user.

## Usage

Common report controls:

- `-ReportType UserSummary,PasswordAge,LockedOut,AuditEvents,PrivilegedUsers,DisabledUsers,StaleUsers`
- `-SearchBase`
- `-SearchBaseList`
- `-ExcludeOU`
- `-UserListPath`
- `-IncludeDisabled`
- `-IncludeGroupMembership`
- `-PasswordAgeWarningDays`
- `-StaleUserDays`

Event controls:

- `-IncludeEvents`
- `-DaysBack`
- `-DomainControllers`
- `-MaxEventsPerDomainController`
- `-AllowPartialResults`
- `-AllowUnverifiedDomainController`
- `-IncludeMessage`

Reset controls:

- `-Identity`
- `-Unlock`
- `-Enable`
- `-ResetPassword`
- `-NewPassword`
- `-GenerateTemporaryPassword`
- `-ShowGeneratedPassword`
- `-ChangePasswordAtLogon`
- `-WhatIf`

Output and safety controls:

- `-ExportFormat Csv,Json,Html`
- `-OutputDirectory`
- `-LogPath`
- `-OutputPrefix`
- `-NoClobber`
- `-ForceOverwrite`
- `-AllowNetworkInputPath`
- Credential reuse with `-CredentialSecretName` or `-CredentialPath`; use only one of these or `-Credential`
- `-AllowNetworkOutputPath`
- `-DisableCsvSanitization`

## JSON Config Support

This script does not currently support `-ConfigPath` or JSON configuration
files. Use command-line parameters or a wrapper script for repeatable presets.

Example wrapper variables:

```powershell
$userReportOutput = ".\reports\ad-user-accounts\scoped-users"
$userScopes = @("OU=Users,DC=domain,DC=local", "OU=Admins,DC=domain,DC=local")

.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge,PrivilegedUsers `
  -SearchBaseList $userScopes `
  -OutputDirectory $userReportOutput `
  -ExportFormat Csv,Html `
  -NoClobber
```

## Examples

User summary and password age:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge `
  -PasswordAgeWarningDays 90 `
  -ExportFormat Csv,Html
```

Privileged, disabled, and stale users:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType PrivilegedUsers,DisabledUsers,StaleUsers `
  -SearchBase "OU=Users,DC=domain,DC=local" `
  -IncludeDisabled `
  -ExportFormat Csv
```

Currently locked accounts:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -ExportFormat Csv,Html
```

Locked accounts with recent lockout events:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -IncludeEvents `
  -DaysBack 7 `
  -AllowPartialResults `
  -ExportFormat Csv,Html
```

Single-user audit summary with event details:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode UserAudit `
  -Identity jsmith `
  -IncludeEvents `
  -DaysBack 30 `
  -ExportFormat Csv,Json
```

Reset a password using a generated temporary password:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -ResetPassword `
  -GenerateTemporaryPassword `
  -ShowGeneratedPassword `
  -ChangePasswordAtLogon
```

Reset a password using a supplied secure string:

```powershell
$password = Read-Host "Temporary password" -AsSecureString

.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -ResetPassword `
  -NewPassword $password `
  -ChangePasswordAtLogon
```

Unlock and enable an account:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -Unlock `
  -Enable
```

## Outputs

Default output location:

```text
reports\ad-user-accounts
```

File name pattern:

```text
ADUsers_<ReportName>_<domain>_<timestamp>.<ext>
```

Common report names:

- `UserSummary`
- `PasswordAge`
- `LockedOut`
- `LockedOutEvents`
- `UserAuditSummary`
- `UserAuditEvents`
- `AuditEvents`
- `PrivilegedUsers`
- `DisabledUsers`
- `StaleUsers`
- `ResetActions`

Supported export formats:

- `Csv`
- `Json`
- `Html`

Default run log location:

```text
logs\manage-ad-user-accounts\Manage-ADUserAccounts_<timestamp>.log
```

## Safety Notes

- `Reset` mode supports one `-Identity` at a time.
- Bulk reset from `-UserListPath` is intentionally not enabled.
- `-ResetPassword` requires either `-NewPassword` or `-GenerateTemporaryPassword`.
- Generated temporary passwords require `-ShowGeneratedPassword`, are displayed once in the console, and are not written to report files.
- Network output and custom log paths require `-AllowNetworkOutputPath`.
- Network user list and credential paths require `-AllowNetworkInputPath`.
- `-CredentialPath` files must be outside the repository directory and should be readable only by the account running the script.
- Existing report and log files are not overwritten unless `-ForceOverwrite` is supplied.
- CSV exports are sanitized by default unless `-DisableCsvSanitization` is supplied.
- Domain Controller names and discovered AD DNS roots are validated before event-log queries.
