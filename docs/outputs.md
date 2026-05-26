# Outputs and Report Files

## Scan-ADComputers.ps1

`Scan-ADComputers.ps1` can export:

- Inventory data
- Targeted audit data
- Separate targeted status reports
- Summary reports
- Delta reports

Supported formats:

- `Csv`
- `Json`
- `Html`

CSV exports are sanitized by default to reduce spreadsheet formula injection risk.
Use `-DisableCsvSanitization` only when raw values are required and the output
will not be opened in spreadsheet software.

Network output paths are rejected by default unless `-AllowNetworkOutputPath` is
supplied. Existing files are not overwritten unless `-ForceOverwrite` is supplied.

## User Account Management Reports

`Manage-ADUserAccounts.ps1` writes reports to
`%LOCALAPPDATA%\AdminTools\ADUserAccountReports` by default, or to
`-OutputDirectory` when supplied.

Network output paths are rejected by default unless `-AllowNetworkOutputPath` is
supplied. Existing files are not overwritten unless `-ForceOverwrite` is supplied.
User list files from UNC paths are rejected unless `-AllowNetworkInputPath` is supplied.

File name pattern:

- `ADUsers_<ReportName>_<domain>_<timestamp>.<ext>`

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

Common user fields:

- `SamAccountName`
- `UserPrincipalName`
- `DisplayName`
- `Enabled`
- `LockedOut`
- `Created`
- `AccountAgeDays`
- `PasswordLastSet`
- `PasswordAgeDays`
- `PasswordAgeStatus`
- `PasswordNeverExpires`
- `CannotChangePassword`
- `LastLogonDate`
- `DaysSinceLastLogon`
- `LastBadPasswordAttempt`
- `AccountExpirationDate`
- `AdminCount`
- `DistinguishedName`
- `ObjectSID`
- `ObjectGUID`
- `QueriedAt`

## AD Admin Activity Reports

`Get-ADAdminActivity.ps1` writes a CSV report to `-OutputCsv`, or to
`%LOCALAPPDATA%\AdminTools\ADAdminActivityReports` by default.

File name pattern:

- `AD_Admin_Activity_Report_<timestamp>.csv`

Common fields:

- `TimeCreated`
- `DomainController`
- `EventId`
- `Action`
- `ActorAccount`
- `ActorSamAccountName`
- `ActorSid`
- `SubjectLogonId`
- `TargetObject`
- `TargetAccount`
- `MemberName`
- `ObjectDistinguishedName`
- `AttributeName`
- `OperationType`
- `AttributeValue`
- `EventRecordId`
- `RenderedMessage`

## Script Coverage

For setup, quick start, usage, examples, and safety notes for each script, see:

- [Scan-ADComputers Manual](scan-adcomputers.md)
- [Get-ADAdminActivity Manual](ad-admin-activity.md)
- [Manage-ADUserAccounts Manual](user-account-management.md)

## Computer Inventory File Name Patterns

- Main inventory: `<Servers|Workstations>_<domain>_<timestamp>.<ext>`
- Targeted audit: `<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit.<ext>`
- Separate targeted status exports (when enabled):
  - `_Matched`
  - `_Unreachable`
  - `_NotFoundInAD`

## Main Inventory Fields

- `ComputerType`
- `Name`
- `CN`
- `DNSHostName`
- `Description`
- `OUPath`
- `CanonicalName`
- `DistinguishedName`
- `Created`
- `Enabled`
- `IPv4Address`
- `LastLogonDate`
- `LastLogonTimestamp`
- `LastSeenDate`
- `DaysSinceLastSeen`
- `InactiveThresholdDays`
- `IsStale`
- `StaleStatus`
- `ObjectGUID`
- `OperatingSystem`
- `OperatingSystemVersion`
- `ConnectivityMethod`
- `ConnectivityStatus`
- `ConnectivityReachable`
- `ConnectivityDetail`
- `DnsStatus`
- `DnsResolvedIPs`
- `DnsMatchesAdIPv4`
- `PortStatus`
- `RemoteInventoryStatus`
- `RemoteUptimeDays`
- `SerialNumber`
- `Model`
- `TotalMemoryGB`
- `SystemDriveFreeGB`
- `PendingReboot`
- `LoggedOnUser`
- `QueriedAt`

## Targeted List File Format

The targeted input file is plain text:

- one host per line
- short name or FQDN under the discovered or supplied AD DNS suffix
- blank lines allowed
- comment lines starting with `#` are ignored
- names with invalid DNS characters are rejected

Example:

```text
# Core production servers
Server01
sql01.domain.local
Server02
```
