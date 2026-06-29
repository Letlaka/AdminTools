# Outputs and Report Files

## Scan-ADComputers.ps1

`Scan-ADComputers.ps1` writes reports to `reports/ad-computers` by default, or to `-OutputDirectory` when supplied. The default run log is written under `logs/scan-ad-computers/`.

`Scan-ADComputers.ps1` can export:

- Inventory data
- Targeted audit data
- Separate targeted status reports
- Summary reports
- Delta reports
- Optional performance summaries when `-PerformanceSummary` is supplied

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
`reports/ad-user-accounts` by default, or to
`-OutputDirectory` when supplied. The default run log is written under
`logs/manage-ad-user-accounts/`.

Network output paths are rejected by default unless `-AllowNetworkOutputPath` is
supplied. Existing report and log files are not overwritten unless `-ForceOverwrite` is supplied.
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
`reports/ad-admin-activity` by default. The default run log is written under
`logs/get-ad-admin-activity/`.

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


## AD Excel Reporting Outputs

`scripts/build_ad_excel_reports.py` reads `Scan-ADComputers.ps1` CSV or JSON exports and writes Excel reporting output under `reports/<financial-year>/<run-date>/` by default. Logs are written under `logs/excel-reporting/<financial-year>/<run-date>/` by default.

Output folders:

- `source`: copies of the input scan exports
- `consolidated`: one `AD_Dashboard_<financial-year>_<run-date>.xlsx` workbook
- `departments`: one workbook folder per matched department
- `logs/excel-reporting/<financial-year>/<run-date>/unmatched_devices.csv`: records that did not match a department

The reporting utility requires environment-local department runtime files:

- `config/dept_list.txt` copied from `config/dept_list.sample.txt`
- `config/dept_codes.txt` copied from `config/dept_codes.sample.txt`

The runtime files are ignored by git so public clones get the samples without receiving environment-specific department data.

Excel detail sheets add normalized reporting fields near the front of each device record:

- `LastSeenDate`: preserves a valid source `LastSeenDate` when present, otherwise uses source `LastLogonDate`, with `LastLogonTimestamp` as a final fallback.
- `InactivityDays`: source `DaysSinceLastSeen` when present, otherwise calculated from the normalized `LastSeenDate`.
- `InactivityStatus`: `Stale` at 90 days or more, `Fresh` below 90 days, or `Unknown` when no usable source exists. Source `StaleStatus = Stale` takes priority for this status.

Raw scan fields such as `DaysSinceLastSeen`, `StaleStatus`, and `IsStale` remain in the detail sheets later in the row for traceability.

## Script Coverage

For setup, quick start, usage, examples, and safety notes for each script, see:

- [Scan-ADComputers Manual](scan-adcomputers.md)
- [Get-ADAdminActivity Manual](ad-admin-activity.md)
- [Manage-ADUserAccounts Manual](user-account-management.md)

## Computer Inventory File Name Patterns

- Main inventory: `<Servers|Workstations>_<domain>_<timestamp>.<ext>`
- Performance summary: `<Servers|Workstations>_<domain>_<timestamp>_Performance.csv` and `.json`
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
