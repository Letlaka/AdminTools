# Get-ADAdminActivity.ps1 Manual

`Get-ADAdminActivity.ps1` reads Domain Controller Security logs and exports a
CSV report of Active Directory administrative activity such as account changes,
group membership changes, computer changes, and policy updates.

## Setup

Run this script from Windows PowerShell 5.1 or later on a domain-joined admin
workstation or management server.

Required components:

- Windows PowerShell 5.1+
- RSAT Active Directory module
- Permission to query AD
- Permission to read Security logs on Domain Controllers
- Advanced Audit Policy configured to record account management events

Recommended first-time check:

```powershell
powershell
Get-Module -ListAvailable ActiveDirectory
Get-ADDomainController -Filter *
```

## Quick Start

Export the last seven days of administrative activity:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

Filter to privileged admin activity:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 30 -AdminOnly
```

Write to a specific CSV path:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 7 `
  -OutputCsv ".\reports\ad-admin-activity\AD_Admin_Activity_Custom.csv"
```

## High-Level Flow

1. Read command-line parameters.
2. Validate output path, Domain Controller names, group names, and admin names.
3. Load the trusted RSAT `ActiveDirectory` module.
4. Discover writable Domain Controllers unless `-DomainControllers` is supplied.
5. Resolve privileged admin membership when `-AdminOnly` is used.
6. Query Security logs for supported AD administrative event IDs.
7. Parse event properties into normalized report rows.
8. Apply optional privileged-admin filtering.
9. Export sanitized CSV output.

## Usage

Common query controls:

- `-DaysBack`
- `-DomainControllers`
- `-Credential`
- `-CredentialSecretName` or `-CredentialPath` for reusable credentials
- `-MaxEventsPerDomainController`
- `-AllowPartialResults`
- `-AllowUnverifiedDomainController`

Admin filtering controls:

- `-AdminOnly`
- `-AdminSamAccountNames`
- `-PrivilegedGroupNames`

Output controls:

- `-OutputCsv`
- `-LogPath`
- `-IncludeMessage`
- `-MaxAttributeValueLength`
- `-NoClobber`
- `-ForceOverwrite`
- `-AllowNetworkOutputPath`
- `-DisableCsvSanitization`

## JSON Config Support

This script does not currently support `-ConfigPath` or JSON configuration
files. Use command-line parameters or a wrapper script if you need repeatable
presets.

Example wrapper variables:

```powershell
$reportPath = ".\reports\ad-admin-activity\AD_Admin_Activity_Custom.csv"
$domainControllers = @("dc01.domain.local", "dc02.domain.local")

.\Get-ADAdminActivity.ps1 `
  -DaysBack 14 `
  -DomainControllers $domainControllers `
  -AdminOnly `
  -OutputCsv $reportPath `
  -AllowPartialResults
```

## Examples

Privileged admin activity using default privileged groups:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly
```

Privileged admin activity with additional admin names:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly `
  -AdminSamAccountNames "svc-ad-admin","tier0.operator"
```

Query specific Domain Controllers:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 7 `
  -DomainControllers "dc01.domain.local","dc02.domain.local" `
  -AllowPartialResults
```

Include rendered event messages:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 3 `
  -IncludeMessage `
  -OutputCsv ".\reports\ad-admin-activity\AD_Admin_Activity_WithMessages.csv"
```

## Outputs

Default output location:

```text
reports\ad-admin-activity
```

Default file pattern:

```text
AD_Admin_Activity_Report_<timestamp>.csv
```

Default run log location:

```text
logs\get-ad-admin-activity\Get-ADAdminActivity_<timestamp>.log
```

Common fields:

- `TimeCreated`
- `DomainController`
- `EventId`
- `Action`
- `ActorAccount`
- `ActorSamAccountName`
- `ActorSid`
- `TargetObject`
- `TargetAccount`
- `MemberName`
- `AttributeName`
- `OperationType`
- `AttributeValue`
- `EventRecordId`
- `RenderedMessage`

## Safety Notes

- CSV values are sanitized by default.
- Existing CSV and log files are not overwritten unless `-ForceOverwrite` is supplied.
- UNC output and custom log paths require `-AllowNetworkOutputPath`.
- UNC credential paths require `-AllowNetworkInputPath`; credential files must be outside the repository directory.
- Supplied Domain Controllers are verified through AD unless `-AllowUnverifiedDomainController` is supplied.
- Use `-IncludeMessage` carefully because rendered event text can contain sensitive values.
