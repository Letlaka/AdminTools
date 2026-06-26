# Troubleshooting

## `ActiveDirectory` Module Not Found

**Problem**

- The machine does not have RSAT AD tools installed.

**Fix**

- Install RSAT Active Directory tools.
- Confirm `Get-Module -ListAvailable ActiveDirectory` returns a result.

## PowerShell Version Error

**Problem**

- `Scan-ADComputers.ps1` is running under Windows PowerShell 5.1 or an earlier version.

**Fix**

- Run `Scan-ADComputers.ps1` in PowerShell 7+.

For `Manage-ADUserAccounts.ps1` and `Get-ADAdminActivity.ps1`, Windows
PowerShell 5.1 or later is supported.

## No Results Returned

Possible causes:

- wrong `ComputerType`
- narrow `SearchBase`
- too many exclusions in `ExcludeOU`
- targeted names not found in AD
- connectivity filtering removed targeted results

Things to check:

- run with `-TestMethod None` in targeted mode
- inspect the `_TargetedAudit` export
- confirm the input list format
- verify queried OU scopes

For user account reports:

- use `-IncludeDisabled` if disabled users should appear in broad scoped reports
- use `-Mode LockedOut -IncludeEvents` when you need lockout event source details
- confirm `-SearchBase` points at user OUs, not computer-only OUs
- use `-AllowPartialResults` for event reports when one DC is unavailable

## User Audit Events Missing

Possible causes:

- the account cannot read Security logs on Domain Controllers
- the events have aged out of the Security logs
- Advanced Audit Policy is not logging account management events
- lockout events were written to a different writable Domain Controller

Things to check:

- increase `-DaysBack`
- provide `-DomainControllers` explicitly
- run from an elevated shell
- verify event IDs such as `4724`, `4738`, and `4740` exist on the DC Security logs

## Password Reset Fails

Possible causes:

- insufficient delegated rights
- password does not meet domain policy
- the account is protected or managed by another process
- the target account is outside the delegated OU scope

Things to check:

- preview with `-WhatIf`
- use `-GenerateTemporaryPassword` or provide `-NewPassword` as a secure string
- add `-ShowGeneratedPassword` when using `-GenerateTemporaryPassword`
- add `-ChangePasswordAtLogon` for temporary password workflows

## Remote Inventory Fails

Possible causes:

- WinRM not enabled
- firewall blocked
- insufficient credential rights
- CIM permissions denied

Things to check:

- use `-TestMethod WinRM`
- lower `-ThrottleLimit` or `-RemoteInventoryThrottleLimit`
- increase `-TimeoutSeconds`
- run without `-RemoteInventory` first
- add `-PerformanceSummary` and compare the `OperationalConnectivity`, `DnsResolution`, `PortChecks`, and `RemoteInventory` timings

## Remote Inventory Skipped As Untrusted

**Problem**

- `RemoteInventoryStatus` shows `SkippedUntrustedTarget`.

**Fix**

- Confirm the computer object's `DNSHostName` is under the AD DNS suffix.
- Supply the correct `-DomainName` if discovery is not returning the expected suffix.
- Restrict scans with `-SearchBase` or `-SearchBaseList` so only trusted computer OUs are queried.

## Credential Path Rejected

**Problem**

- `CredentialPath` is rejected because it is under the repository or points at a UNC path.

**Fix**

- Store CLIXML credential files outside the repository, such as under `$env:USERPROFILE\.admintools`.
- Use `-AllowNetworkInputPath` only when a UNC credential location is trusted and access-controlled.
- Use only one credential source: `-Credential`, `-CredentialSecretName`, or `-CredentialPath`.

## Output Path Rejected

**Problem**

- A report or log path is rejected because it already exists or is a network path.

**Fix**

- Use a new output path, or pass `-ForceOverwrite` when replacement is intentional.
- Use `-AllowNetworkOutputPath` only when the UNC location is trusted and access-controlled.
- Use `-AllowNetworkInputPath` only when an input file UNC location is trusted and access-controlled.

## Compare Report Fails

Possible causes:

- previous file path is wrong
- previous file format is not CSV or JSON
- previous file schema is incompatible with current data

## Recommended Starting Commands

Basic full server inventory:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

Basic targeted server scan:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Targeted -ComputerListPath ".\serverlist.txt"
```

Targeted scan with audit and diagnostics:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -ResolveDns `
  -TestPorts 445,3389,5985 `
  -SeparateStatusExports
```

Workstation stale-device summary:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90 `
  -SummaryOnly `
  -ExportFormat Csv,Html
```

Basic admin activity audit:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

Privileged admin audit:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly `
  -AllowPartialResults
```

User password age report:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType PasswordAge `
  -ExportFormat Csv,Html
```

Locked-out user detail:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -IncludeEvents `
  -DaysBack 7 `
  -AllowPartialResults
```
