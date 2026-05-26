# Usage Guide

This guide gives the common operating patterns for all scripts. For the deepest
per-script details, use the dedicated manuals:

- [Scan-ADComputers Manual](scan-adcomputers.md)
- [Get-ADAdminActivity Manual](ad-admin-activity.md)
- [Manage-ADUserAccounts Manual](user-account-management.md)

## Setup for All Scripts

Run from a domain-joined administrative workstation or management server.

Required baseline:

- RSAT Active Directory module installed
- A credential with the delegated rights required by the task
- Local output paths unless a trusted UNC output path is explicitly allowed
- PowerShell execution policy and script signing controls aligned with your environment

PowerShell version requirements:

- `Scan-ADComputers.ps1`: PowerShell 7+
- `Get-ADAdminActivity.ps1`: Windows PowerShell 5.1+
- `Manage-ADUserAccounts.ps1`: Windows PowerShell 5.1+

Baseline checks:

```powershell
Get-Module -ListAvailable ActiveDirectory
Get-ADDomain
```

## Quick Start for All Scripts

Computer inventory:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

Admin activity audit:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

User account password age report:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType PasswordAge `
  -ExportFormat Csv,Html
```

Locked-out account details:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -IncludeEvents `
  -DaysBack 7 `
  -ExportFormat Csv,Html
```

Preview a reset action:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -Unlock `
  -WhatIf
```

## Scan-ADComputers.ps1 Usage

Run a full scan:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Full
```

Run a targeted scan:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt"
```

Use OU scope controls:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -SearchBase "OU=Servers,DC=domain,DC=local" `
  -ExcludeOU "OU=Decommissioned,OU=Servers,DC=domain,DC=local"
```

Add operational validation:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -ResolveDns `
  -TestPorts 445,5985,5986 `
  -SeparateStatusExports
```

Use a JSON config file:

```powershell
.\Scan-ADComputers.ps1 -ConfigPath ".\Scan-ADComputers.json"
```

`Scan-ADComputers.ps1` is the only current script with native `-ConfigPath`
support.

## Get-ADAdminActivity.ps1 Usage

Run a default audit:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

Limit results to privileged admins:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly
```

Query specific Domain Controllers and allow partial results:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 14 `
  -DomainControllers "dc01.domain.local","dc02.domain.local" `
  -AllowPartialResults
```

Write to a specific CSV:

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 7 `
  -OutputCsv "C:\Temp\ADReports\AD_Admin_Activity.csv"
```

This script does not currently support JSON config files.

## Manage-ADUserAccounts.ps1 Usage

Generate multiple user reports:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge,PrivilegedUsers `
  -SearchBase "OU=Users,DC=domain,DC=local" `
  -ExportFormat Csv,Html
```

Audit one user:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode UserAudit `
  -Identity jsmith `
  -IncludeEvents `
  -DaysBack 30 `
  -ExportFormat Csv,Json
```

Report lockouts with event details:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -IncludeEvents `
  -DaysBack 7 `
  -AllowPartialResults `
  -ExportFormat Csv,Html
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

This script does not currently support JSON config files.

## WhatIf Support

`Scan-ADComputers.ps1` supports `-WhatIf` for report and log creation behavior:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv,Html `
  -WhatIf
```

`Manage-ADUserAccounts.ps1` supports `-WhatIf` for account-changing reset
actions:

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -Unlock `
  -WhatIf
```

`Get-ADAdminActivity.ps1` is read-only against AD and event logs. It does not
modify AD objects.

## Common Safety Defaults

- CSV exports are sanitized by default.
- Existing outputs are not overwritten unless `-ForceOverwrite` is supplied.
- Network output paths are rejected unless `-AllowNetworkOutputPath` is supplied.
- Network input paths are rejected unless the relevant script supports `-AllowNetworkInputPath` and it is supplied.
- Domain Controller names are validated before event-log queries.
- Generated temporary passwords are displayed only when `-ShowGeneratedPassword` is explicitly supplied.
