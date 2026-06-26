# AdminTools Documentation

This repository contains Active Directory administration scripts for inventory,
audit reporting, and basic user account management.

- `Scan-ADComputers.ps1`: PowerShell 7 script for Active Directory computer inventory, validation, and reporting.
- `Get-ADAdminActivity.ps1`: Windows PowerShell 5.1+ script for Domain Controller Security log audit reporting.
- `Manage-ADUserAccounts.ps1`: Windows PowerShell 5.1+ script for AD user account reports, lockout details, single-user audit lookups, and explicit reset actions.
- `build_ad_excel_reports.py`: Python reporting utility for Excel dashboards and department workbooks from computer inventory exports.

## Documentation Index

- [Overview](docs/overview.md)
- [Usage Guide](docs/usage.md)
- [Parameters Reference](docs/parameters.md)
- [Outputs and Report Files](docs/outputs.md)
- [Examples](docs/examples.md)
- [Scan-ADComputers Manual](docs/scan-adcomputers.md)
- [AD Excel Reporting](docs/ad-excel-reporting.md)
- [Get-ADAdminActivity Manual](docs/ad-admin-activity.md)
- [User Account Management](docs/user-account-management.md)
- [Troubleshooting](docs/troubleshooting.md)

## Repository Standards

- [Contributing Guide](CONTRIBUTING.md)
- [Security Policy](SECURITY.md)
- [Changelog](CHANGELOG.md)

## Quick Start

Computer inventory:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

Admin activity audit:

```powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

User account reporting:

```powershell
.\Manage-ADUserAccounts.ps1 -Mode Report -ReportType PasswordAge -ExportFormat Csv,Html
```

Preview a user unlock:

```powershell
.\Manage-ADUserAccounts.ps1 -Mode Reset -Identity jsmith -Unlock -WhatIf
```

Build Excel dashboards and department workbooks from `Scan-ADComputers` exports:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --servers reports/ad-computers/Servers_example_corp_local_20260625.csv `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260625.csv
```

The scripts default to local input/output paths, CSV sanitization, and no silent
overwrites. Network paths and generated password display require explicit opt-in
switches.

Regenerate the combined and script-specific manuals from the Markdown sources:

```powershell
.\tools\New-AdminToolsManual.ps1
```

For prerequisites, setup, and environment requirements, start with [Overview](docs/overview.md).

