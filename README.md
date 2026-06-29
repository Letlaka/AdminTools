# AdminTools

AdminTools is a set of Active Directory administration and reporting tools for
computer inventory, administrative activity auditing, user account reporting, and
Excel dashboard generation.

## Features

- `Scan-ADComputers.ps1`: AD computer inventory, targeted scans, stale-device reporting, DNS and port validation, optional remote inventory, CSV/JSON/HTML exports, and performance summaries.
- `Get-ADAdminActivity.ps1`: Domain Controller Security log reporting for AD administrative events, with optional privileged-admin filtering.
- `Manage-ADUserAccounts.ps1`: AD user account reports, lockout details, single-user audit lookups, and explicit unlock, enable, and password reset actions.
- `scripts/build_ad_excel_reports.py`: Excel dashboards and per-department workbooks from `Scan-ADComputers.ps1` CSV or JSON exports.

## Requirements

Run the PowerShell tools from a domain-joined administrative workstation or
management server.

Required for the AD PowerShell tools:

- Windows with RSAT Active Directory tools installed.
- The `ActiveDirectory` PowerShell module.
- A credential with the delegated rights needed for the task.
- Permission to read Domain Controller Security logs when using audit/event reports.
- PowerShell 7+ for `Scan-ADComputers.ps1`.
- Windows PowerShell 5.1+ for `Get-ADAdminActivity.ps1` and `Manage-ADUserAccounts.ps1`.

Required for Excel report generation:

- Python 3.13+.
- `uv`, or another Python environment with `openpyxl` installed.

Optional, depending on feature use:

- CIM/WinRM access and firewall rules for `Scan-ADComputers.ps1 -RemoteInventory`.
- DNS resolution from the admin workstation for `-ResolveDns`.
- TCP reachability for `-TestPorts` or `-TestMethod WinRM`.
- Microsoft.PowerShell.SecretManagement and SecretStore for reusable credentials.

## Install and Prepare

Clone the repository and switch into it:

```powershell
git clone https://github.com/<owner>/<repo>.git
Set-Location .\AdminTools
```

If you downloaded a ZIP instead of cloning, extract it and run the commands from
the extracted repository root.

Check the AD module from the shell you plan to use:

```powershell
Get-Module -ListAvailable ActiveDirectory
Get-ADDomain
```

If the module is missing, install RSAT Active Directory tools on the workstation
or management server first.

If your execution policy blocks local scripts, use the policy that matches your
environment. For a current-user development workstation, this is the typical
local setup:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Install the Python dependency for Excel reporting:

```powershell
uv sync
```

If you are not using `uv`, create a Python 3.13+ environment and install the
package dependency declared in `pyproject.toml`:

```powershell
python -m pip install openpyxl
```

## Configure Credentials

All three AD PowerShell tools accept one credential source at a time:

- `-Credential` for an explicit `PSCredential` object.
- `-CredentialSecretName` for a SecretManagement secret containing a `PSCredential`.
- `-CredentialPath` for a DPAPI-protected credential file created with `Export-Clixml`.

SecretManagement is the preferred reusable setup for an administrator
workstation:

```powershell
Install-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -Scope CurrentUser
Register-SecretVault -Name AdminToolsVault -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
Set-Secret -Name ExampleADCredential -Secret (Get-Credential)
```

Use the secret with any AD tool:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full -CredentialSecretName ExampleADCredential
```

A local DPAPI credential file is useful for single-user, single-machine runs.
Store it outside the repository:

```powershell
$credentialPath = Join-Path $env:USERPROFILE ".admintools\ad-reporting.credential.xml"
New-Item -ItemType Directory -Force -Path (Split-Path $credentialPath) | Out-Null
Get-Credential | Export-Clixml -LiteralPath $credentialPath

.\Get-ADAdminActivity.ps1 -DaysBack 7 -CredentialPath $credentialPath
```

Do not commit credential files. Credential files must stay outside the repository
and are not portable across Windows users or computers.

## Configure Runtime Files

`Scan-ADComputers.ps1` can run from command-line parameters or a JSON config
file. Start from the committed samples and edit the local runtime copies:

```powershell
Copy-Item config\server_config.sample.json config\server_config.json
Copy-Item config\workstation.sample.json config\workstation.json
Copy-Item config\servers.sample.txt config\servers.txt
Copy-Item config\workstations.sample.txt config\workstations.txt
```

Excel department reporting uses local department files. Copy the samples and
replace the sample values with your environment's department names and codes:

```powershell
Copy-Item config\dept_list.sample.txt config\dept_list.txt
Copy-Item config\dept_codes.sample.txt config\dept_codes.txt
```

Runtime config files are ignored by git so environment-specific names, codes,
server lists, and workstation lists do not get committed.

## First Runs

Run a computer inventory scan:

```powershell
pwsh
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full -ExportFormat Csv,Html
```

Run a targeted workstation scan from a list:

```powershell
pwsh
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Targeted `
  -ComputerListPath .\config\workstations.txt `
  -ResolveDns `
  -SeparateStatusExports
```

Run a scan from JSON config:

```powershell
pwsh
.\Scan-ADComputers.ps1 -ConfigPath .\config\server_config.json
```

Export recent AD administrative activity:

```powershell
powershell
.\Get-ADAdminActivity.ps1 -DaysBack 7
```

Export user account reports:

```powershell
powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge,PrivilegedUsers `
  -ExportFormat Csv,Html
```

Preview an account unlock before making a change:

```powershell
powershell
.\Manage-ADUserAccounts.ps1 -Mode Reset -Identity jsmith -Unlock -WhatIf
```

Build Excel dashboards and department workbooks from `Scan-ADComputers.ps1`
exports:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --servers reports/ad-computers/Servers_example_corp_local_20260625.csv `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260625.csv `
  --as-of-date 2026-06-25
```

Generate only one Excel report scope when you do not need the full workbook set:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260625.csv `
  --report-scope main
```

## Outputs

Default output locations:

- Computer inventory reports: `reports/ad-computers/`.
- AD admin activity reports: `reports/ad-admin-activity/`.
- User account reports: `reports/ad-user-accounts/`.
- Excel dashboard reports: `reports/<financial-year>/<run-date>/`.
- Run logs: `logs/<tool-name>/` or `logs/excel-reporting/<financial-year>/<run-date>/`.

The tools default to local paths, CSV sanitization, and no silent overwrites.
Network paths, existing-file overwrites, and generated password display require
explicit opt-in switches.

## Safety Defaults

- CSV exports are sanitized by default to reduce spreadsheet formula injection risk.
- Existing report and log files are not overwritten unless `-ForceOverwrite` is supplied.
- Network output paths are rejected unless `-AllowNetworkOutputPath` is supplied.
- Network input paths are rejected unless `-AllowNetworkInputPath` is supplied.
- Stored credential files must be outside the repository directory.
- `Manage-ADUserAccounts.ps1 -Mode Reset` supports `-WhatIf` and handles one identity at a time.
- Generated temporary passwords are shown only when `-ShowGeneratedPassword` is supplied.

## Documentation

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

Regenerate the combined and script-specific manuals from the Markdown sources:

```powershell
.\tools\New-AdminToolsManual.ps1
```
