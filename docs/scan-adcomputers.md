# Scan-ADComputers.ps1 Manual

`Scan-ADComputers.ps1` inventories Active Directory computer objects, validates
operational reachability, enriches records with optional DNS/port/CIM details,
and exports inventory, summary, targeted audit, and delta reports.

## Setup

Run this script from PowerShell 7 or later on a domain-joined admin workstation
or management server.

Required components:

- PowerShell 7+
- RSAT Active Directory module
- Permission to query AD computer objects
- ICMP, WinRM, or TCP access only when those checks are enabled
- Local output directory unless `-AllowNetworkOutputPath` is explicitly used

Recommended first-time check:

```powershell
pwsh
Get-Module -ListAvailable ActiveDirectory
Get-ADDomain
```

## Quick Start

Full server inventory:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

Full workstation inventory:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Full
```

Targeted server inventory from a text file:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt"
```

## High-Level Flow

1. Read command-line parameters and optional JSON config.
2. Validate input values, paths, option sets, and safety switches.
3. Load the trusted RSAT `ActiveDirectory` module.
4. Discover domain metadata when `-DomainController` or `-DomainName` is omitted.
5. Build the AD query from `ComputerType`, `Mode`, and OU scope controls.
6. Query AD computer objects or resolve targeted list entries.
7. Enrich records with stale-device, connectivity, DNS, port, and remote inventory data as requested.
8. Export inventory, targeted audit, status split, summary, and delta reports.
9. Write the run log and exit with a structured exit code.

## Usage

Common scope controls:

- `-ComputerType Server|Workstation`
- `-Mode Full|Targeted`
- `-ComputerListPath` for targeted scans
- `-SearchBase` or `-SearchBaseList` for OU scope
- `-ExcludeOU` for post-query exclusions
- `-IncludeDisabled` to include disabled computer objects
- `-InactiveDays` to mark stale devices

Operational checks:

- `-TestMethod Ping|WinRM|None`
- `-ResolveDns`
- `-TestPorts 445,3389,5985`
- `-RemoteInventory`
- `-TimeoutSeconds`
- `-ThrottleLimit`
- `-ConnectivityThrottleLimit`, `-DnsThrottleLimit`, `-PortThrottleLimit`, and `-RemoteInventoryThrottleLimit` for phase-specific parallelism; each defaults to `-ThrottleLimit` when omitted

AD discovery tuning:

- `-AdResultPageSize` defaults to `1000`
- `-AdSearchScope Base|OneLevel|Subtree` defaults to `Subtree`
- `-TargetedQueryChunkSize` defaults to `40` for targeted list query batching

Performance reporting:

- Stage timings are always written to the run log.
- `-PerformanceSummary` writes CSV and JSON timing summaries beside the normal report.

Reporting controls:

- `-ExportFormat Csv,Json,Html`
- `-SummaryOnly`
- `-SeparateStatusExports`
- `-CompareWithPrevious`
- `-OutputDirectory`
- `-LogPath`
- `-NoProgress` to suppress transient progress displays in unattended runs

Safety controls:

- `-NoClobber`
- `-ForceOverwrite`
- `-AllowNetworkInputPath`
- Credential reuse with `-CredentialSecretName` or `-CredentialPath`; use only one of these or `-Credential`
- `-AllowNetworkOutputPath`
- `-DisableCsvSanitization`

## Targeted List Format

Use one computer name per line. Blank lines are allowed and comment lines
starting with `#` are ignored.

```text
# Production servers
server01
sql01.domain.local
server02
```

Targeted FQDN entries must be under the discovered or supplied AD DNS suffix.

## JSON Config Support

This script supports JSON config files with `-ConfigPath`. Explicit command-line
parameters override values from the config file. Public samples are committed as
`config/server_config.sample.json` and `config/workstation.sample.json`; copy a
sample to an ignored runtime file before storing environment-specific values.

Example full inventory config:

```json
{
  "ComputerType": "Server",
  "Mode": "Full",
  "CredentialSecretName": "ExampleADCredential",
  "ExportFormat": ["Csv", "Json", "Html"],
  "SearchBase": "OU=Servers,DC=domain,DC=local",
  "ExcludeOU": ["OU=Decommissioned,OU=Servers,DC=domain,DC=local"],
  "InactiveDays": 90,
  "TestMethod": "Ping",
  "PingCount": 2,
  "ResolveDns": true,
  "TestPorts": [445, 3389, 5985],
  "SummaryOnly": false,
  "NoProgress": false,
  "NoClobber": true
}
```

Example targeted diagnostic config:

```json
{
  "ComputerType": "Server",
  "Mode": "Targeted",
  "ComputerListPath": ".\\serverlist.txt",
  "ExportFormat": ["Csv", "Html"],
  "TestMethod": "WinRM",
  "TimeoutSeconds": 5,
  "ThrottleLimit": 12,
  "RemoteInventoryThrottleLimit": 8,
  "PerformanceSummary": true,
  "ResolveDns": true,
  "TestPorts": [445, 5985, 5986],
  "RemoteInventory": true,
  "SeparateStatusExports": true,
  "InactiveDays": 90,
  "NoProgress": true,
  "NoClobber": true
}
```

Run with:

```powershell
Copy-Item config/server_config.sample.json config/server_config.json
.\Scan-ADComputers.ps1 -ConfigPath ".\config\server_config.json"
```

Progress is enabled by default and shows live AD discovery, targeted
connectivity and matching, record preparation, operational enrichment, and
export status. For a scheduled or redirected run, suppress transient progress:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full -NoProgress -ExportFormat Csv
```

## Examples

OU-scoped workstation report:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBase "OU=Workstations,DC=domain,DC=local" `
  -InactiveDays 60 `
  -ExportFormat Csv,Html
```

Targeted server diagnostics:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -ResolveDns `
  -TestPorts 445,5985 `
  -TimeoutSeconds 5
```

Targeted server diagnostics with status split exports:

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

Remote inventory with trusted scope controls:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -SearchBase "OU=Production Servers,DC=domain,DC=local" `
  -TestMethod WinRM `
  -RemoteInventory `
  -RemoteInventoryThrottleLimit 8 `
  -TimeoutSeconds 5
```

Delta report from a previous CSV export:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -CompareWithPrevious ".\reports\ad-computers\Servers_domain_local_20260501090000.csv"
```

## Outputs

Default output location is `reports/ad-computers` unless `-OutputDirectory` is supplied. The default log path is in the resolved output directory. When `-PerformanceSummary` is supplied, timing files are written as `<Servers|Workstations>_<domain>_<timestamp>_Performance.csv` and `.json`.

Common outputs:

- `<Servers|Workstations>_<domain>_<timestamp>.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit_Matched.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit_Unreachable.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit_NotFoundInAD.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_Summary.<ext>`
- `<Servers|Workstations>_<domain>_<timestamp>_Delta.<ext>`

## Safety Notes

- Existing output and log files are not overwritten unless `-ForceOverwrite` is supplied.
- UNC output paths require `-AllowNetworkOutputPath`.
- UNC config, targeted list, comparison, and credential input paths require `-AllowNetworkInputPath`.
- CSV values are sanitized by default.
- Remote inventory skips targets outside the AD DNS suffix.
- `-CredentialPath` files must be outside the repository directory and should be readable only by the account running the script.
