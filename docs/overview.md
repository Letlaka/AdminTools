# Overview

`Scan-ADComputers.ps1` is a PowerShell 7 inventory and validation script for Active Directory computer objects. It can scan `Server` or `Workstation` objects, run in full-domain or targeted-list mode, perform optional operational checks, and export reports in multiple formats.

## What the Script Does

The script is built in three capability layers:

1. AD inventory and reporting
2. Operational validation
3. Usability and automation support

In practical terms, it can:

- Query enabled or all AD computer objects
- Scan either all matching computers or only a supplied list
- Restrict scanning to one or more OUs
- Exclude specific OUs from final results
- Flag stale devices based on inactivity age
- Export reports as CSV, JSON, and HTML
- Generate targeted audit reports
- Compare the current export with a previous export and generate delta output
- Produce summary-only reports
- Validate DNS resolution
- Test TCP ports
- Attempt remote inventory collection through CIM/WinRM
- Write a run log with timestamps
- Load parameters from a JSON config file
- Return structured exit codes for automation

## Requirements

### PowerShell

- PowerShell 7 or later

### Modules

- `ActiveDirectory`
- RSAT Active Directory tools installed on the machine

### Access Requirements

- A credential with permission to query AD
- If `-RemoteInventory` is used:
  - Remote CIM/WinRM access must be allowed
  - Firewalls must permit required traffic
  - The credential must have access on target machines

### Network Requirements

- If `-TestMethod Ping` is used, ICMP must be allowed
- If `-TestMethod WinRM` or `-RemoteInventory` is used, WinRM ports must be reachable
- If `-ResolveDns` is used, DNS resolution must work from the machine running the script

## High-Level Flow

The script runs in this order:

1. Read parameters and optional JSON config
2. Prepare output and log paths
3. Prompt for credentials if no credential was supplied
4. Build the AD query scope and filters
5. Run either a full scan or targeted scan
6. Build inventory records with AD metadata
7. Optionally enrich records with DNS, port, and remote inventory checks
8. Export inventory, audit, summary, and delta reports as requested
9. Write a run summary and exit with a structured code

## Functional Layers

### 1. Parameters

Main scan controls:

- `ComputerType`
- `Mode`
- `ComputerListPath`
- `SearchBase`
- `SearchBaseList`
- `ExcludeOU`
- `IncludeDisabled`
- `InactiveDays`

Export and reporting controls:

- `ExportFormat`
- `CompareWithPrevious`
- `SummaryOnly`
- `SeparateStatusExports`

Operational checks:

- `TestMethod`
- `SkipPing`
- `PingCount`
- `TimeoutSeconds`
- `ThrottleLimit`
- `ResolveDns`
- `TestPorts`
- `RemoteInventory`

Usability and connection:

- `ConfigPath`
- `LogPath`
- `DomainController`
- `DomainName`
- `Credential`
- `OutputDirectory`

### 2. Logging and Exit Codes

The script writes a timestamped log file with run details and errors.

| Exit Code | Meaning |
|---|---|
| `0` | Success |
| `1` | General failure |
| `2` | Prerequisite failure |
| `3` | Config file failure |
| `4` | Validation failure |
| `5` | AD query failure |
| `6` | Operational enrichment failure |
| `7` | Export failure |
| `8` | Compare/delta failure |

### 3. AD Query Layer

- `Full` mode queries all matching computer objects in scope.
- `Targeted` mode reads names from a file and queries AD by `Name` and `DNSHostName`.
- Targeted mode also produces an audit export for requested inputs.

### 4. Filtering Layer

- Type filter (`Server` vs `Workstation`)
- Enabled-only by default (override with `-IncludeDisabled`)
- Scope filters (`-SearchBase`, `-SearchBaseList`)
- Post-query OU exclusions (`-ExcludeOU`)
- Optional stale-device fields (`-InactiveDays`)

### 5. Operational Enrichment Layer

Optional checks include:

- Connectivity (`Ping`, `WinRM`, `None`)
- DNS validation (`-ResolveDns`)
- TCP checks (`-TestPorts`)
- Remote inventory (`-RemoteInventory`)

### 6. Export Layer

Supports:

- Inventory exports
- Targeted audit exports
- Separate targeted status exports
- Summary exports
- Delta exports (`-CompareWithPrevious`)

Supported formats: `Csv`, `Json`, `Html`.
