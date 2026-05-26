# Overview

This repository contains three Active Directory administration scripts:

- `Scan-ADComputers.ps1`: PowerShell 7 inventory and validation for AD computer objects.
- `Get-ADAdminActivity.ps1`: Windows PowerShell 5.1+ Domain Controller Security log audit reporting.
- `Manage-ADUserAccounts.ps1`: Windows PowerShell 5.1+ user account reports, lockout details, focused audit lookup, and explicit reset actions.

## What the Scripts Do

`Scan-ADComputers.ps1` is built in three capability layers:

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

`Get-ADAdminActivity.ps1` can:

- Query writable Domain Controller Security logs
- Report AD administrative events such as user, computer, group, and policy changes
- Filter to current privileged admins with `-AdminOnly`
- Export sanitized CSV audit reports

`Manage-ADUserAccounts.ps1` can:

- Report user account summary, password age, locked-out accounts, privileged users, disabled users, and stale users
- Export single-user audit summaries and optional Security event details
- Unlock one account, enable one account, reset one password, and require password change at next logon
- Export reports as CSV, JSON, and HTML

## Requirements

### PowerShell

- `Scan-ADComputers.ps1`: PowerShell 7 or later
- `Get-ADAdminActivity.ps1`: Windows PowerShell 5.1 or later
- `Manage-ADUserAccounts.ps1`: Windows PowerShell 5.1 or later

### Modules

- `ActiveDirectory`
- RSAT Active Directory tools installed on the machine

### Access Requirements

- A credential with permission to query AD
- Permission to read Domain Controller Security logs for audit reports
- If `-RemoteInventory` is used:
  - Remote CIM/WinRM access must be allowed
  - Firewalls must permit required traffic
  - The credential must have access on target machines

### Network Requirements

- If `-TestMethod Ping` is used, ICMP must be allowed
- If `-TestMethod WinRM` or `-RemoteInventory` is used, WinRM ports must be reachable
- If `-ResolveDns` is used, DNS resolution must work from the machine running the script

## Security Defaults

- Active Directory modules are loaded only from the trusted Windows RSAT module path.
- CSV exports are sanitized by default to reduce spreadsheet formula injection risk.
- Existing report and log files are not overwritten unless `-ForceOverwrite` is supplied.
- Network output paths are rejected unless `-AllowNetworkOutputPath` is supplied.
- Network input paths are rejected unless `-AllowNetworkInputPath` is supplied.
- User-supplied text inputs are checked for control characters before they are used.
- Domain Controller and DNS names are validated before network calls.
- `Scan-ADComputers.ps1 -RemoteInventory` skips targets outside the discovered or supplied AD DNS suffix.
- Generated temporary passwords in `Manage-ADUserAccounts.ps1` require `-ShowGeneratedPassword`.

## High-Level Flow

`Scan-ADComputers.ps1` runs in this order:

1. Read parameters and optional JSON config
2. Prepare output and log paths
3. Prompt for credentials if no credential was supplied
4. Build the AD query scope and filters
5. Run either a full scan or targeted scan
6. Build inventory records with AD metadata
7. Optionally enrich records with DNS, port, and remote inventory checks
8. Export inventory, audit, summary, and delta reports as requested
9. Write a run summary and exit with a structured code

`Get-ADAdminActivity.ps1` runs in this order:

1. Read parameters and validate output controls
2. Load the trusted RSAT `ActiveDirectory` module
3. Discover writable Domain Controllers unless they are supplied
4. Resolve privileged group membership when `-AdminOnly` is used
5. Query Domain Controller Security logs for AD administrative event IDs
6. Normalize event properties into CSV rows
7. Apply privileged-admin filtering when requested
8. Export sanitized CSV output

`Manage-ADUserAccounts.ps1` runs in this order:

1. Read parameters and validate identities, OU scope, paths, and safety switches
2. Load the trusted RSAT `ActiveDirectory` module
3. Discover domain metadata and Domain Controllers as needed
4. Build user query scope from `SearchBase`, `SearchBaseList`, `ExcludeOU`, and `UserListPath`
5. Run the selected mode: `Report`, `UserAudit`, `LockedOut`, or `Reset`
6. Query Security logs only when event details are requested
7. Export CSV, JSON, and HTML reports
8. For reset mode, perform explicit account actions through `ShouldProcess` and write an action report

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
- `NoClobber`
- `ForceOverwrite`
- `AllowNetworkInputPath`
- `AllowNetworkOutputPath`
- `DisableCsvSanitization`

### 2. Logging and Exit Codes

`Scan-ADComputers.ps1` writes a timestamped log file with run details and errors.

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

## Script Manuals

Use the script-specific manuals for complete setup, quick start, usage, examples,
outputs, and safety details:

- [Scan-ADComputers Manual](scan-adcomputers.md)
- [Get-ADAdminActivity Manual](ad-admin-activity.md)
- [Manage-ADUserAccounts Manual](user-account-management.md)
