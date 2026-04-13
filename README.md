# Scan-ADComputers.ps1

`Scan-ADComputers.ps1` is a PowerShell 7 inventory and validation script for Active Directory computer objects. It can scan either `Server` or `Workstation` objects, run in full-domain or targeted-list mode, export multiple report formats, perform optional operational checks, and generate change reports against a previous run.

## What The Script Does

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
- Generate targeted audit reports explaining why names were or were not exported
- Compare the current export with a previous export and generate delta output
- Produce summary-only reports
- Validate DNS resolution
- Test TCP ports
- Attempt remote inventory collection through CIM/WinRM
- Write a run log with timestamps
- Load parameters from a JSON config file
- Return structured exit codes for automation

## Requirements

## PowerShell

- PowerShell 7 or later

## Modules

- `ActiveDirectory`
- RSAT Active Directory tools installed on the machine

## Access Requirements

- A credential with permission to query AD
- If `-RemoteInventory` is used:
  - Remote CIM/WinRM access must be allowed
  - Firewalls must permit the required traffic
  - The credential must have enough access on the target machines

## Network Requirements

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
7. Optionally enrich those records with DNS, port, and remote inventory checks
8. Export inventory, audit, summary, and delta reports as requested
9. Write a run summary and exit with a structured code

## Script Sections

## 1. Parameters

This is the public interface of the script.

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

Usability:

- `ConfigPath`
- `LogPath`

Connection settings:

- `DomainController`
- `DomainName`
- `Credential`
- `OutputDirectory`

## 2. Logging And Exit Codes

The script initializes a timestamped log file and writes status lines such as:

- run start
- mode
- computer type
- scope
- export formats
- scan summary
- errors

Exit codes used by the script:

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

## 3. AD Query Layer

The AD layer decides what objects to fetch.

### Full Mode

In `-Mode Full`, the script queries all matching computer objects of the selected type.

Examples:

- all servers in the domain
- all workstations in a specific OU
- all computers of a type across multiple OUs

### Targeted Mode

In `-Mode Targeted`, the script reads names from a text file and queries AD only for those requested names. It matches by:

- `Name`
- `DNSHostName`

It then produces:

- the main inventory export for machines that qualify
- a targeted audit export for every requested input

The audit report includes data like:

- input name
- resolved short name
- resolved FQDN
- connectivity status
- found in AD
- matched on field
- excluded by OU
- exported or skipped
- skip reason

## 4. Filtering Layer

The script applies several filters.

### Computer Type Filter

- `Server` means `OperatingSystem -like "*Server*"`
- `Workstation` means `OperatingSystem -like "Windows*"` and not server

### Enabled Filter

- Default behavior is enabled-only
- `-IncludeDisabled` removes that restriction

### Search Scope Filter

You can control scope with:

- `-SearchBase`
- `-SearchBaseList`

If both are supplied, the script queries all supplied scopes.

### Exclude OU Filter

`-ExcludeOU` is applied after AD results are returned. Matching objects under excluded OUs are removed from the final export. In targeted mode, they still appear in the audit report with `ExcludedOU`.

### Inactive Device Flagging

If `-InactiveDays` is set, the script computes:

- `LastSeenDate`
- `DaysSinceLastSeen`
- `IsStale`
- `StaleStatus`

The script uses `LastLogonDate` and `LastLogonTimestamp` and takes the most recent available value as the effective "last seen" date.

## 5. Operational Enrichment Layer

This layer is optional. It runs after the AD inventory records are built.

### Connectivity Methods

`-TestMethod` controls connectivity validation.

Valid values:

- `Ping`
- `WinRM`
- `None`

Behavior:

- `Ping` uses ICMP
- `WinRM` checks ports `5985` and `5986`
- `None` skips connectivity checks

`-SkipPing` is a compatibility shortcut that forces `-TestMethod None`.

### DNS Checks

If `-ResolveDns` is used, the script attempts to resolve A records and records:

- `DnsStatus`
- `DnsResolvedIPs`
- `DnsMatchesAdIPv4`

### Port Checks

If `-TestPorts` is supplied, the script checks each port and writes results into:

- `PortStatus`

Example `PortStatus` value:

`445:Open;3389:Closed;5985:Open`

### Remote Inventory

If `-RemoteInventory` is enabled, the script attempts to collect:

- uptime
- serial number
- model
- total memory
- system drive free space
- pending reboot status
- logged-on user

Remote inventory status is captured in:

- `RemoteInventoryStatus`

Possible values include:

- `Success`
- `NoTarget`
- `SkippedUnreachable`
- `Failed: <message>`

## 6. Export Layer

The script can export:

- inventory data
- targeted audit data
- separate targeted status reports
- summary reports
- delta reports

Supported formats:

- `Csv`
- `Json`
- `Html`

### Inventory Export

The main report name pattern is:

`<Servers|Workstations>_<domain>_<timestamp>.<ext>`

Example:

`Servers_domain_20260413133000.csv`

### Targeted Audit Export

When `-Mode Targeted` is used, an audit report is generated:

`<Servers|Workstations>_<domain>_<timestamp>_TargetedAudit.<ext>`

### Separate Targeted Status Exports

When `-SeparateStatusExports` is used in targeted mode, the script also writes:

- `_Matched`
- `_Unreachable`
- `_NotFoundInAD`

### Summary Exports

When `-SummaryOnly` is used, the script skips the main inventory export and writes:

- summary totals
- summary breakdowns

### Delta Exports

When `-CompareWithPrevious` is used, the script generates a delta report showing:

- `Added`
- `Removed`
- `Disabled`
- `ReEnabled`
- `OperatingSystemChanged`

The compare input can be:

- CSV
- JSON

## Parameters Reference

| Parameter | Type | Description |
|---|---|---|
| `ComputerType` | `Server` or `Workstation` | Selects the AD computer category |
| `Mode` | `Full` or `Targeted` | Full scan or list-driven scan |
| `DomainController` | `string` | Domain controller or LDAP server to query |
| `DomainName` | `string` | Domain suffix used when building FQDNs |
| `Credential` | `PSCredential` | Credential for AD query and optional remote inventory |
| `ComputerListPath` | `string` | Required in targeted mode |
| `OutputDirectory` | `string` | Where exports and logs are written |
| `SearchBase` | `string` | Single OU/container scope |
| `SearchBaseList` | `string[]` | Multiple OU/container scopes |
| `ExcludeOU` | `string[]` | OUs to exclude from final results |
| `InactiveDays` | `int` | Marks stale devices based on age |
| `IncludeDisabled` | `switch` | Includes disabled AD objects |
| `ExportFormat` | `Csv`, `Json`, `Html` | One or more export formats |
| `CompareWithPrevious` | `string` | Previous CSV or JSON inventory for delta reporting |
| `SummaryOnly` | `switch` | Export summary instead of inventory |
| `SeparateStatusExports` | `switch` | Write separate targeted reports |
| `ResolveDns` | `switch` | Resolve forward DNS and compare IPs |
| `TestPorts` | `int[]` | TCP ports to test |
| `RemoteInventory` | `switch` | Collect remote machine details |
| `TimeoutSeconds` | `int` | Timeout for connectivity and CIM operations |
| `ThrottleLimit` | `int` | Parallelism limit for connectivity and enrichment |
| `PingCount` | `int` | Number of ICMP echo requests |
| `TestMethod` | `Ping`, `WinRM`, `None` | Connectivity method |
| `SkipPing` | `switch` | Forces `TestMethod` to `None` |
| `ConfigPath` | `string` | JSON config file path |
| `LogPath` | `string` | Custom log file path |

## Output Fields

The main inventory export can include these fields:

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

## File Formats Used By The Script

## Targeted List File

The targeted input file is a plain text file:

- one host per line
- short name or FQDN
- blank lines allowed
- comment lines starting with `#` are ignored

Example:

```text
# Core production servers
Server01
sql01.domain.local
Server02
```

## JSON Config File

The script can load parameters from a JSON config file. Explicit command-line arguments take precedence over config values.

Example:

```json
{
  "ComputerType": "Server",
  "Mode": "Targeted",
  "ComputerListPath": ".\\serverlist.txt",
  "OutputDirectory": ".\\Output",
  "ExportFormat": ["Csv", "Json", "Html"],
  "TestMethod": "Ping",
  "PingCount": 2,
  "ResolveDns": true,
  "TestPorts": [445, 3389, 5985],
  "InactiveDays": 90,
  "SeparateStatusExports": true,
  "TimeoutSeconds": 5,
  "ThrottleLimit": 12
}
```

## Detailed Examples

## Basic Full Scan

Scan all enabled server objects in the domain and export CSV:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

Scan all enabled workstation objects:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Full
```

## Full Scan With A Specific Domain Controller

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -DomainController "DomainController.domain.local"
```

## Full Scan With Custom Output Directory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -OutputDirectory "C:\Temp\ADReports"
```

## Full Scan Limited To One OU

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBase "OU=Workstations,DC=domain,DC=local"
```

## Full Scan Across Multiple OUs

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -SearchBaseList `
    "OU=Production Servers,DC=domain,DC=local", `
    "OU=Test Servers,DC=domain,DC=local"
```

## Full Scan With One Included Scope And One Excluded Scope

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBase "DC=domain,DC=local" `
  -ExcludeOU "OU=Disabled Devices,DC=domain,DC=local"
```

## Include Disabled Objects

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -IncludeDisabled
```

## Stale Device Reporting

Mark devices stale if they have not been seen in 90 days:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90
```

Mark devices stale in a targeted server scan:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -InactiveDays 60
```

## Targeted Scan Basics

Targeted scan using a list file:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt"
```

Targeted workstation scan:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Targeted `
  -ComputerListPath ".\workstationlist.txt"
```

## Targeted Scan Without Connectivity Checks

This is useful if ICMP is blocked or you only care about AD matching:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod None
```

Equivalent compatibility form:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -SkipPing
```

## Targeted Scan Using Ping

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod Ping `
  -PingCount 2 `
  -TimeoutSeconds 5 `
  -ThrottleLimit 10
```

## Targeted Scan Using WinRM Reachability

This checks whether WinRM ports are reachable, rather than using ICMP:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -TimeoutSeconds 5 `
  -ThrottleLimit 12
```

## Targeted Audit And Separate Status Exports

This writes:

- main inventory export
- targeted audit export
- matched export
- unreachable export
- not-found-in-AD export

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -SeparateStatusExports
```

## Export To Multiple Formats

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv,Json,Html
```

## Summary-Only Reporting

This skips the main inventory export and writes summary totals plus summary breakdowns:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90 `
  -SummaryOnly `
  -ExportFormat Csv,Html
```

## Compare Current Run With A Previous Export

Compare against a previous CSV:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -CompareWithPrevious "C:\Temp\ADReports\Servers_domain_local_20260412090000.csv"
```

Compare against a previous JSON export:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -ExportFormat Json `
  -CompareWithPrevious "C:\Temp\ADReports\Workstations_domain_local_20260412090000.json"
```

## DNS Validation

Resolve DNS and compare the resolved A record IPs to AD IPv4 values:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -ResolveDns
```

## TCP Port Checks

Check SMB, RDP, and WinRM ports:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestPorts 445,3389,5985
```

Check ports during a full workstation scan:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBase "OU=Workstations,DC=domain,DC=local" `
  -TestMethod Ping `
  -TestPorts 445,3389
```

## Remote Inventory

Collect remote inventory details through CIM:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -RemoteInventory `
  -TimeoutSeconds 8 `
  -ThrottleLimit 8
```

Remote inventory with DNS and port checks:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -ResolveDns `
  -TestPorts 445,5985,5986 `
  -RemoteInventory `
  -InactiveDays 90
```

## Configuration File Usage

Run entirely from config:

```powershell
.\Scan-ADComputers.ps1 -ConfigPath ".\Scan-ADComputers.json"
```

Run from config but override one value at the command line:

```powershell
.\Scan-ADComputers.ps1 `
  -ConfigPath ".\Scan-ADComputers.json" `
  -ComputerType Workstation
```

Run from config but override output format and mode:

```powershell
.\Scan-ADComputers.ps1 `
  -ConfigPath ".\Scan-ADComputers.json" `
  -Mode Full `
  -ExportFormat Csv,Html
```

## Custom Log File

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -LogPath "C:\Temp\ADReports\ServerScan.log"
```

## WhatIf

The script supports `-WhatIf` because it declares `SupportsShouldProcess`. In practice, this is most useful for previewing file-writing behavior such as report and log creation.

Example:

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv,Html `
  -WhatIf
```

## Example End-To-End Scenarios

## Scenario 1: Daily Server Inventory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -SearchBase "OU=Servers,DC=domain,DC=local" `
  -ExportFormat Csv,Json `
  -InactiveDays 30 `
  -LogPath ".\Logs\DailyServerInventory.log"
```

## Scenario 2: Investigate A Small List Of Problem Servers

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod Ping `
  -ResolveDns `
  -TestPorts 445,3389,5985 `
  -SeparateStatusExports `
  -ExportFormat Csv,Html
```

## Scenario 3: Monthly Workstation Hygiene Report

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBaseList `
    "OU=Workstations,DC=domain,DC=local", `
    "OU=Laptops,DC=domain,DC=local" `
  -ExcludeOU "OU=Retired,DC=domain,DC=local" `
  -InactiveDays 90 `
  -SummaryOnly `
  -ExportFormat Csv,Html
```

## Scenario 4: Change Tracking Between Runs

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv `
  -CompareWithPrevious "C:\Temp\ADReports\Servers_domain_local_20260401080000.csv"
```

## Scenario 5: Full Validation Run For Critical Servers

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\critical-servers.txt" `
  -TestMethod WinRM `
  -ResolveDns `
  -TestPorts 445,135,3389,5985,5986 `
  -RemoteInventory `
  -TimeoutSeconds 10 `
  -ThrottleLimit 6 `
  -ExportFormat Csv,Json,Html
```

## Troubleshooting

## `ActiveDirectory` Module Not Found

Problem:

- The machine does not have RSAT AD tools installed

Fix:

- Install RSAT Active Directory tools
- Confirm `Get-Module -ListAvailable ActiveDirectory` returns a result

## PowerShell Version Error

Problem:

- The script is running under Windows PowerShell 5.1 or an earlier PowerShell version

Fix:

- Run the script in PowerShell 7+

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
- verify the queried OU scopes

## Remote Inventory Fails

Possible causes:

- WinRM not enabled
- firewall blocked
- insufficient credential rights
- CIM permissions denied

Things to check:

- use `-TestMethod WinRM`
- lower `-ThrottleLimit`
- increase `-TimeoutSeconds`
- run without `-RemoteInventory` first

## Compare Report Fails

Possible causes:

- previous file path is wrong
- previous file format is not CSV or JSON
- previous file schema is incompatible with current data

## Recommended Starting Commands

If you are unsure where to start, use one of these:

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
