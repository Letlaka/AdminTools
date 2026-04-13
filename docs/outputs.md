# Outputs and Report Files

## Export Types

The script can export:

- Inventory data
- Targeted audit data
- Separate targeted status reports
- Summary reports
- Delta reports

Supported formats:

- `Csv`
- `Json`
- `Html`

## File Name Patterns

- Main inventory: `<Servers|Workstations>_<domain>_<timestamp>.<ext>`
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
