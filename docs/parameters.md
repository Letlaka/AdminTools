# Parameters Reference

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
