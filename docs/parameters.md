# Parameters Reference

## Scan-ADComputers.ps1

| Parameter | Type | Description |
|---|---|---|
| `ComputerType` | `Server` or `Workstation` | Selects the AD computer category |
| `Mode` | `Full` or `Targeted` | Full scan or list-driven scan |
| `DomainController` | `string` | Domain controller or LDAP server to query; discovered from AD when omitted |
| `DomainName` | `string` | Domain suffix used when building FQDNs; discovered from AD when omitted |
| `Credential` | `PSCredential` | Credential for AD query and optional remote inventory |
| `CredentialSecretName` | `string` | SecretManagement secret name containing a `PSCredential` |
| `CredentialPath` | `string` | Path to a DPAPI-protected `Export-Clixml` credential file outside the repo |
| `ComputerListPath` | `string` | Required in targeted mode |
| `OutputDirectory` | `string` | Where report exports are written |
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
| `ThrottleLimit` | `int` | Default parallelism limit for connectivity and enrichment phases |
| `ConnectivityThrottleLimit` | `int` | Optional connectivity phase parallelism; inherits `ThrottleLimit` when `0` |
| `DnsThrottleLimit` | `int` | Optional DNS phase parallelism; inherits `ThrottleLimit` when `0` |
| `PortThrottleLimit` | `int` | Optional TCP port phase parallelism; inherits `ThrottleLimit` when `0` |
| `RemoteInventoryThrottleLimit` | `int` | Optional CIM remote inventory parallelism; inherits `ThrottleLimit` when `0` |
| `AdResultPageSize` | `int` | AD query page size; default `1000` |
| `AdSearchScope` | `Base`, `OneLevel`, `Subtree` | AD search scope; default `Subtree` |
| `TargetedQueryChunkSize` | `int` | Targeted list AD filter chunk size; default `40` |
| `PerformanceSummary` | `switch` | Writes CSV and JSON stage timing summaries beside normal reports |
| `PingCount` | `int` | Number of ICMP echo requests |
| `TestMethod` | `Ping`, `WinRM`, `None` | Connectivity method |
| `SkipPing` | `switch` | Forces `TestMethod` to `None` |
| `ConfigPath` | `string` | JSON config file path |
| `LogPath` | `string` | Custom log file path |
| `NoProgress` | `switch` | Suppresses transient progress displays for unattended runs |
| `NoClobber` | `switch` | Fails if an output or log file already exists |
| `ForceOverwrite` | `switch` | Overwrites existing output or log files |
| `AllowNetworkOutputPath` | `switch` | Allows writing reports and custom log paths to UNC paths |
| `AllowNetworkInputPath` | `switch` | Allows reading config, list, comparison, or credential files from UNC paths |
| `DisableCsvSanitization` | `switch` | Exports raw CSV strings without spreadsheet formula protection |

## Manage-ADUserAccounts.ps1

| Parameter | Type | Description |
|---|---|---|
| `Mode` | `Report`, `UserAudit`, `Reset`, `LockedOut` | Selects the user account operation |
| `Identity` | `string` | User identity for audit or reset actions |
| `UserListPath` | `string` | Optional text file of user identities for reporting |
| `SearchBase` | `string` | Single OU/container scope |
| `SearchBaseList` | `string[]` | Multiple OU/container scopes |
| `ExcludeOU` | `string[]` | OUs to exclude from scoped reports |
| `DomainControllers` | `string[]` | Domain Controllers to query for Security events |
| `Credential` | `PSCredential` | Credential for AD queries and event reads |
| `CredentialSecretName` | `string` | SecretManagement secret name containing a `PSCredential` |
| `CredentialPath` | `string` | Path to a DPAPI-protected `Export-Clixml` credential file outside the repo |
| `ReportType` | `UserSummary`, `PasswordAge`, `LockedOut`, `AuditEvents`, `PrivilegedUsers`, `DisabledUsers`, `StaleUsers` | One or more reports to generate |
| `ExportFormat` | `Csv`, `Json`, `Html` | One or more export formats |
| `DaysBack` | `int` | Event query lookback window |
| `PasswordAgeWarningDays` | `int` | Password age threshold for warning status |
| `StaleUserDays` | `int` | Last-logon threshold for stale-user reporting |
| `MaxEventsPerDomainController` | `int` | Optional event limit per DC; `0` means no explicit limit |
| `PrivilegedGroupNames` | `string[]` | Groups used for privileged user reporting |
| `OutputDirectory` | `string` | Where user account reports are written |
| `OutputPrefix` | `string` | Report filename prefix |
| `LogPath` | `string` | Custom run log file path |
| `Unlock` | `switch` | Unlocks the selected user in `Reset` mode |
| `Enable` | `switch` | Enables the selected user in `Reset` mode |
| `ResetPassword` | `switch` | Resets the selected user's password |
| `ChangePasswordAtLogon` | `switch` | Requires password change at next logon |
| `NewPassword` | `securestring` | Password used with `ResetPassword` |
| `GenerateTemporaryPassword` | `switch` | Generates a temporary password for `ResetPassword` |
| `ShowGeneratedPassword` | `switch` | Required with `GenerateTemporaryPassword` to explicitly display the generated password once |
| `IncludeEvents` | `switch` | Adds Security event reports to `UserAudit` or `LockedOut` mode |
| `IncludeGroupMembership` | `switch` | Includes semicolon-delimited group DNs in summary output |
| `IncludeDisabled` | `switch` | Includes disabled users in scoped reports |
| `IncludeMessage` | `switch` | Includes rendered event messages in audit exports |
| `AllowPartialResults` | `switch` | Allows event reports when some DCs cannot be queried |
| `AllowUnverifiedDomainController` | `switch` | Uses supplied DC names without AD verification |
| `NoClobber` | `switch` | Fails if an output or log file already exists |
| `ForceOverwrite` | `switch` | Overwrites existing output or log files |
| `AllowNetworkOutputPath` | `switch` | Allows writing reports and custom log paths to UNC paths |
| `AllowNetworkInputPath` | `switch` | Allows reading user list or credential files from UNC paths |
| `DisableCsvSanitization` | `switch` | Exports raw CSV strings without spreadsheet formula protection |

## Get-ADAdminActivity.ps1

| Parameter | Type | Description |
|---|---|---|
| `DaysBack` | `int` | Security event query lookback window |
| `DomainControllers` | `string[]` | Domain Controllers to query for Security events |
| `Credential` | `PSCredential` | Credential for AD discovery, group lookups, and Security log reads |
| `CredentialSecretName` | `string` | SecretManagement secret name containing a `PSCredential` |
| `CredentialPath` | `string` | Path to a DPAPI-protected `Export-Clixml` credential file outside the repo |
| `AdminOnly` | `switch` | Includes only events performed by current privileged admins or supplied admin names |
| `AdminSamAccountNames` | `string[]` | Extra admin account names for `AdminOnly` matching |
| `PrivilegedGroupNames` | `string[]` | Groups used to resolve current privileged admins |
| `OutputCsv` | `string` | CSV report path |
| `LogPath` | `string` | Custom run log file path |
| `IncludeMessage` | `switch` | Includes rendered event messages in the export |
| `MaxAttributeValueLength` | `int` | Maximum exported event attribute value length |
| `MaxEventsPerDomainController` | `int` | Optional event limit per DC; `0` means no explicit limit |
| `AllowPartialResults` | `switch` | Allows export when some DCs cannot be queried |
| `NoClobber` | `switch` | Fails if the output CSV or log file already exists |
| `ForceOverwrite` | `switch` | Overwrites an existing output CSV or log file |
| `AllowNetworkOutputPath` | `switch` | Allows writing the CSV report and custom log paths to UNC paths |
| `AllowNetworkInputPath` | `switch` | Allows reading credential files from UNC paths |
| `AllowUnverifiedDomainController` | `switch` | Uses supplied DC names without AD discovery verification |
| `DisableCsvSanitization` | `switch` | Exports raw CSV strings without spreadsheet formula protection |
