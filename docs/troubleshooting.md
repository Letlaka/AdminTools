# Troubleshooting

## `ActiveDirectory` Module Not Found

**Problem**

- The machine does not have RSAT AD tools installed.

**Fix**

- Install RSAT Active Directory tools.
- Confirm `Get-Module -ListAvailable ActiveDirectory` returns a result.

## PowerShell Version Error

**Problem**

- The script is running under Windows PowerShell 5.1 or an earlier version.

**Fix**

- Run the script in PowerShell 7+.

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
- verify queried OU scopes

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
