# Usage Guide

## Running the Script

From the repository root:

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

## Modes

- `Full`: scans all matching computers in query scope.
- `Targeted`: scans only names listed in `-ComputerListPath`.

## Scope and Filtering

Use these controls to target results:

- `-SearchBase` for a single DN scope
- `-SearchBaseList` for multiple DN scopes
- `-ExcludeOU` to remove matching OUs from final export
- `-IncludeDisabled` to include disabled AD objects
- `-InactiveDays` to calculate stale-device fields

## Operational Validation

- `-TestMethod Ping|WinRM|None`
- `-SkipPing` (compatibility shortcut for `-TestMethod None`)
- `-ResolveDns`
- `-TestPorts`
- `-RemoteInventory`
- `-TimeoutSeconds`
- `-ThrottleLimit`

## Export Behavior

- `-ExportFormat Csv,Json,Html`
- `-SummaryOnly` for summary output instead of full inventory
- `-SeparateStatusExports` for targeted mode status splits
- `-CompareWithPrevious` to generate delta reports

## Configuration File

Parameters can be loaded from JSON with `-ConfigPath`. Explicit command-line parameters override config values.

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

## WhatIf Support

The script supports `-WhatIf` and can be used to preview file-writing behavior for report and log creation.
