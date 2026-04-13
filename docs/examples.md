# Examples

## Basic Full Scan

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

```powershell
.\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Full
```

## Full Scan with Specific Domain Controller

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -DomainController "DomainController.domain.local"
```

## Full Scan with Custom Output Directory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -OutputDirectory "C:\Temp\ADReports"
```

## Full Scan with Scope Controls

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -SearchBase "OU=Workstations,DC=domain,DC=local"
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -SearchBaseList `
    "OU=Production Servers,DC=domain,DC=local", `
    "OU=Test Servers,DC=domain,DC=local"
```

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

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90
```

## Targeted Scan Basics

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt"
```

## Targeted Scan Without Connectivity Checks

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod None
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -SkipPing
```

## Targeted Scan with Ping or WinRM

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

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -TimeoutSeconds 5 `
  -ThrottleLimit 12
```

## Targeted Audit and Status Exports

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -SeparateStatusExports
```

## Multi-Format Export and Summary-Only

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv,Json,Html
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90 `
  -SummaryOnly `
  -ExportFormat Csv,Html
```

## Compare with Previous Export

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -CompareWithPrevious "C:\Temp\ADReports\Servers_domain_local_20260412090000.csv"
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -ExportFormat Json `
  -CompareWithPrevious "C:\Temp\ADReports\Workstations_domain_local_20260412090000.json"
```

## DNS, Ports, and Remote Inventory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -ResolveDns
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestPorts 445,3389,5985
```

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

```powershell
.\Scan-ADComputers.ps1 -ConfigPath ".\Scan-ADComputers.json"
```

```powershell
.\Scan-ADComputers.ps1 `
  -ConfigPath ".\Scan-ADComputers.json" `
  -Mode Full `
  -ExportFormat Csv,Html
```

## Custom Log File and WhatIf

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -LogPath "C:\Temp\ADReports\ServerScan.log"
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -ExportFormat Csv,Html `
  -WhatIf
```
