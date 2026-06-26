# Examples

## Scan-ADComputers.ps1

### Basic Full Scans

```powershell
.\Scan-ADComputers.ps1 -ComputerType Server -Mode Full
```

```powershell
.\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Full
```


### Reusable Credentials

```powershell
Install-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -Scope CurrentUser
Register-SecretVault -Name AdminToolsVault -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
Set-Secret -Name ExampleADCredential -Secret (Get-Credential)

.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -CredentialSecretName ExampleADCredential
```

### Scope Controls

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
    "OU=Test Servers,DC=domain,DC=local" `
  -ExcludeOU "OU=Retired Servers,DC=domain,DC=local"
```

### Targeted Scan

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt"
```

### DNS, Ports, and Remote Inventory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Targeted `
  -ComputerListPath ".\serverlist.txt" `
  -TestMethod WinRM `
  -ResolveDns `
  -TestPorts 445,5985,5986 `
  -RemoteInventory `
  -RemoteInventoryThrottleLimit 8 `
  -TimeoutSeconds 5 `
  -InactiveDays 90
```

### Summary and Delta Reports

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Workstation `
  -Mode Full `
  -InactiveDays 90 `
  -SummaryOnly `
  -ExportFormat Csv,Html
```

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -CompareWithPrevious ".\reports\ad-computers\Servers_domain_local_20260501090000.csv"
```

### Fast Inventory

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -NoProgress `
  -ExportFormat Csv
```

### Unattended Scan Without Progress

```powershell
.\Scan-ADComputers.ps1 `
  -ComputerType Server `
  -Mode Full `
  -NoProgress
```

### JSON Config

`Scan-ADComputers.ps1` supports JSON config files.

```json
{
  "ComputerType": "Server",
  "Mode": "Targeted",
  "ComputerListPath": ".\\serverlist.txt",
  "CredentialSecretName": "ExampleADCredential",
  "ExportFormat": ["Csv", "Json", "Html"],
  "NoClobber": true,
  "NoProgress": true,
  "TestMethod": "WinRM",
  "PingCount": 2,
  "ResolveDns": true,
  "TestPorts": [445, 3389, 5985],
  "RemoteInventory": true,
  "InactiveDays": 90,
  "SeparateStatusExports": true,
  "TimeoutSeconds": 5,
  "ThrottleLimit": 12,
  "RemoteInventoryThrottleLimit": 8,
  "PerformanceSummary": true
}
```

```powershell
.\Scan-ADComputers.ps1 -ConfigPath ".\Scan-ADComputers.json"
```

## Get-ADAdminActivity.ps1

### Basic Audit Report

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 7 `
  -OutputCsv ".\reports\ad-admin-activity\AD_Admin_Activity_Custom.csv"
```

### Privileged Admin Activity

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly `
  -AllowPartialResults
```

### Additional Admin Names

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 30 `
  -AdminOnly `
  -AdminSamAccountNames "svc-ad-admin","tier0.operator"
```

### Specific Domain Controllers

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 14 `
  -DomainControllers "dc01.domain.local","dc02.domain.local" `
  -AllowPartialResults
```

### Include Rendered Event Messages

```powershell
.\Get-ADAdminActivity.ps1 `
  -DaysBack 3 `
  -IncludeMessage `
  -OutputCsv ".\reports\ad-admin-activity\AD_Admin_Activity_WithMessages.csv"
```

### JSON Config

`Get-ADAdminActivity.ps1` does not currently support JSON config files. Use a
wrapper script for repeatable presets:

```powershell
$domainControllers = @("dc01.domain.local", "dc02.domain.local")

.\Get-ADAdminActivity.ps1 `
  -DaysBack 14 `
  -DomainControllers $domainControllers `
  -AdminOnly `
  -AllowPartialResults
```

## Manage-ADUserAccounts.ps1

### User Account Reports

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge `
  -PasswordAgeWarningDays 90 `
  -ExportFormat Csv,Html
```

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType PrivilegedUsers,DisabledUsers,StaleUsers `
  -SearchBase "OU=Users,DC=domain,DC=local" `
  -IncludeDisabled `
  -ExportFormat Csv
```

### Locked-Out User Details

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode LockedOut `
  -IncludeEvents `
  -DaysBack 7 `
  -ExportFormat Csv,Html
```

### Single-User Audit

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode UserAudit `
  -Identity jsmith `
  -IncludeEvents `
  -DaysBack 30 `
  -ExportFormat Csv,Json
```

### User Reset Actions

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -Unlock `
  -WhatIf
```

```powershell
.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -ResetPassword `
  -GenerateTemporaryPassword `
  -ShowGeneratedPassword `
  -ChangePasswordAtLogon
```

```powershell
$password = Read-Host "Temporary password" -AsSecureString

.\Manage-ADUserAccounts.ps1 `
  -Mode Reset `
  -Identity jsmith `
  -ResetPassword `
  -NewPassword $password `
  -ChangePasswordAtLogon
```

### JSON Config

`Manage-ADUserAccounts.ps1` does not currently support JSON config files. Use a
wrapper script for repeatable presets:

```powershell
$userScopes = @("OU=Users,DC=domain,DC=local", "OU=Admins,DC=domain,DC=local")

.\Manage-ADUserAccounts.ps1 `
  -Mode Report `
  -ReportType UserSummary,PasswordAge,PrivilegedUsers `
  -SearchBaseList $userScopes `
  -OutputDirectory ".\reports\ad-user-accounts\scoped-users" `
  -ExportFormat Csv,Html `
  -NoClobber
```

## AD Excel Reporting

### Excel Dashboard Reports

```powershell
uv sync

uv run python scripts/build_ad_excel_reports.py `
  --servers reports/ad-computers/Servers_example_corp_local_20260625.csv `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260625.csv `
  --as-of-date 2026-06-25
```

