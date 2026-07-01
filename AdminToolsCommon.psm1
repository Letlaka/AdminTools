#requires -Version 5.1
<#
.SYNOPSIS
    Shared utility functions for AdminTools scripts.
.DESCRIPTION
    Import this module at the top of each AdminTools script:
        Import-Module (Join-Path $PSScriptRoot "AdminToolsCommon.psm1") -Force
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseSingularNouns",
    "",
    Justification = "These established public helpers return or validate collections; renaming them would break callers."
)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSAvoidUsingPlainTextForPassword",
    "",
    Justification = "CredentialSecretName is a SecretManagement lookup key and CredentialPath is a file path; neither parameter carries a password value."
)]
param()

Set-StrictMode -Version Latest

$Script:DefaultPrivilegedGroupNames = @(
    "Domain Admins",
    "Enterprise Admins",
    "Schema Admins",
    "Administrators",
    "Account Operators",
    "Server Operators",
    "Backup Operators",
    "Group Policy Creator Owners"
)

function Get-DefaultPrivilegedGroupNames {
    return @($Script:DefaultPrivilegedGroupNames)
}

# --- AD Module ---

function Test-IsTrustedActiveDirectoryModulePath {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($env:WINDIR)) {
        return $false
    }

    $windowsRoot = (Join-Path -Path $env:WINDIR -ChildPath "System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory").TrimEnd('\')
    $modulePath = $Path.TrimEnd('\')

    return (
        $modulePath.Equals($windowsRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $modulePath.StartsWith("$windowsRoot\", [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Import-ActiveDirectoryModule {
    [CmdletBinding()]
    param(
        [scriptblock]$OnError = { param($Message) Write-Warning $Message }
    )

    $LoadedModules = @(Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)
    if ($LoadedModules.Count -gt 0) {
        foreach ($LoadedModule in $LoadedModules) {
            if (-not (Test-IsTrustedActiveDirectoryModulePath -Path $LoadedModule.ModuleBase)) {
                & $OnError "ActiveDirectory module is already loaded from an untrusted path: $($LoadedModule.ModuleBase)"
                return $false
            }
        }

        return $true
    }

    $TrustedModule = @(Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue |
            Where-Object { Test-IsTrustedActiveDirectoryModulePath -Path $_.ModuleBase } |
            Sort-Object -Property Version -Descending |
            Select-Object -First 1)

    if ($TrustedModule.Count -eq 0) {
        & $OnError "Trusted ActiveDirectory module could not be found in the Windows RSAT module path."
        return $false
    }

    try {
        $ModuleSpecification = @{
            ModuleName    = "ActiveDirectory"
            ModuleVersion = $TrustedModule[0].Version
        }

        if ($TrustedModule[0].Guid -and $TrustedModule[0].Guid -ne [guid]::Empty) {
            $ModuleSpecification["Guid"] = $TrustedModule[0].Guid
        }

        Import-Module -FullyQualifiedName $ModuleSpecification -ErrorAction Stop
        return $true
    }
    catch {
        & $OnError "ActiveDirectory module could not be loaded from trusted path '$($TrustedModule[0].ModuleBase)'. Error: $($_.Exception.Message)"
        return $false
    }
}

# --- Event Log Helpers ---

function ConvertTo-EventDataMap {
    param(
        [object]$EventRecord
    )

    $EventDataMap = @{}
    try {
        [xml]$EventXml = $EventRecord.ToXml()
        foreach ($DataElement in $EventXml.Event.EventData.Data) {
            if ($null -ne $DataElement.Name) {
                $EventDataMap[$DataElement.Name] = [string]$DataElement.'#text'
            }
        }
    }
    catch {
        Write-Warning ("Skipped event record {0} (ID {1}) due to XML parse failure: {2}" -f
            $EventRecord.RecordId, $EventRecord.Id, $_.Exception.Message)
    }
    return $EventDataMap
}

function Get-MapValue {
    param(
        [hashtable]$Map,
        [string]$Key
    )

    if ($Map.ContainsKey($Key)) {
        return $Map[$Key]
    }

    return $null
}

function Join-DomainAccount {
    param(
        [string]$DomainName,
        [string]$AccountName
    )

    if ([string]::IsNullOrWhiteSpace($AccountName) -or $AccountName -eq "-") {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($DomainName) -or $DomainName -eq "-") {
        return $AccountName
    }

    return "$DomainName\$AccountName"
}

function Limit-TextLength {
    param(
        [string]$Text,
        [int]$MaximumLength
    )

    if ([string]::IsNullOrEmpty($Text) -or $Text.Length -le $MaximumLength) {
        return $Text
    }

    return $Text.Substring(0, $MaximumLength) + "...[truncated]"
}

function Split-CredentialUserName {
    param(
        [PSCredential]$Credential
    )

    $UserName = $Credential.UserName

    if ($UserName -match '\\') {
        $Parts = $UserName.Split('\', 2)
        return [pscustomobject]@{
            Domain = $Parts[0]
            User   = $Parts[1]
        }
    }

    return [pscustomobject]@{
        Domain = $null
        User   = $UserName
    }
}

function New-SecurityEventXPathQuery {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "This pure helper constructs and returns an XPath string; it does not change system state."
    )]
    param(
        [int[]]$Ids,
        [string]$ProviderName,
        [datetime]$StartDate
    )

    $EventIdFilter = (($Ids | ForEach-Object { "EventID=$_" }) -join " or ")
    $StartUtc = $StartDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", [System.Globalization.CultureInfo]::InvariantCulture)

    return "*[System[Provider[@Name='$ProviderName'] and ($EventIdFilter) and TimeCreated[@SystemTime >= '$StartUtc']]]"
}

function Get-SecurityAuditEvents {
    param(
        [string]$ComputerName,
        [int[]]$Ids,
        [string]$ProviderName,
        [datetime]$StartDate,
        [int]$MaxEvents,
        [PSCredential]$Credential,
        [scriptblock]$ProcessEvent
    )

    $ProcessedRecords = New-Object System.Collections.Generic.List[object]
    $EventsRead = 0
    $LimitReached = $false
    $UnlimitedEvents = $MaxEvents -le 0

    function Add-SecurityEventQueryMetadata {
        param(
            [object]$Record,
            [int]$MaxEventsPerDomainController,
            [bool]$LimitReached,
            [int]$EventsReadFromDomainController,
            [bool]$UnlimitedEvents
        )

        if ($null -eq $Record) {
            return
        }

        $Metadata = @{
            EventQueryMaxEventsPerDomainController   = $MaxEventsPerDomainController
            EventQueryLimitReached                   = $LimitReached
            EventQueryEventsReadFromDomainController = $EventsReadFromDomainController
            EventQueryUnlimitedEvents                = $UnlimitedEvents
        }

        foreach ($Entry in $Metadata.GetEnumerator()) {
            $ExistingProperty = $Record.PSObject.Properties[$Entry.Key]
            if ($ExistingProperty) {
                $ExistingProperty.Value = $Entry.Value
            }
            else {
                $Record | Add-Member -NotePropertyName $Entry.Key -NotePropertyValue $Entry.Value
            }
        }
    }

    function Complete-SecurityEventQuery {
        foreach ($Record in $ProcessedRecords) {
            Add-SecurityEventQueryMetadata `
                -Record $Record `
                -MaxEventsPerDomainController $MaxEvents `
                -LimitReached $LimitReached `
                -EventsReadFromDomainController $EventsRead `
                -UnlimitedEvents $UnlimitedEvents
        }

        [pscustomobject]@{
            Records                        = @($ProcessedRecords.ToArray())
            EventsReadFromDomainController = $EventsRead
            MaxEventsPerDomainController   = $MaxEvents
            LimitReached                   = $LimitReached
            UnlimitedEvents                = $UnlimitedEvents
        }
    }

    if (-not $Credential) {
        $GetWinEventParameters = @{
            ComputerName   = $ComputerName
            FilterHashtable = @{
                LogName      = "Security"
                ProviderName = $ProviderName
                Id           = $Ids
                StartTime    = $StartDate
            }
            ErrorAction    = "Stop"
        }

        foreach ($EventRecord in Get-WinEvent @GetWinEventParameters) {
            try {
                if ($MaxEvents -gt 0 -and $EventsRead -ge $MaxEvents) {
                    $LimitReached = $true
                    break
                }

                $EventsRead++
                $ProcessedOutput = if ($ProcessEvent) { & $ProcessEvent $EventRecord } else { $EventRecord }
                foreach ($ProcessedRecord in @($ProcessedOutput)) {
                    if ($null -ne $ProcessedRecord) {
                        [void]$ProcessedRecords.Add($ProcessedRecord)
                    }
                }
            }
            finally {
                if ($null -ne $EventRecord) {
                    $EventRecord.Dispose()
                }
            }
        }

        return Complete-SecurityEventQuery
    }

    $CredentialName = Split-CredentialUserName -Credential $Credential
    $Session = $null
    $Reader = $null

    try {
        $Session = [System.Diagnostics.Eventing.Reader.EventLogSession]::new(
            $ComputerName,
            $CredentialName.Domain,
            $CredentialName.User,
            $Credential.Password,
            [System.Diagnostics.Eventing.Reader.SessionAuthentication]::Default
        )

        $XPathQuery = New-SecurityEventXPathQuery `
            -Ids $Ids `
            -ProviderName $ProviderName `
            -StartDate $StartDate

        $Query = [System.Diagnostics.Eventing.Reader.EventLogQuery]::new(
            "Security",
            [System.Diagnostics.Eventing.Reader.PathType]::LogName,
            $XPathQuery
        )
        $Query.Session = $Session
        $Query.ReverseDirection = $true

        $Reader = [System.Diagnostics.Eventing.Reader.EventLogReader]::new($Query)

        while ($true) {
            $EventRecord = $Reader.ReadEvent()
            if ($null -eq $EventRecord) {
                break
            }

            try {
                if ($MaxEvents -gt 0 -and $EventsRead -ge $MaxEvents) {
                    $LimitReached = $true
                    break
                }

                $EventsRead++
                $ProcessedOutput = if ($ProcessEvent) { & $ProcessEvent $EventRecord } else { $EventRecord }
                foreach ($ProcessedRecord in @($ProcessedOutput)) {
                    if ($null -ne $ProcessedRecord) {
                        [void]$ProcessedRecords.Add($ProcessedRecord)
                    }
                }
            }
            finally {
                if ($null -ne $EventRecord) {
                    $EventRecord.Dispose()
                }
            }
        }

        return Complete-SecurityEventQuery
    }
    finally {
        if ($null -ne $Reader) {
            $Reader.Dispose()
        }

        if ($null -ne $Session) {
            $Session.Dispose()
        }
    }
}

# --- CSV Safety ---

function ConvertTo-SafeCsvValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or -not ($Value -is [string])) {
        return $Value
    }

    if ($Value -match '^[\t\r\n]' -or $Value -match '^\s*[=+\-@]') {
        return "'$Value"
    }

    return $Value
}

function ConvertTo-SafeCsvRecord {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [psobject]$InputObject
    )

    process {
        $Properties = [ordered]@{}

        foreach ($Property in $InputObject.PSObject.Properties) {
            $Properties[$Property.Name] = ConvertTo-SafeCsvValue -Value $Property.Value
        }

        [pscustomobject]$Properties
    }
}

# --- Path and Name Validation ---

function Test-IsUncPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    try {
        $uri = [System.Uri]$Path
        return $uri.IsUnc
    }
    catch {
        return $Path.StartsWith("\\", [System.StringComparison]::Ordinal)
    }
}

function Get-ResolvedOutputPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path -Path (Get-Location).Path -ChildPath $Path)
}

function Test-IsSafeDnsName {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name) -or
        $Name.Length -gt 253 -or
        (Test-ContainsControlCharacter -Value $Name)) {
        return $false
    }

    if ($Name.StartsWith(".") -or $Name.EndsWith(".")) {
        return $false
    }

    foreach ($label in $Name.Split(".")) {
        if ([string]::IsNullOrWhiteSpace($label) -or $label.Length -gt 63) {
            return $false
        }

        if ($label -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?$') {
            return $false
        }
    }

    return $true
}

function Test-ContainsControlCharacter {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Value
    )

    return ($null -ne $Value -and $Value -match '[\x00-\x1F\x7F]')
}

function Resolve-AdminToolsPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$BaseDirectory
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $BaseDirectory -ChildPath $Path))
}

function Test-IsPathUnderDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)

    return (
        $fullPath.Equals($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullDirectory + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Get-AdminToolsSecret {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return Get-Secret -Name $Name -ErrorAction Stop
}

function Resolve-AdminToolsSecretCredential {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CredentialSecretName
    )

    $secretModule = Get-Module -ListAvailable -Name Microsoft.PowerShell.SecretManagement -ErrorAction SilentlyContinue |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

    if (-not $secretModule) {
        throw "Microsoft.PowerShell.SecretManagement is required for -CredentialSecretName. Install it and register a vault, or use -CredentialPath."
    }

    Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
    $secret = Get-AdminToolsSecret -Name $CredentialSecretName

    if ($secret -isnot [PSCredential]) {
        throw "Secret '$CredentialSecretName' must contain a PSCredential."
    }

    return $secret
}

function Resolve-AdminToolsCredential {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [PSCredential]$Credential,

        [Parameter(Mandatory = $false)]
        [string]$CredentialSecretName,

        [Parameter(Mandatory = $false)]
        [string]$CredentialPath,

        [Parameter(Mandatory = $true)]
        [string]$BaseDirectory,

        [Parameter(Mandatory = $false)]
        [switch]$AllowNetworkInputPath
    )

    $sourceCount = 0
    if ($Credential) { $sourceCount++ }
    if (-not [string]::IsNullOrWhiteSpace($CredentialSecretName)) { $sourceCount++ }
    if (-not [string]::IsNullOrWhiteSpace($CredentialPath)) { $sourceCount++ }

    if ($sourceCount -gt 1) {
        throw "Only one credential source can be used: -Credential, -CredentialSecretName, or -CredentialPath."
    }

    if ($Credential) {
        return $Credential
    }

    if (-not [string]::IsNullOrWhiteSpace($CredentialSecretName)) {
        Assert-SafeTextValues -Purpose "CredentialSecretName" -Values @($CredentialSecretName) -MaximumLength 256
        return Resolve-AdminToolsSecretCredential -CredentialSecretName $CredentialSecretName
    }

    if ([string]::IsNullOrWhiteSpace($CredentialPath)) {
        return $null
    }

    Assert-SafeTextValues -Purpose "CredentialPath" -Values @($CredentialPath)
    $resolvedCredentialPath = Resolve-AdminToolsPath -Path $CredentialPath -BaseDirectory $BaseDirectory
    $resolvedBaseDirectory = [System.IO.Path]::GetFullPath($BaseDirectory)

    if (Test-IsPathUnderDirectory -Path $resolvedCredentialPath -Directory $resolvedBaseDirectory) {
        throw "CredentialPath must point to a credential file outside the repository directory."
    }

    if ((Test-IsUncPath -Path $resolvedCredentialPath) -and -not $AllowNetworkInputPath) {
        throw "Network credential paths are not allowed by default: $resolvedCredentialPath. Re-run with -AllowNetworkInputPath only if the location is trusted."
    }

    if (-not (Test-Path -LiteralPath $resolvedCredentialPath -PathType Leaf)) {
        throw "CredentialPath file not found: $resolvedCredentialPath"
    }

    $loadedCredential = Import-Clixml -LiteralPath $resolvedCredentialPath -ErrorAction Stop
    if ($loadedCredential -isnot [PSCredential]) {
        throw "CredentialPath must contain a PSCredential exported with Export-Clixml."
    }

    return $loadedCredential
}
function Assert-SafeDomainControllerNames {
    param(
        [string[]]$Names
    )

    foreach ($Name in @($Names)) {
        if ([string]::IsNullOrWhiteSpace($Name)) {
            continue
        }

        if (-not (Test-IsSafeDnsName -Name $Name)) {
            throw "DomainControllers contains an invalid DNS name: $Name"
        }
    }
}

function Assert-SafeTextValues {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Purpose,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Values,

        [int]$MaximumLength = 4096
    )

    foreach ($Value in @($Values)) {
        if ([string]::IsNullOrWhiteSpace($Value)) {
            continue
        }

        if ($Value.Length -gt $MaximumLength) {
            throw "$Purpose value exceeds maximum length $MaximumLength."
        }

        if (Test-ContainsControlCharacter -Value $Value) {
            throw "$Purpose contains control characters and was rejected."
        }
    }
}

# Exported to satisfy the retry test scaffold; Scan-ADComputers keeps its
# logging-aware implementation because its Write-AdminToolsLog function is script-local.
function Get-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5
    )

    $PermanentPatterns = @(
        "Access is denied",
        "Insufficient access",
        "UnauthorizedAccessException",
        "invalid filter",
        "No such object",
        "The server does not support the requested critical extension"
    )

    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        try {
            return & $ScriptBlock
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($PermanentPatterns | Where-Object { $errorMessage -match $_ }) {
                throw
            }

            $attempt++
            if ($attempt -lt $MaxAttempts) {
                Start-Sleep -Seconds $DelaySeconds
            }
            else {
                throw
            }
        }
    }
}

Export-ModuleMember -Function *

