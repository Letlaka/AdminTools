#requires -Version 5.1
<#
.SYNOPSIS
    Shared utility functions for AdminTools scripts.
.DESCRIPTION
    Import this module at the top of each AdminTools script:
        Import-Module (Join-Path $PSScriptRoot "AdminToolsCommon.psm1") -Force
#>

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
        [PSCredential]$Credential
    )

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

        if ($MaxEvents -gt 0) {
            $GetWinEventParameters["MaxEvents"] = $MaxEvents
        }

        return @(Get-WinEvent @GetWinEventParameters)
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
        $Events = New-Object System.Collections.Generic.List[object]

        while ($true) {
            $Event = $Reader.ReadEvent()
            if ($null -eq $Event) {
                break
            }

            [void]$Events.Add($Event)

            if ($MaxEvents -gt 0 -and $Events.Count -ge $MaxEvents) {
                break
            }
        }

        return @($Events.ToArray())
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
# logging-aware implementation because its Write-Log function is script-local.
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
