using module ./AdminToolsCommon.psm1
<#
.SYNOPSIS
    Scan Active Directory for server or workstation computer objects and export reports.

.DESCRIPTION
    Supports three broad feature phases in one script:
    - AD inventory and reporting controls.
    - Optional operational checks such as DNS, ports, and remote inventory.
    - Usability controls such as config files, logging, and structured exit codes.

.PARAMETER ComputerType
    The AD computer type to scan. Valid values are Server and Workstation.

.PARAMETER Mode
    Full scans the selected computer type across the query scope.
    Targeted scans only the names listed in ComputerListPath.

.PARAMETER ComputerListPath
    Path to a text file containing computer names or FQDNs, one per line.
    Required when Mode is Targeted.

.PARAMETER SearchBase
    Optional distinguished name that limits the AD query scope to a specific OU or container.

.PARAMETER SearchBaseList
    Optional list of distinguished names to query. Can be combined with SearchBase.

.PARAMETER ExcludeOU
    Optional list of OU distinguished names to exclude from the final output.

.PARAMETER InactiveDays
    Optional inactivity threshold used to flag stale devices.

.PARAMETER IncludeDisabled
    Include disabled AD computer objects in the scan.

.PARAMETER ExportFormat
    One or more export formats: Csv, Json, Html.

.PARAMETER CompareWithPrevious
    Path to a previous Csv or Json inventory export. A delta report is generated.

.PARAMETER SummaryOnly
    Skip the main inventory export and export summary breakdowns instead.

.PARAMETER SeparateStatusExports
    In Targeted mode, export separate matched, unreachable, and not-found status reports.

.PARAMETER ResolveDns
    Resolve forward DNS for each exported computer and flag mismatches.

.PARAMETER TestPorts
    One or more TCP ports to test for each exported computer.

.PARAMETER RemoteInventory
    Attempt remote inventory collection for each exported computer.

.PARAMETER TimeoutSeconds
    Timeout for connectivity and remote inventory operations.

.PARAMETER ThrottleLimit
    Throttle limit used for parallel connectivity and operational checks.

.PARAMETER PingCount
    Number of ICMP echo requests to use when TestMethod is Ping.

.PARAMETER TestMethod
    Connectivity method used in Targeted mode and optional operational enrichment.
    Valid values are Ping, WinRM, and None.

.PARAMETER SkipPing
    Backward-compatible shortcut that forces TestMethod to None.

.PARAMETER ConfigPath
    Optional path to a Json configuration file.

.PARAMETER LogPath
    Optional path to the run log file.

.PARAMETER NoProgress
    Suppress transient Write-Progress feedback for unattended or redirected runs.

.PARAMETER NoClobber
    Fail if an output or log file already exists instead of overwriting it.

.PARAMETER ForceOverwrite
    Overwrite existing output or log files. Existing files are not overwritten by default.

.PARAMETER AllowNetworkOutputPath
    Allow writing reports and logs to UNC paths. Network output paths are rejected by default.

.PARAMETER AllowNetworkInputPath
    Allow reading configuration, target list, or comparison files from UNC paths.
    Network input paths are rejected by default.

.PARAMETER DisableCsvSanitization
    Export raw string values without Excel formula injection protection.

.EXAMPLE
    .\Scan-ADComputers.ps1 -ComputerType Server -Mode Full

.EXAMPLE
    .\Scan-ADComputers.ps1 -ComputerType Workstation -Mode Targeted `
        -ComputerListPath .\workstationlist.txt `
        -ResolveDns `
        -TestPorts 445,3389 `
        -InactiveDays 90

.EXAMPLE
    .\Scan-ADComputers.ps1 -ConfigPath .\Scan-ADComputers.json

.EXAMPLE
    .\Scan-ADComputers.ps1 -ComputerType Server -Mode Full -NoProgress
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("Server", "Workstation")]
    [string]$ComputerType = "Server",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Full", "Targeted")]
    [string]$Mode = "Full",

    [Parameter(Mandatory = $false)]
    [string]$DomainController,

    [Parameter(Mandatory = $false)]
    [string]$DomainName,

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$ComputerListPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $false)]
    [string]$SearchBase,

    [Parameter(Mandatory = $false)]
    [string[]]$SearchBaseList,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeOU,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabled,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Csv", "Json", "Html")]
    [string[]]$ExportFormat = @("Csv"),

    [Parameter(Mandatory = $false)]
    [string]$CompareWithPrevious,

    [Parameter(Mandatory = $false)]
    [switch]$SummaryOnly,

    [Parameter(Mandatory = $false)]
    [switch]$SeparateStatusExports,

    [Parameter(Mandatory = $false)]
    [switch]$ResolveDns,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 65535)]
    [int[]]$TestPorts,

    [Parameter(Mandatory = $false)]
    [switch]$RemoteInventory,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 300)]
    [int]$TimeoutSeconds = 5,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 128)]
    [int]$ThrottleLimit = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$PingCount = 1,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Ping", "WinRM", "None")]
    [string]$TestMethod = "Ping",

    [Parameter(Mandatory = $false)]
    [switch]$SkipPing,

    [Parameter(Mandatory = $false)]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [switch]$NoProgress,

    [Parameter(Mandatory = $false)]
    [switch]$NoClobber,

    [Parameter(Mandatory = $false)]
    [switch]$ForceOverwrite,

    [Parameter(Mandatory = $false)]
    [switch]$AllowNetworkOutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$AllowNetworkInputPath,

    [Parameter(Mandatory = $false)]
    [switch]$DisableCsvSanitization
)

Import-Module (Join-Path $PSScriptRoot "AdminToolsCommon.psm1") -Force -ErrorAction Stop

$ScriptDirectory = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$RunTimestamp = (Get-Date).ToString("yyyyMMddHHmmss")
$RunStartedAt = Get-Date

$ExitCodes = @{
    General     = 1
    Prereq      = 2
    Config      = 3
    Validation  = 4
    ADQuery     = 5
    Operational = 6
    Export      = 7
    Compare     = 8
}

$script:InitialBoundParameters = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $script:InitialBoundParameters[$key] = $PSBoundParameters[$key]
}

$script:LogFilePath = $null
$script:FileLoggingEnabled = $false

if ($ForceOverwrite -and $NoClobber) {
    Write-Error "ForceOverwrite and NoClobber cannot be used together."
    exit $ExitCodes.Validation
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error", "Verbose")]
        [string]$Level = "Info"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[{0}] [{1}] {2}" -f $timestamp, $Level.ToUpperInvariant(), $Message

    switch ($Level) {
        "Info" { Write-Host $Message }
        "Warning" { Write-Warning $Message }
        "Error" { Write-Error $Message }
        "Verbose" { Write-Verbose $Message }
    }

    if ($script:FileLoggingEnabled -and -not [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
        Add-Content -LiteralPath $script:LogFilePath -Value $entry
    }
}

function Write-ErrorAndExit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("General", "Prereq", "Config", "Validation", "ADQuery", "Operational", "Export", "Compare")]
        [string]$CodeKey = "General"
    )

    Write-Log -Message $Message -Level Error
    exit $ExitCodes[$CodeKey]
}

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$BaseDirectory
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    $currentLocationCandidate = Join-Path (Get-Location).Path $Path
    if (Test-Path -LiteralPath $currentLocationCandidate) {
        return (Resolve-Path -LiteralPath $currentLocationCandidate).Path
    }

    return (Join-Path $BaseDirectory $Path)
}

function Resolve-OutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$BaseDirectory
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return (Join-Path $BaseDirectory $Path)
}

function Get-SafeFileNamePart {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "domain"
    }

    return ($Value -replace '[^a-zA-Z0-9._-]', '_')
}

function Test-IsInDnsSuffix {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$DnsSuffix
    )

    return (
        $Name.Equals($DnsSuffix, [System.StringComparison]::OrdinalIgnoreCase) -or
        $Name.EndsWith(".$DnsSuffix", [System.StringComparison]::OrdinalIgnoreCase)
    )
}

function Assert-SafeDnsName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Purpose
    )

    if (-not (Test-IsSafeDnsName -Name $Name)) {
        Write-ErrorAndExit -Message "$Purpose contains invalid DNS name characters or length: $Name" -CodeKey Validation
    }
}

function Assert-AllowedValues {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Purpose,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string[]]$Values,

        [Parameter(Mandatory = $true)]
        [string[]]$AllowedValues
    )

    $AllowedSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($AllowedValue in $AllowedValues) {
        [void]$AllowedSet.Add($AllowedValue)
    }

    foreach ($Value in @($Values)) {
        if ([string]::IsNullOrWhiteSpace($Value) -or $AllowedSet.Contains($Value)) {
            continue
        }

        Write-ErrorAndExit -Message "$Purpose contains unsupported value '$Value'. Allowed values: $($AllowedValues -join ', ')." -CodeKey Validation
    }
}

function Assert-IntegerRange {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Purpose,

        [Parameter(Mandatory = $true)]
        [int]$Value,

        [Parameter(Mandatory = $true)]
        [int]$Minimum,

        [Parameter(Mandatory = $true)]
        [int]$Maximum
    )

    if ($Value -lt $Minimum -or $Value -gt $Maximum) {
        Write-ErrorAndExit -Message "$Purpose must be between $Minimum and $Maximum." -CodeKey Validation
    }
}

function Assert-OutputPathAllowed {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Purpose
    )

    if ((Test-IsUncPath -Path $Path) -and -not $AllowNetworkOutputPath) {
        Write-ErrorAndExit -Message "Network $Purpose paths are not allowed by default: $Path. Re-run with -AllowNetworkOutputPath only if the location is trusted." -CodeKey Validation
    }

    if ((Test-Path -LiteralPath $Path) -and ($NoClobber -or -not $ForceOverwrite)) {
        Write-ErrorAndExit -Message "$Purpose path already exists: $Path. Re-run with -ForceOverwrite only if replacing it is intended." -CodeKey Validation
    }

    if ((Test-Path -LiteralPath $Path) -and $ForceOverwrite) {
        Write-Warning "Overwriting existing $Purpose path: $Path"
    }
}

function Assert-InputPathAllowed {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Purpose
    )

    if ((Test-IsUncPath -Path $Path) -and -not $AllowNetworkInputPath) {
        Write-ErrorAndExit -Message "Network $Purpose paths are not allowed by default: $Path. Re-run with -AllowNetworkInputPath only if the location is trusted." -CodeKey Validation
    }
}

function Import-TrustedActiveDirectoryModule {
    $loadedModules = @(Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)
    if ($loadedModules.Count -gt 0) {
        foreach ($loadedModule in $loadedModules) {
            if (-not (Test-IsTrustedActiveDirectoryModulePath -Path $loadedModule.ModuleBase)) {
                Write-ErrorAndExit -Message "ActiveDirectory module is already loaded from an untrusted path: $($loadedModule.ModuleBase)" -CodeKey Prereq
            }
        }

        return
    }

    $trustedModule = @(Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue |
            Where-Object { Test-IsTrustedActiveDirectoryModulePath -Path $_.ModuleBase } |
            Sort-Object -Property Version -Descending |
            Select-Object -First 1)

    if ($trustedModule.Count -eq 0) {
        Write-ErrorAndExit -Message '"ActiveDirectory" module not found in the trusted Windows RSAT module path. Install RSAT Active Directory Tools on this machine.' -CodeKey Prereq
    }

    try {
        $moduleSpecification = @{
            ModuleName    = "ActiveDirectory"
            ModuleVersion = $trustedModule[0].Version
        }

        if ($trustedModule[0].Guid -and $trustedModule[0].Guid -ne [guid]::Empty) {
            $moduleSpecification["Guid"] = $trustedModule[0].Guid
        }

        Import-Module -FullyQualifiedName $moduleSpecification -ErrorAction Stop
    }
    catch {
        Write-ErrorAndExit -Message "ActiveDirectory module could not be loaded from trusted path '$($trustedModule[0].ModuleBase)'. Error: $($_.Exception.Message)" -CodeKey Prereq
    }
}

$script:ConfigSourcedParameters = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

function Set-ParameterFromConfig {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$ConfigObject,

        [Parameter(Mandatory = $true)]
        [string]$ParameterName
    )

    if ($script:InitialBoundParameters.ContainsKey($ParameterName)) {
        return
    }

    $property = $ConfigObject.PSObject.Properties[$ParameterName]
    if ($null -eq $property -or $null -eq $property.Value) {
        return
    }

    Set-Variable -Name $ParameterName -Value $property.Value -Scope Script
    [void]$script:ConfigSourcedParameters.Add($ParameterName)
}

function Write-ScanProgress {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Activity,

        [Parameter(Mandatory = $true)]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [int]$CompletedCount = 0,

        [Parameter(Mandatory = $false)]
        [int]$TotalCount = 0
    )

    if ($NoProgress) {
        return
    }

    $progressParameters = @{
        Id       = $Id
        Activity = $Activity
        Status   = $Status
    }

    if ($TotalCount -gt 0) {
        $boundedCompleted = [Math]::Min([Math]::Max($CompletedCount, 0), $TotalCount)
        $progressParameters["PercentComplete"] = [int][Math]::Floor(($boundedCompleted / $TotalCount) * 100)
    }

    Write-Progress @progressParameters
}

function Complete-ScanProgress {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Activity
    )

    if ($NoProgress) {
        return
    }

    Write-Progress -Id $Id -Activity $Activity -Completed
}

function Initialize-LogFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($WhatIfPreference) {
        Write-Host "WhatIf: would initialize log file -> $Path"
        return
    }

    Assert-OutputPathAllowed -Path $Path -Purpose "log"

    $logDirectory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force -ErrorAction Stop | Out-Null
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path -Force -ErrorAction Stop | Out-Null
    }

    $script:LogFilePath = $Path
    $script:FileLoggingEnabled = $true
}

function Get-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [int]$MaxAttempts = 3,

        [int]$DelaySeconds = 5
    )

    # These patterns indicate transient conditions worth retrying.
    # Authorization, filter, and object-not-found errors are permanent
    # and are re-thrown immediately.
    $TransientPatterns = @(
        "timeout",
        "RPC",
        "network",
        "busy",
        "temporarily unavailable",
        "server is unavailable"
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

            $isPermanent = $PermanentPatterns | Where-Object { $errorMessage -match $_ }
            if ($isPermanent) {
                Write-Log -Message ("Permanent failure, not retrying: {0}" -f $errorMessage) -Level Warning
                throw
            }

            $attempt++
            Write-Log -Message ("Attempt {0} failed: {1}" -f $attempt, $errorMessage) -Level Warning

            if ($attempt -lt $MaxAttempts) {
                $isTransient = $TransientPatterns | Where-Object { $errorMessage -match $_ }
                if (-not $isTransient) {
                    Write-Log -Message "Failure does not match known transient patterns; retrying anyway." -Level Warning
                }
                Start-Sleep -Seconds $DelaySeconds
            }
            else {
                throw
            }
        }
    }
}

function Get-ComputerTypeFilter {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Server", "Workstation")]
        [string]$Type
    )

    switch ($Type) {
        "Server" {
            return 'OperatingSystem -like "*Server*"'
        }
        "Workstation" {
            return 'OperatingSystem -like "Windows*" -and OperatingSystem -notlike "*Server*"'
        }
    }
}

function Get-ComputerExportPrefix {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Server", "Workstation")]
        [string]$Type
    )

    switch ($Type) {
        "Server" { "Servers" }
        "Workstation" { "Workstations" }
    }
}

function Convert-ToEscapedAdFilterValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return ($Value -replace "'", "''")
}

function Get-QuerySearchBases {
    $bases = New-Object 'System.Collections.Generic.List[string]'

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $bases.Add($SearchBase.Trim())
    }

    foreach ($searchBaseEntry in @($SearchBaseList)) {
        if (-not [string]::IsNullOrWhiteSpace($searchBaseEntry)) {
            $bases.Add($searchBaseEntry.Trim())
        }
    }

    return @($bases | Sort-Object -Unique)
}

function Test-IsExcludedByOu {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$DistinguishedName,

        [Parameter(Mandatory = $false)]
        [string[]]$ExcludedOuList
    )

    if ([string]::IsNullOrWhiteSpace($DistinguishedName) -or @($ExcludedOuList).Count -eq 0) {
        return $false
    }

    $normalizedDn = $DistinguishedName.Trim().ToUpperInvariant()

    foreach ($excludedOu in @($ExcludedOuList)) {
        if ([string]::IsNullOrWhiteSpace($excludedOu)) {
            continue
        }

        $normalizedOu = $excludedOu.Trim().ToUpperInvariant()
        if ($normalizedDn -eq $normalizedOu -or $normalizedDn.EndsWith(",$normalizedOu")) {
            return $true
        }
    }

    return $false
}

function Get-RequestedComputerList {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListPath,

        [Parameter(Mandatory = $true)]
        [string]$DefaultDomainName
    )

    $seenInputs = @{}
    $requestedComputers = New-Object System.Collections.Generic.List[object]

    foreach ($line in Get-Content -LiteralPath $ListPath) {
        $inputName = $line.Trim()

        if ([string]::IsNullOrWhiteSpace($inputName) -or $inputName -match '^#') {
            continue
        }

        $dedupeKey = $inputName.ToUpperInvariant()
        if ($seenInputs.ContainsKey($dedupeKey)) {
            continue
        }

        $seenInputs[$dedupeKey] = $true

        if ($inputName -match '\.') {
            if (-not (Test-IsSafeDnsName -Name $inputName)) {
                Write-ErrorAndExit -Message "ComputerListPath contains an invalid DNS name: $inputName" -CodeKey Validation
            }

            if (-not (Test-IsInDnsSuffix -Name $inputName -DnsSuffix $DefaultDomainName)) {
                Write-ErrorAndExit -Message "ComputerListPath target '$inputName' is outside the allowed AD DNS suffix '$DefaultDomainName'." -CodeKey Validation
            }

            $shortName = $inputName.Split('.')[0].ToUpperInvariant()
            $fqdn = $inputName
        }
        else {
            if ($inputName -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$') {
                Write-ErrorAndExit -Message "ComputerListPath contains an invalid short computer name: $inputName" -CodeKey Validation
            }

            $shortName = $inputName.ToUpperInvariant()
            $fqdn = "$inputName.$DefaultDomainName"
        }

        $requestedComputers.Add([PSCustomObject]@{
                InputName = $inputName
                ShortName = $shortName
                FQDN      = $fqdn
            })
    }

    return @($requestedComputers.ToArray())
}

function Get-RequestedComputerFilters {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$RequestedComputers,

        [Parameter(Mandatory = $true)]
        [string]$BaseFilter,

        [int]$ChunkSize = 40
    )

    $seenClauses = @{}
    $clauses = New-Object 'System.Collections.Generic.List[string]'

    foreach ($requestedComputer in $RequestedComputers) {
        foreach ($candidate in @(
                @{ Attribute = "Name"; Value = $requestedComputer.ShortName },
                @{ Attribute = "DNSHostName"; Value = $requestedComputer.FQDN }
            )) {
            if ([string]::IsNullOrWhiteSpace($candidate.Value)) {
                continue
            }

            $escapedValue = Convert-ToEscapedAdFilterValue -Value $candidate.Value
            $clause = "{0} -eq '{1}'" -f $candidate.Attribute, $escapedValue

            if (-not $seenClauses.ContainsKey($clause)) {
                $seenClauses[$clause] = $true
                $clauses.Add($clause)
            }
        }
    }

    if ($clauses.Count -eq 0) {
        return @()
    }

    $filters = New-Object 'System.Collections.Generic.List[string]'

    for ($offset = 0; $offset -lt $clauses.Count; $offset += $ChunkSize) {
        $batchSize = [Math]::Min($ChunkSize, $clauses.Count - $offset)
        $batchClauses = $clauses.GetRange($offset, $batchSize)
        $nameFilter = [string]::Join(' -or ', $batchClauses)
        $filters.Add("($BaseFilter) -and ($nameFilter)")
    }

    return @($filters.ToArray())
}

function Get-ComputerIdentityKey {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$InputObject
    )

    foreach ($propertyName in @("ObjectGUID", "objectGUID")) {
        if ($InputObject.PSObject.Properties[$propertyName]) {
            $value = $InputObject.$propertyName
            if ($null -ne $value) {
                if ($value -is [guid]) {
                    return $value.Guid
                }

                if ($value.PSObject -and $value.PSObject.Properties["Guid"]) {
                    return $value.Guid
                }

                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    return ([string]$value).Trim()
                }
            }
        }
    }

    foreach ($propertyName in @("DNSHostName", "dNSHostName", "DistinguishedName", "distinguishedName", "CN", "cn", "Name", "name")) {
        if ($InputObject.PSObject.Properties[$propertyName]) {
            $value = [string]$InputObject.$propertyName
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value.Trim().ToUpperInvariant()
            }
        }
    }

    return $null
}

function Add-ComputerLookupEntry {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Lookup,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [psobject]$Computer
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return
    }

    $normalizedKey = $Key.Trim().ToUpperInvariant()
    if (-not $Lookup.ContainsKey($normalizedKey)) {
        $Lookup[$normalizedKey] = $Computer
    }
}

function Get-EffectiveLastSeenDate {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Computer
    )

    $candidates = New-Object System.Collections.Generic.List[datetime]

    foreach ($propertyName in @("lastLogonDate", "LastLogonDate", "lastLogonTimestamp", "LastLogonTimestamp")) {
        if (-not $Computer.PSObject.Properties[$propertyName]) {
            continue
        }

        $value = $Computer.$propertyName
        if ($null -eq $value) {
            continue
        }

        if ($value -is [datetime]) {
            $candidates.Add($value)
            continue
        }

        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParse([string]$value, [ref]$parsedDate)) {
            $candidates.Add($parsedDate)
        }
    }

    if ($candidates.Count -eq 0) {
        return $null
    }

    return ($candidates | Sort-Object -Descending | Select-Object -First 1)
}

function Get-OrganizationalUnitPath {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Computer
    )

    if ($Computer.PSObject.Properties["canonicalName"] -and -not [string]::IsNullOrWhiteSpace([string]$Computer.canonicalName)) {
        $canonicalName = [string]$Computer.canonicalName
        $segments = $canonicalName.Split('/')
        if ($segments.Count -gt 1) {
            return ($segments[0..($segments.Count - 2)] -join '/')
        }
    }

    if ($Computer.PSObject.Properties["distinguishedName"] -and -not [string]::IsNullOrWhiteSpace([string]$Computer.distinguishedName)) {
        $dn = [string]$Computer.distinguishedName
        $parts = $dn.Split(',')
        if ($parts.Count -gt 1) {
            return ($parts[1..($parts.Count - 1)] -join ',')
        }
    }

    return $null
}

function Get-StaleStatus {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [Nullable[datetime]]$LastSeenDate,

        [Parameter(Mandatory = $false)]
        [int]$ThresholdDays = 0
    )

    if ($ThresholdDays -le 0) {
        return "NotEvaluated"
    }

    if ($null -eq $LastSeenDate) {
        return "Unknown"
    }

    $daysSinceLastSeen = [Math]::Floor(((Get-Date) - $LastSeenDate).TotalDays)
    if ($daysSinceLastSeen -ge $ThresholdDays) {
        return "Stale"
    }

    return "Active"
}

function Get-ComputerRecord {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Computer,

        [Parameter(Mandatory = $true)]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [int]$InactiveThresholdDays = 0
    )

    $lastSeenDate = Get-EffectiveLastSeenDate -Computer $Computer
    $daysSinceLastSeen = $null

    if ($null -ne $lastSeenDate) {
        $daysSinceLastSeen = [Math]::Floor(((Get-Date) - $lastSeenDate).TotalDays)
    }

    $staleStatus = Get-StaleStatus -LastSeenDate $lastSeenDate -ThresholdDays $InactiveThresholdDays
    $isStale = if ($InactiveThresholdDays -gt 0) { $staleStatus -eq "Stale" } else { $null }

    [PSCustomObject]@{
        ComputerType           = $Type
        Name                   = $Computer.name
        CN                     = $Computer.cn
        DNSHostName            = $Computer.dNSHostName
        Description            = $Computer.description
        OUPath                 = (Get-OrganizationalUnitPath -Computer $Computer)
        CanonicalName          = $Computer.canonicalName
        DistinguishedName      = $Computer.distinguishedName
        Created                = $Computer.whenCreated
        Enabled                = $Computer.enabled
        IPv4Address            = (@($Computer.ipv4Address) | Where-Object { $_ }) -join ","
        LastLogonDate          = $Computer.lastLogonDate
        LastLogonTimestamp     = $Computer.lastLogonTimestamp
        LastSeenDate           = $lastSeenDate
        DaysSinceLastSeen      = $daysSinceLastSeen
        InactiveThresholdDays  = if ($InactiveThresholdDays -gt 0) { $InactiveThresholdDays } else { $null }
        IsStale                = $isStale
        StaleStatus            = $staleStatus
        ObjectGUID             = $Computer.objectGUID
        OperatingSystem        = $Computer.operatingSystem
        OperatingSystemVersion = $Computer.operatingSystemVersion
        ConnectivityMethod     = $null
        ConnectivityStatus     = $null
        ConnectivityReachable  = $null
        ConnectivityDetail     = $null
        DnsStatus              = $null
        DnsResolvedIPs         = $null
        DnsMatchesAdIPv4       = $null
        PortStatus             = $null
        RemoteInventoryStatus  = $null
        RemoteUptimeDays       = $null
        SerialNumber           = $null
        Model                  = $null
        TotalMemoryGB          = $null
        SystemDriveFreeGB      = $null
        PendingReboot          = $null
        LoggedOnUser           = $null
        QueriedAt              = $RunStartedAt
    }
}

function Export-DataSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string[]]$Formats,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object[]]$Data,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $rows = @($Data)
    $exportedPaths = New-Object 'System.Collections.Generic.List[string]'
    $formatList = @($Formats)
    $exportIndex = 0
    $progressActivity = "Exporting report: $Title"

    try {
        foreach ($format in $formatList) {
            Write-ScanProgress -Id 6 -Activity $progressActivity `
                -Status ("Exporting {0} ({1} of {2})" -f $format, ($exportIndex + 1), $formatList.Count) `
                -CompletedCount $exportIndex -TotalCount $formatList.Count

            $lowerFormat = $format.ToLowerInvariant()
            $targetPath = "{0}.{1}" -f $BasePath, $lowerFormat

            if ($WhatIfPreference) {
                Write-Host "WhatIf: would export $Title -> $targetPath"
                $exportIndex++
                Write-ScanProgress -Id 6 -Activity $progressActivity `
                    -Status ("Completed {0} of {1} format(s)" -f $exportIndex, $formatList.Count) `
                    -CompletedCount $exportIndex -TotalCount $formatList.Count
                continue
            }

            Assert-OutputPathAllowed -Path $targetPath -Purpose "output"

            switch ($format) {
                "Csv" {
                    if ($rows.Count -gt 0) {
                        $csvRows = if ($DisableCsvSanitization) {
                            Write-Warning "CSV sanitization is disabled. Opening this report in spreadsheet software may evaluate AD data as formulas."
                            $rows
                        }
                        else {
                            @($rows | ConvertTo-SafeCsvRecord)
                        }

                        $csvRows | Export-Csv -LiteralPath $targetPath -NoTypeInformation
                    }
                    else {
                        Set-Content -LiteralPath $targetPath -Value "" -Encoding UTF8
                    }
                }
                "Json" {
                    $json = if ($rows.Count -gt 0) {
                        $rows | ConvertTo-Json -Depth 8
                    }
                    else {
                        "[]"
                    }

                    Set-Content -LiteralPath $targetPath -Value $json -Encoding UTF8
                }
                "Html" {
                    $style = @"
<style>
body { font-family: Segoe UI, sans-serif; margin: 24px; color: #1f2937; }
h1 { margin-bottom: 4px; }
p.meta { color: #6b7280; margin-top: 0; }
table { border-collapse: collapse; width: 100%; font-size: 12px; }
th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background: #f3f4f6; }
tr:nth-child(even) { background: #f9fafb; }
</style>
"@

                    $body = if ($rows.Count -gt 0) {
                        $rows | ConvertTo-Html -Head $style -Title $Title -PreContent "<h1>$Title</h1><p class='meta'>Generated $RunStartedAt</p>"
                    }
                    else {
                        @"
<html>
<head>$style<title>$Title</title></head>
<body>
<h1>$Title</h1>
<p class='meta'>Generated $RunStartedAt</p>
<p>No records found.</p>
</body>
</html>
"@
                    }

                    Set-Content -LiteralPath $targetPath -Value $body -Encoding UTF8
                }
            }

            Write-Log -Message ("Exported {0}: {1}" -f $Title, $targetPath) -Level Info
            $exportedPaths.Add($targetPath)
            $exportIndex++
            Write-ScanProgress -Id 6 -Activity $progressActivity `
                -Status ("Completed {0} of {1} format(s)" -f $exportIndex, $formatList.Count) `
                -CompletedCount $exportIndex -TotalCount $formatList.Count
        }
    }
    finally {
        Complete-ScanProgress -Id 6 -Activity $progressActivity
    }

    return @($exportedPaths.ToArray())
}

function Import-PreviousDataSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    switch ($extension) {
        ".csv" {
            return @(Import-Csv -LiteralPath $Path)
        }
        ".json" {
            $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($json)) {
                return @()
            }

            return @(ConvertFrom-Json -InputObject $json -ErrorAction Stop)
        }
        default {
            Write-ErrorAndExit -Message "CompareWithPrevious supports only Csv and Json files." -CodeKey Compare
        }
    }
}

function Convert-ToNullableBool {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [bool]) {
        return $Value
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = $false
    if ([bool]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Get-DeltaRecords {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$PreviousRecords,

        [Parameter(Mandatory = $true)]
        [object[]]$CurrentRecords
    )

    $previousLookup = @{}
    foreach ($previousRecord in @($PreviousRecords)) {
        $key = Get-ComputerIdentityKey -InputObject $previousRecord
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $previousLookup.ContainsKey($key)) {
            $previousLookup[$key] = $previousRecord
        }
    }

    $currentLookup = @{}
    foreach ($currentRecord in @($CurrentRecords)) {
        $key = Get-ComputerIdentityKey -InputObject $currentRecord
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $currentLookup.ContainsKey($key)) {
            $currentLookup[$key] = $currentRecord
        }
    }

    $changes = New-Object System.Collections.Generic.List[object]

    foreach ($key in $currentLookup.Keys) {
        $currentRecord = $currentLookup[$key]

        if (-not $previousLookup.ContainsKey($key)) {
            $changes.Add([PSCustomObject]@{
                    ChangeType              = "Added"
                    IdentityKey             = $key
                    CN                      = $currentRecord.CN
                    DNSHostName             = $currentRecord.DNSHostName
                    PreviousEnabled         = $null
                    CurrentEnabled          = $currentRecord.Enabled
                    PreviousOperatingSystem = $null
                    CurrentOperatingSystem  = $currentRecord.OperatingSystem
                    Details                 = "Present in current inventory only."
                })
            continue
        }

        $previousRecord = $previousLookup[$key]
        $previousEnabled = Convert-ToNullableBool -Value $previousRecord.Enabled
        $currentEnabled = Convert-ToNullableBool -Value $currentRecord.Enabled

        if ($previousEnabled -ne $null -and $currentEnabled -ne $null -and $previousEnabled -ne $currentEnabled) {
            $changes.Add([PSCustomObject]@{
                    ChangeType              = if ($currentEnabled) { "ReEnabled" } else { "Disabled" }
                    IdentityKey             = $key
                    CN                      = $currentRecord.CN
                    DNSHostName             = $currentRecord.DNSHostName
                    PreviousEnabled         = $previousEnabled
                    CurrentEnabled          = $currentEnabled
                    PreviousOperatingSystem = $previousRecord.OperatingSystem
                    CurrentOperatingSystem  = $currentRecord.OperatingSystem
                    Details                 = "Enabled state changed."
                })
        }

        $previousOperatingSystem = [string]$previousRecord.OperatingSystem
        $currentOperatingSystem = [string]$currentRecord.OperatingSystem
        $previousOperatingSystemVersion = [string]$previousRecord.OperatingSystemVersion
        $currentOperatingSystemVersion = [string]$currentRecord.OperatingSystemVersion

        if ($previousOperatingSystem -ne $currentOperatingSystem -or $previousOperatingSystemVersion -ne $currentOperatingSystemVersion) {
            $changes.Add([PSCustomObject]@{
                    ChangeType              = "OperatingSystemChanged"
                    IdentityKey             = $key
                    CN                      = $currentRecord.CN
                    DNSHostName             = $currentRecord.DNSHostName
                    PreviousEnabled         = $previousEnabled
                    CurrentEnabled          = $currentEnabled
                    PreviousOperatingSystem = $previousOperatingSystem
                    CurrentOperatingSystem  = $currentOperatingSystem
                    Details                 = "OS changed from '$previousOperatingSystem $previousOperatingSystemVersion' to '$currentOperatingSystem $currentOperatingSystemVersion'."
                })
        }
    }

    foreach ($key in $previousLookup.Keys) {
        if ($currentLookup.ContainsKey($key)) {
            continue
        }

        $previousRecord = $previousLookup[$key]
        $changes.Add([PSCustomObject]@{
                ChangeType              = "Removed"
                IdentityKey             = $key
                CN                      = $previousRecord.CN
                DNSHostName             = $previousRecord.DNSHostName
                PreviousEnabled         = $previousRecord.Enabled
                CurrentEnabled          = $null
                PreviousOperatingSystem = $previousRecord.OperatingSystem
                CurrentOperatingSystem  = $null
                Details                 = "Present in previous inventory only."
            })
    }

    return @($changes.ToArray())
}

function Get-SummaryRecords {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Records
    )

    $summary = New-Object System.Collections.Generic.List[object]

    foreach ($group in @($Records | Group-Object -Property OUPath | Sort-Object Name)) {
        $summary.Add([PSCustomObject]@{
                Category = "OU"
                Value    = if ([string]::IsNullOrWhiteSpace([string]$group.Name)) { "<blank>" } else { $group.Name }
                Count    = $group.Count
            })
    }

    foreach ($group in @($Records | Group-Object -Property OperatingSystem | Sort-Object Name)) {
        $summary.Add([PSCustomObject]@{
                Category = "OperatingSystem"
                Value    = if ([string]::IsNullOrWhiteSpace([string]$group.Name)) { "<blank>" } else { $group.Name }
                Count    = $group.Count
            })
    }

    foreach ($group in @($Records | Group-Object -Property Enabled | Sort-Object Name)) {
        $summary.Add([PSCustomObject]@{
                Category = "Enabled"
                Value    = [string]$group.Name
                Count    = $group.Count
            })
    }

    foreach ($group in @($Records | Group-Object -Property StaleStatus | Sort-Object Name)) {
        $summary.Add([PSCustomObject]@{
                Category = "StaleStatus"
                Value    = [string]$group.Name
                Count    = $group.Count
            })
    }

    return @($summary.ToArray())
}

function Get-ConnectivityLookup {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$TargetItems,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Ping", "WinRM", "None")]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [int]$TimeoutSecondsValue,

        [Parameter(Mandatory = $true)]
        [int]$PingCountValue,

        [Parameter(Mandatory = $true)]
        [int]$ThrottleLimitValue
    )

    if (@($TargetItems).Count -eq 0) {
        return @{}
    }

    $completedTargets = 0
    $progressActivity = "Testing targeted connectivity"

    try {
        $results = @($TargetItems | ForEach-Object -Parallel {
            function Test-TcpPort {
                param(
                    [string]$ComputerName,
                    [int]$Port,
                    [int]$TimeoutSeconds
                )

                $client = [System.Net.Sockets.TcpClient]::new()

                try {
                    $task = $client.ConnectAsync($ComputerName, $Port)
                    if (-not $task.Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
                        return $false
                    }

                    return $client.Connected
                }
                catch {
                    return $false
                }
                finally {
                    $client.Dispose()
                }
            }

            $item = $_
            $reachable = $null
            $status = "NotRequested"
            $detail = $null

            switch ($using:Method) {
                "None" {
                    $reachable = $true
                    $status = "Skipped"
                    $detail = "Connectivity check skipped."
                }
                "Ping" {
                    try {
                        $pingResult = Test-Connection -TargetName $item.TargetName -Count $using:PingCountValue -Quiet -TimeoutSeconds $using:TimeoutSecondsValue -ErrorAction SilentlyContinue
                        if ($pingResult -is [System.Array]) {
                            $reachable = ($pingResult -contains $true)
                        }
                        else {
                            $reachable = [bool]$pingResult
                        }
                    }
                    catch {
                        $reachable = $false
                    }

                    $status = if ($reachable) { "Reachable" } else { "Unreachable" }
                    $detail = "ICMP ping check."
                }
                "WinRM" {
                    $openPorts = New-Object System.Collections.Generic.List[string]
                    foreach ($port in 5985, 5986) {
                        if (Test-TcpPort -ComputerName $item.TargetName -Port $port -TimeoutSeconds $using:TimeoutSecondsValue) {
                            $openPorts.Add([string]$port)
                        }
                    }

                    $reachable = $openPorts.Count -gt 0
                    $status = if ($reachable) { "Reachable" } else { "Unreachable" }
                    $detail = if ($reachable) { "Open WinRM ports: $([string]::Join(',', $openPorts))" } else { "WinRM ports 5985/5986 unavailable." }
                }
            }

            [PSCustomObject]@{
                LookupKey = $item.LookupKey
                Method    = $using:Method
                Reachable = $reachable
                Status    = $status
                Detail    = $detail
            }
            } -ThrottleLimit $ThrottleLimitValue | ForEach-Object {
                $completedTargets++
                Write-ScanProgress -Id 2 -Activity $progressActivity `
                    -Status ("Completed {0} of {1} target(s)" -f $completedTargets, $TargetItems.Count) `
                    -CompletedCount $completedTargets -TotalCount $TargetItems.Count
                $_
            })
    }
    finally {
        Complete-ScanProgress -Id 2 -Activity $progressActivity
    }

    $lookup = @{}
    foreach ($result in $results) {
        $lookup[$result.LookupKey] = $result
    }

    return $lookup
}

function Invoke-OperationalEnrichment {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Records,

        [Parameter(Mandatory = $true)]
        [bool]$PerformConnectivityCheck
    )

    if (@($Records).Count -eq 0) {
        return @()
    }

    $needsEnrichment = $PerformConnectivityCheck -or $ResolveDns.IsPresent -or @($TestPorts).Count -gt 0 -or $RemoteInventory.IsPresent
    if (-not $needsEnrichment) {
        return @($Records)
    }

    $indexedRecords = for ($index = 0; $index -lt $Records.Count; $index++) {
        [PSCustomObject]@{
            Index  = $index
            Record = $Records[$index]
        }
    }

    $completedRecords = 0
    $progressActivity = "Running operational enrichment"

    try {
        $results = @($indexedRecords | ForEach-Object -Parallel {
            function Test-TcpPort {
                param(
                    [string]$ComputerName,
                    [int]$Port,
                    [int]$TimeoutSeconds
                )

                $client = [System.Net.Sockets.TcpClient]::new()

                try {
                    $task = $client.ConnectAsync($ComputerName, $Port)
                    if (-not $task.Wait([TimeSpan]::FromSeconds($TimeoutSeconds))) {
                        return $false
                    }

                    return $client.Connected
                }
                catch {
                    return $false
                }
                finally {
                    $client.Dispose()
                }
            }

            function Get-RemotePendingReboot {
                param(
                    [Microsoft.Management.Infrastructure.CimSession]$CimSession,
                    [int]$TimeoutSeconds
                )

                $hklm = [uint32]2147483650
                $probeFailures = 0

                $componentResult = $null
                try {
                    $componentResult = Invoke-CimMethod -CimSession $CimSession -Namespace "root/default" `
                        -ClassName "StdRegProv" -MethodName "EnumKey" `
                        -Arguments @{ hDefKey = $hklm; sSubKeyName = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Component Based Servicing" } `
                        -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop
                } catch { $probeFailures++ }

                $componentPending = $componentResult -and @($componentResult.sNames) -contains "RebootPending"

                $windowsUpdateResult = $null
                try {
                    $windowsUpdateResult = Invoke-CimMethod -CimSession $CimSession -Namespace "root/default" `
                        -ClassName "StdRegProv" -MethodName "EnumKey" `
                        -Arguments @{ hDefKey = $hklm; sSubKeyName = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\Auto Update" } `
                        -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop
                } catch { $probeFailures++ }

                $windowsUpdatePending = $windowsUpdateResult -and @($windowsUpdateResult.sNames) -contains "RebootRequired"

                $renameResult = $null
                try {
                    $renameResult = Invoke-CimMethod -CimSession $CimSession -Namespace "root/default" `
                        -ClassName "StdRegProv" -MethodName "GetMultiStringValue" `
                        -Arguments @{ hDefKey = $hklm; sSubKeyName = "SYSTEM\\CurrentControlSet\\Control\\Session Manager"; sValueName = "PendingFileRenameOperations" } `
                        -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop
                } catch { $probeFailures++ }

                $pendingRename = $renameResult -and @($renameResult.sValue).Count -gt 0

                # If any probe failed, return $null to distinguish from a confirmed
                # "no reboot pending" result. The caller maps $null to "Unknown".
                if ($probeFailures -gt 0) {
                    return $null
                }

                return ($componentPending -or $windowsUpdatePending -or $pendingRename)
            }

            $wrapper = $_
            $record = $wrapper.Record
            $resultProperties = [ordered]@{}

            foreach ($property in $record.PSObject.Properties) {
                $resultProperties[$property.Name] = $property.Value
            }

            $targetName = if (-not [string]::IsNullOrWhiteSpace([string]$record.DNSHostName)) {
                [string]$record.DNSHostName
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$record.CN) -and -not [string]::IsNullOrWhiteSpace([string]$using:DomainName)) {
                "{0}.{1}" -f $record.CN, $using:DomainName
            }
            else {
                [string]$record.CN
            }

            if ($using:PerformConnectivityCheck) {
                $resultProperties["ConnectivityMethod"] = $using:TestMethod
                $resultProperties["ConnectivityStatus"] = "NotRequested"
                $resultProperties["ConnectivityReachable"] = $null
                $resultProperties["ConnectivityDetail"] = $null

                switch ($using:TestMethod) {
                    "Ping" {
                        try {
                            $pingResult = Test-Connection -TargetName $targetName -Count $using:PingCount -Quiet -TimeoutSeconds $using:TimeoutSeconds -ErrorAction SilentlyContinue
                            if ($pingResult -is [System.Array]) {
                                $reachable = ($pingResult -contains $true)
                            }
                            else {
                                $reachable = [bool]$pingResult
                            }
                        }
                        catch {
                            $reachable = $false
                        }

                        $resultProperties["ConnectivityReachable"] = $reachable
                        $resultProperties["ConnectivityStatus"] = if ($reachable) { "Reachable" } else { "Unreachable" }
                        $resultProperties["ConnectivityDetail"] = "ICMP ping check."
                    }
                    "WinRM" {
                        $openPorts = New-Object System.Collections.Generic.List[string]
                        foreach ($winRmPort in 5985, 5986) {
                            if (Test-TcpPort -ComputerName $targetName -Port $winRmPort -TimeoutSeconds $using:TimeoutSeconds) {
                                $openPorts.Add([string]$winRmPort)
                            }
                        }

                        $reachable = $openPorts.Count -gt 0
                        $resultProperties["ConnectivityReachable"] = $reachable
                        $resultProperties["ConnectivityStatus"] = if ($reachable) { "Reachable" } else { "Unreachable" }
                        $resultProperties["ConnectivityDetail"] = if ($reachable) { "Open WinRM ports: $([string]::Join(',', $openPorts))" } else { "WinRM ports 5985/5986 unavailable." }
                    }
                }
            }

            if ($using:ResolveDns) {
                if ([string]::IsNullOrWhiteSpace($targetName)) {
                    $resultProperties["DnsStatus"] = "NoTarget"
                    $resultProperties["DnsResolvedIPs"] = $null
                    $resultProperties["DnsMatchesAdIPv4"] = $null
                }
                else {
                    try {
                        $dnsResults = @(Resolve-DnsName -Name $targetName -Type A -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress -Unique)
                        $resultProperties["DnsStatus"] = if ($dnsResults.Count -gt 0) { "Resolved" } else { "NoARecord" }
                        $resultProperties["DnsResolvedIPs"] = ($dnsResults -join ",")

                        $adIps = @([string]$record.IPv4Address -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        if ($dnsResults.Count -gt 0 -and $adIps.Count -gt 0) {
                            $resultProperties["DnsMatchesAdIPv4"] = @($dnsResults | Where-Object { $adIps -contains $_ }).Count -gt 0
                        }
                        else {
                            $resultProperties["DnsMatchesAdIPv4"] = $null
                        }
                    }
                    catch {
                        $resultProperties["DnsStatus"] = "Failed"
                        $resultProperties["DnsResolvedIPs"] = $null
                        $resultProperties["DnsMatchesAdIPv4"] = $null
                    }
                }
            }

            if (@($using:TestPorts).Count -gt 0) {
                if ([string]::IsNullOrWhiteSpace($targetName)) {
                    $resultProperties["PortStatus"] = "NoTarget"
                }
                else {
                    $portStatuses = New-Object System.Collections.Generic.List[string]
                    foreach ($port in $using:TestPorts) {
                        $isOpen = Test-TcpPort -ComputerName $targetName -Port $port -TimeoutSeconds $using:TimeoutSeconds
                        $portStatuses.Add(("{0}:{1}" -f $port, $(if ($isOpen) { "Open" } else { "Closed" })))
                    }

                    $resultProperties["PortStatus"] = [string]::Join(";", $portStatuses)
                }
            }

            if ($using:RemoteInventory) {
                if ([string]::IsNullOrWhiteSpace($targetName)) {
                    $resultProperties["RemoteInventoryStatus"] = "NoTarget"
                }
                elseif ([string]::IsNullOrWhiteSpace([string]$using:DomainName)) {
                    $resultProperties["RemoteInventoryStatus"] = "SkippedUntrustedTarget: domain name unavailable"
                }
                elseif (-not $targetName.EndsWith(".$($using:DomainName)", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $resultProperties["RemoteInventoryStatus"] = "SkippedUntrustedTarget: target is outside the discovered AD DNS suffix"
                }
                elseif ($resultProperties.Contains("ConnectivityReachable") -and $resultProperties["ConnectivityReachable"] -eq $false) {
                    $resultProperties["RemoteInventoryStatus"] = "SkippedUnreachable"
                }
                else {
                    $cimSession = $null

                    try {
                        $cimSession = New-CimSession -ComputerName $targetName -Credential $using:Credential -OperationTimeoutSec $using:TimeoutSeconds -ErrorAction Stop

                        $operatingSystem = Get-CimInstance -ClassName "Win32_OperatingSystem" -CimSession $cimSession -OperationTimeoutSec $using:TimeoutSeconds -ErrorAction Stop
                        $computerSystem = Get-CimInstance -ClassName "Win32_ComputerSystem" -CimSession $cimSession -OperationTimeoutSec $using:TimeoutSeconds -ErrorAction Stop
                        $bios = Get-CimInstance -ClassName "Win32_BIOS" -CimSession $cimSession -OperationTimeoutSec $using:TimeoutSeconds -ErrorAction Stop

                        $systemDrive = if (-not [string]::IsNullOrWhiteSpace([string]$operatingSystem.SystemDrive)) {
                            [string]$operatingSystem.SystemDrive
                        }
                        else {
                            "C:"
                        }

                        $logicalDisk = Get-CimInstance -ClassName "Win32_LogicalDisk" -Filter "DeviceID = '$systemDrive'" -CimSession $cimSession -OperationTimeoutSec $using:TimeoutSeconds -ErrorAction SilentlyContinue
                        $pendingRebootRaw = Get-RemotePendingReboot -CimSession $cimSession -TimeoutSeconds $using:TimeoutSeconds

                        $resultProperties["RemoteInventoryStatus"] = "Success"
                        $resultProperties["RemoteUptimeDays"] = [Math]::Round(((Get-Date) - $operatingSystem.LastBootUpTime).TotalDays, 2)
                        $resultProperties["SerialNumber"] = [string]$bios.SerialNumber
                        $resultProperties["Model"] = [string]$computerSystem.Model
                        $resultProperties["TotalMemoryGB"] = if ($computerSystem.TotalPhysicalMemory) { [Math]::Round(([double]$computerSystem.TotalPhysicalMemory / 1GB), 2) } else { $null }
                        $resultProperties["SystemDriveFreeGB"] = if ($logicalDisk -and $logicalDisk.FreeSpace) { [Math]::Round(([double]$logicalDisk.FreeSpace / 1GB), 2) } else { $null }
                        $resultProperties["PendingReboot"] = if ($null -eq $pendingRebootRaw) { "Unknown" } else { $pendingRebootRaw }
                        $resultProperties["LoggedOnUser"] = [string]$computerSystem.UserName
                    }
                    catch {
                        $resultProperties["RemoteInventoryStatus"] = "Failed: $($_.Exception.Message)"
                    }
                    finally {
                        if ($null -ne $cimSession) {
                            $cimSession | Remove-CimSession
                        }
                    }
                }
            }

            [PSCustomObject]@{
                Index  = $wrapper.Index
                Record = [PSCustomObject]$resultProperties
            }
            } -ThrottleLimit $ThrottleLimit | ForEach-Object {
                $completedRecords++
                Write-ScanProgress -Id 5 -Activity $progressActivity `
                    -Status ("Completed {0} of {1} device(s)" -f $completedRecords, $Records.Count) `
                    -CompletedCount $completedRecords -TotalCount $Records.Count
                $_
            })
    }
    finally {
        Complete-ScanProgress -Id 5 -Activity $progressActivity
    }

    Write-Log -Message ("Operational enrichment complete: {0} device record(s) processed." -f $results.Count) -Level Info
    return @($results | Sort-Object -Property Index | ForEach-Object { $_.Record })
}

function Invoke-AdComputerQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Filters,

        [Parameter(Mandatory = $true)]
        [hashtable]$BaseQueryParameters,

        [Parameter(Mandatory = $false)]
        [string[]]$QuerySearchBases
    )

    $allResultsById = @{}
    $allResults = New-Object System.Collections.Generic.List[object]
    $queryWorkItems = New-Object System.Collections.Generic.List[object]
    $effectiveSearchBases = @(
        foreach ($querySearchBase in $QuerySearchBases) {
            if (-not [string]::IsNullOrWhiteSpace([string]$querySearchBase)) {
                $querySearchBase
            }
        }
    )

    foreach ($filter in @($Filters)) {
        if ($effectiveSearchBases.Count -eq 0) {
            $queryWorkItems.Add([PSCustomObject]@{
                    Filter     = $filter
                    SearchBase = $null
                })
        }
        else {
            foreach ($querySearchBase in $effectiveSearchBases) {
                $queryWorkItems.Add([PSCustomObject]@{
                        Filter     = $filter
                        SearchBase = $querySearchBase
                    })
            }
        }
    }

    $queryIndex = 0
    $progressActivity = "Discovering Active Directory devices"

    try {
        foreach ($queryWorkItem in $queryWorkItems) {
            $queryParameters = @{}
            foreach ($key in $BaseQueryParameters.Keys) {
                $queryParameters[$key] = $BaseQueryParameters[$key]
            }

            if ([string]::IsNullOrWhiteSpace([string]$queryWorkItem.SearchBase)) {
                Write-Log -Message "Running AD query against full scope." -Level Verbose
            }
            else {
                $queryParameters["SearchBase"] = $queryWorkItem.SearchBase
                Write-Log -Message "Running AD query against search base '$($queryWorkItem.SearchBase)'." -Level Verbose
            }

            Write-ScanProgress -Id 1 -Activity $progressActivity `
                -Status ("Query {0} of {1}; {2} unique device(s) discovered" -f ($queryIndex + 1), $queryWorkItems.Count, $allResults.Count) `
                -CompletedCount $queryIndex -TotalCount $queryWorkItems.Count

            Get-WithRetry -ScriptBlock {
                Get-ADComputer -Filter $queryWorkItem.Filter @queryParameters
            } | ForEach-Object {
                $computer = $_
                $identityKey = Get-ComputerIdentityKey -InputObject $computer
                if (-not [string]::IsNullOrWhiteSpace($identityKey) -and -not $allResultsById.ContainsKey($identityKey)) {
                    $allResultsById[$identityKey] = $true
                    $allResults.Add($computer)

                    Write-ScanProgress -Id 1 -Activity $progressActivity `
                        -Status ("Query {0} of {1}; {2} unique device(s) discovered" -f ($queryIndex + 1), $queryWorkItems.Count, $allResults.Count) `
                        -CompletedCount $queryIndex -TotalCount $queryWorkItems.Count
                }
            }

            $queryIndex++
            Write-ScanProgress -Id 1 -Activity $progressActivity `
                -Status ("Completed query {0} of {1}; {2} unique device(s) discovered" -f $queryIndex, $queryWorkItems.Count, $allResults.Count) `
                -CompletedCount $queryIndex -TotalCount $queryWorkItems.Count
        }
    }
    finally {
        Complete-ScanProgress -Id 1 -Activity $progressActivity
    }

    Write-Log -Message ("AD discovery complete: {0} unique device(s) discovered across {1} query operation(s)." -f $allResults.Count, $queryWorkItems.Count) -Level Info

    return @($allResults.ToArray())
}

if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Resolve-ExistingPath -Path $ConfigPath -BaseDirectory $ScriptDirectory

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        Write-ErrorAndExit -Message "Configuration file not found: $ConfigPath" -CodeKey Config
    }

    Assert-InputPathAllowed -Path $ConfigPath -Purpose "configuration file"

    try {
        $configObject = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        Write-ErrorAndExit -Message "Failed to read configuration file '$ConfigPath': $($_.Exception.Message)" -CodeKey Config
    }

    foreach ($configParameterName in @(
            "ComputerType",
            "Mode",
            "DomainController",
            "DomainName",
            "ComputerListPath",
            "OutputDirectory",
            "SearchBase",
            "SearchBaseList",
            "ExcludeOU",
            "InactiveDays",
            "IncludeDisabled",
            "ExportFormat",
            "CompareWithPrevious",
            "SummaryOnly",
            "SeparateStatusExports",
            "ResolveDns",
            "TestPorts",
            "RemoteInventory",
            "TimeoutSeconds",
            "ThrottleLimit",
            "PingCount",
            "TestMethod",
            "SkipPing",
            "LogPath",
            "NoProgress",
            "NoClobber",
            "ForceOverwrite",
            "AllowNetworkOutputPath",
            "AllowNetworkInputPath",
            "DisableCsvSanitization"
        )) {
        Set-ParameterFromConfig -ConfigObject $configObject -ParameterName $configParameterName
    }
}

if ($ForceOverwrite -and $NoClobber) {
    Write-ErrorAndExit -Message "ForceOverwrite and NoClobber cannot be used together." -CodeKey Validation
}

Assert-AllowedValues -Purpose "ComputerType" -Values @($ComputerType) -AllowedValues @("Server", "Workstation")
Assert-AllowedValues -Purpose "Mode" -Values @($Mode) -AllowedValues @("Full", "Targeted")
Assert-AllowedValues -Purpose "ExportFormat" -Values @($ExportFormat) -AllowedValues @("Csv", "Json", "Html")
Assert-AllowedValues -Purpose "TestMethod" -Values @($TestMethod) -AllowedValues @("Ping", "WinRM", "None")
Assert-IntegerRange -Purpose "TimeoutSeconds" -Value $TimeoutSeconds -Minimum 1 -Maximum 300
Assert-IntegerRange -Purpose "ThrottleLimit" -Value $ThrottleLimit -Minimum 1 -Maximum 128
Assert-IntegerRange -Purpose "PingCount" -Value $PingCount -Minimum 1 -Maximum 10
Assert-SafeTextValues -Purpose "SearchBase" -Values @($SearchBase)
Assert-SafeTextValues -Purpose "SearchBaseList" -Values $SearchBaseList
Assert-SafeTextValues -Purpose "ExcludeOU" -Values $ExcludeOU

if ($SkipPing.IsPresent) {
    if ($PSBoundParameters.ContainsKey("TestMethod") -and $TestMethod -ne "None") {
        Write-Warning "SkipPing overrides TestMethod to None."
    }

    $TestMethod = "None"
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = $ScriptDirectory
}

$OutputDirectory = Resolve-OutputPath -Path $OutputDirectory -BaseDirectory $ScriptDirectory

if ((Test-IsUncPath -Path $OutputDirectory) -and -not $AllowNetworkOutputPath) {
    Write-ErrorAndExit -Message "Network output directories are not allowed by default: $OutputDirectory. Re-run with -AllowNetworkOutputPath only if the location is trusted." -CodeKey Validation
}

if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $ComputerListPath = Resolve-ExistingPath -Path $ComputerListPath -BaseDirectory $ScriptDirectory
    Assert-InputPathAllowed -Path $ComputerListPath -Purpose "computer list"
}

if (-not [string]::IsNullOrWhiteSpace($CompareWithPrevious)) {
    $CompareWithPrevious = Resolve-ExistingPath -Path $CompareWithPrevious -BaseDirectory $ScriptDirectory
    Assert-InputPathAllowed -Path $CompareWithPrevious -Purpose "comparison file"
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $OutputDirectory "Scan-ADComputers_$RunTimestamp.log"
}
else {
    $LogPath = Resolve-OutputPath -Path $LogPath -BaseDirectory $ScriptDirectory
}

try {
    if (-not $WhatIfPreference -and -not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
    }

    Initialize-LogFile -Path $LogPath
}
catch {
    Write-ErrorAndExit -Message "Failed to prepare output or log location: $($_.Exception.Message)" -CodeKey Export
}

Write-Log -Message "Run started at $RunStartedAt" -Level Info
Write-Log -Message "Script directory: $ScriptDirectory" -Level Verbose

if ($Mode -eq "Targeted") {
    if ([string]::IsNullOrWhiteSpace($ComputerListPath)) {
        Write-ErrorAndExit -Message "Targeted mode requires -ComputerListPath." -CodeKey Validation
    }

    if (-not (Test-Path -LiteralPath $ComputerListPath)) {
        Write-ErrorAndExit -Message "Computer list file not found: $ComputerListPath" -CodeKey Validation
    }
}
elseif (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
    Write-ErrorAndExit -Message "-ComputerListPath can only be used with -Mode Targeted." -CodeKey Validation
}

if (-not [string]::IsNullOrWhiteSpace($CompareWithPrevious) -and -not (Test-Path -LiteralPath $CompareWithPrevious)) {
    Write-ErrorAndExit -Message "CompareWithPrevious file not found: $CompareWithPrevious" -CodeKey Validation
}

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-ErrorAndExit -Message "PowerShell 7 or greater is required." -CodeKey Prereq
}

Import-TrustedActiveDirectoryModule

if (-not $Credential) {
    Write-Log -Message "No credential provided; prompting for credentials..." -Level Info
    $Credential = Get-Credential -Message "Enter credentials for AD query and optional remote inventory"
}

$domainDiscoveryParameters = @{
    Credential = $Credential
}

if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
    $domainDiscoveryParameters["Server"] = $DomainController
}

try {
    $adDomain = Get-ADDomain @domainDiscoveryParameters -ErrorAction Stop
}
catch {
    Write-ErrorAndExit -Message "Failed to discover AD domain details: $($_.Exception.Message)" -CodeKey ADQuery
}

if ([string]::IsNullOrWhiteSpace($DomainName)) {
    $DomainName = $adDomain.DNSRoot
}

if ([string]::IsNullOrWhiteSpace($DomainController)) {
    $DomainController = $adDomain.PDCEmulator
}

Assert-SafeDnsName -Name $DomainName -Purpose "DomainName"
Assert-SafeDnsName -Name $DomainController -Purpose "DomainController"

$querySearchBases = @(Get-QuerySearchBases)
$queryScopeDescription = if ($querySearchBases.Count -gt 0) {
    [string]::Join("; ", $querySearchBases)
}
else {
    "entire domain"
}

$typeFilter = Get-ComputerTypeFilter -Type $ComputerType
$baseFilter = if ($IncludeDisabled.IsPresent) {
    $typeFilter
}
else {
    "Enabled -eq `$true -and $typeFilter"
}

$propertiesToGet = @(
    "canonicalName",
    "cn",
    "description",
    "distinguishedName",
    "dNSHostName",
    "enabled",
    "ipv4Address",
    "lastLogonDate",
    "lastLogonTimestamp",
    "name",
    "objectGUID",
    "operatingSystem",
    "operatingSystemVersion",
    "whenCreated"
)

$adQueryParameters = @{
    Credential  = $Credential
    Properties  = $propertiesToGet
    ErrorAction = "Stop"
}

if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
    $adQueryParameters["Server"] = $DomainController
}

Write-Log -Message "Mode: $Mode" -Level Info
Write-Log -Message "Computer type: $ComputerType" -Level Info
Write-Log -Message "Query scope: $queryScopeDescription" -Level Info
Write-Log -Message "Export formats: $([string]::Join(', ', @($ExportFormat)))" -Level Info

$exportPrefix = Get-ComputerExportPrefix -Type $ComputerType
$domainClean = Get-SafeFileNamePart -Value $DomainName
$inventoryBasePath = Join-Path $OutputDirectory "${exportPrefix}_${domainClean}_$RunTimestamp"
$auditBasePath = Join-Path $OutputDirectory "${exportPrefix}_${domainClean}_${RunTimestamp}_TargetedAudit"
$deltaBasePath = Join-Path $OutputDirectory "${exportPrefix}_${domainClean}_${RunTimestamp}_Delta"
$summaryBasePath = Join-Path $OutputDirectory "${exportPrefix}_${domainClean}_${RunTimestamp}_Summary"
$summaryBreakdownBasePath = Join-Path $OutputDirectory "${exportPrefix}_${domainClean}_${RunTimestamp}_SummaryBreakdown"

$requestedCount = 0
$adReturnedCount = 0
$excludedByOuCount = 0
$reachableCount = $null
$foundInAdCount = 0
$matchedCount = 0
$skippedCount = 0
$deltaCount = 0

$allQueriedComputers = @()
$computers = @()
$auditRecords = @()

try {
    if ($Mode -eq "Full") {
        Write-Log -Message "[MODE] Full AD inventory scan." -Level Info
        $allQueriedComputers = Invoke-AdComputerQuery -Filters @($baseFilter) -BaseQueryParameters $adQueryParameters -QuerySearchBases $querySearchBases
        $adReturnedCount = $allQueriedComputers.Count
        $computers = @($allQueriedComputers | Where-Object { -not (Test-IsExcludedByOu -DistinguishedName $_.distinguishedName -ExcludedOuList $ExcludeOU) })
        $excludedByOuCount = $adReturnedCount - $computers.Count
    }
    else {
        Write-Log -Message "[MODE] Targeted scan using list file: $ComputerListPath" -Level Info
        $requestedComputers = Get-RequestedComputerList -ListPath $ComputerListPath -DefaultDomainName $DomainName
        $requestedCount = $requestedComputers.Count

        if ($requestedCount -eq 0) {
            Write-ErrorAndExit -Message "The computer list file is empty: $ComputerListPath" -CodeKey Validation
        }

        Write-Log -Message "Requested computer names: $requestedCount" -Level Info
        $targetedFilters = Get-RequestedComputerFilters -RequestedComputers $requestedComputers -BaseFilter $baseFilter
        $allQueriedComputers = Invoke-AdComputerQuery -Filters $targetedFilters -BaseQueryParameters $adQueryParameters -QuerySearchBases $querySearchBases
        $adReturnedCount = $allQueriedComputers.Count

        $computerLookup = @{}
        foreach ($computer in $allQueriedComputers) {
            Add-ComputerLookupEntry -Lookup $computerLookup -Key $computer.name -Computer $computer
            Add-ComputerLookupEntry -Lookup $computerLookup -Key $computer.dNSHostName -Computer $computer
        }

        $connectivityLookup = @{}
        if ($TestMethod -ne "None") {
            Write-Log -Message "Running targeted connectivity validation with method '$TestMethod'." -Level Info
            $connectivityTargets = @(
                for ($index = 0; $index -lt $requestedComputers.Count; $index++) {
                    [PSCustomObject]@{
                        LookupKey  = [string]$index
                        TargetName = $requestedComputers[$index].FQDN
                    }
                }
            )

            $connectivityLookup = Get-ConnectivityLookup -TargetItems $connectivityTargets -Method $TestMethod -TimeoutSecondsValue $TimeoutSeconds -PingCountValue $PingCount -ThrottleLimitValue $ThrottleLimit
            $reachableCount = @($connectivityLookup.Values | Where-Object { $_.Reachable -eq $true }).Count
            Write-Log -Message ("Targeted connectivity complete: {0} of {1} target(s) reachable." -f $reachableCount, $requestedCount) -Level Info
        }
        else {
            $reachableCount = $null
        }

        $matchedById = @{}
        $matchedComputers = New-Object System.Collections.Generic.List[object]
        $auditList = New-Object System.Collections.Generic.List[object]
        $matchingActivity = "Matching targeted devices"

        try {
            for ($index = 0; $index -lt $requestedComputers.Count; $index++) {
                $requestedComputer = $requestedComputers[$index]
                $lookupKey = [string]$index
                $connectivityResult = if ($connectivityLookup.ContainsKey($lookupKey)) { $connectivityLookup[$lookupKey] } else { $null }

            $connectivityReachable = if ($TestMethod -eq "None") { $true } elseif ($connectivityResult) { [bool]$connectivityResult.Reachable } else { $false }
            $connectivityStatus = if ($TestMethod -eq "None") { "Skipped" } elseif ($connectivityResult) { $connectivityResult.Status } else { "Unreachable" }
            $connectivityDetail = if ($TestMethod -eq "None") { "Connectivity check skipped." } elseif ($connectivityResult) { $connectivityResult.Detail } else { "No connectivity result." }

            $matchedComputer = $null
            $matchedOn = $null

            if ($computerLookup.ContainsKey($requestedComputer.FQDN.ToUpperInvariant())) {
                $matchedComputer = $computerLookup[$requestedComputer.FQDN.ToUpperInvariant()]
                $matchedOn = "DNSHostName"
            }
            elseif ($computerLookup.ContainsKey($requestedComputer.ShortName.ToUpperInvariant())) {
                $matchedComputer = $computerLookup[$requestedComputer.ShortName.ToUpperInvariant()]
                $matchedOn = "Name"
            }

            $foundInAd = $null -ne $matchedComputer
            $excludedByOu = if ($foundInAd) { Test-IsExcludedByOu -DistinguishedName $matchedComputer.distinguishedName -ExcludedOuList $ExcludeOU } else { $false }
            $shouldExport = $connectivityReachable -and $foundInAd -and (-not $excludedByOu)

            $skipReasons = New-Object System.Collections.Generic.List[string]
            if (-not $connectivityReachable) {
                $skipReasons.Add("Unreachable")
            }
            if (-not $foundInAd) {
                $skipReasons.Add("NotFoundInAD")
            }
            if ($excludedByOu) {
                $skipReasons.Add("ExcludedOU")
            }

            if ($shouldExport) {
                $identityKey = Get-ComputerIdentityKey -InputObject $matchedComputer
                if (-not [string]::IsNullOrWhiteSpace($identityKey) -and -not $matchedById.ContainsKey($identityKey)) {
                    $matchedById[$identityKey] = $true
                    $matchedComputers.Add($matchedComputer)
                }
            }

                $auditList.Add([PSCustomObject]@{
                    IdentityKey            = if ($matchedComputer) { Get-ComputerIdentityKey -InputObject $matchedComputer } else { $null }
                    InputName              = $requestedComputer.InputName
                    ResolvedShortName      = $requestedComputer.ShortName
                    ResolvedFQDN           = $requestedComputer.FQDN
                    ConnectivityMethod     = $TestMethod
                    ConnectivityStatus     = $connectivityStatus
                    ConnectivityReachable  = if ($TestMethod -eq "None") { $null } else { $connectivityReachable }
                    ConnectivityDetail     = $connectivityDetail
                    FoundInAD              = $foundInAd
                    MatchedOn              = $matchedOn
                    ExcludedByOU           = $excludedByOu
                    Exported               = $shouldExport
                    SkipReason             = [string]::Join(";", $skipReasons)
                    ComputerType           = $ComputerType
                    CN                     = if ($matchedComputer) { $matchedComputer.cn } else { $null }
                    DNSHostName            = if ($matchedComputer) { $matchedComputer.dNSHostName } else { $null }
                    Description            = if ($matchedComputer) { $matchedComputer.description } else { $null }
                    CanonicalName          = if ($matchedComputer) { $matchedComputer.canonicalName } else { $null }
                    DistinguishedName      = if ($matchedComputer) { $matchedComputer.distinguishedName } else { $null }
                    Enabled                = if ($matchedComputer) { $matchedComputer.enabled } else { $null }
                    IPv4Address            = if ($matchedComputer) { (@($matchedComputer.ipv4Address) | Where-Object { $_ }) -join "," } else { $null }
                    LastLogonDate          = if ($matchedComputer) { $matchedComputer.lastLogonDate } else { $null }
                    LastLogonTimestamp     = if ($matchedComputer) { $matchedComputer.lastLogonTimestamp } else { $null }
                    ObjectGUID             = if ($matchedComputer) { $matchedComputer.objectGUID } else { $null }
                    OperatingSystem        = if ($matchedComputer) { $matchedComputer.operatingSystem } else { $null }
                    OperatingSystemVersion = if ($matchedComputer) { $matchedComputer.operatingSystemVersion } else { $null }
                    })

                Write-ScanProgress -Id 3 -Activity $matchingActivity `
                    -Status ("Processed {0} of {1} target(s)" -f ($index + 1), $requestedComputers.Count) `
                    -CompletedCount ($index + 1) -TotalCount $requestedComputers.Count
            }
        }
        finally {
            Complete-ScanProgress -Id 3 -Activity $matchingActivity
        }

        $auditRecords = @($auditList.ToArray())
        $computers = @($matchedComputers.ToArray())
        $excludedByOuCount = ($auditRecords | Where-Object { $_.ExcludedByOU }).Count
        $foundInAdCount = ($auditRecords | Where-Object { $_.FoundInAD -eq $true }).Count
        $matchedCount = $computers.Count
        $skippedCount = ($auditRecords | Where-Object { -not $_.Exported }).Count
        Write-Log -Message ("Target matching complete: {0} of {1} requested device(s) selected for export." -f $matchedCount, $requestedCount) -Level Info
    }
}
catch {
    Write-ErrorAndExit -Message "Failed AD query: $($_.Exception.Message)" -CodeKey ADQuery
}

if ($Mode -eq "Full") {
    $requestedCount = $adReturnedCount
    $foundInAdCount = $computers.Count
    $matchedCount = $computers.Count
    $skippedCount = 0
}

if ($Mode -eq "Targeted") {
    try {
        [void](Export-DataSet -BasePath $auditBasePath -Formats @($ExportFormat) -Data $auditRecords -Title "$ComputerType Targeted Audit")
    }
    catch {
        Write-ErrorAndExit -Message "Failed to export targeted audit data: $($_.Exception.Message)" -CodeKey Export
    }

    if ($SeparateStatusExports.IsPresent) {
        try {
            $matchedAudit = @($auditRecords | Where-Object { $_.Exported -eq $true })
            $unreachableAudit = @($auditRecords | Where-Object { $_.ConnectivityReachable -eq $false })
            $notFoundAudit = @($auditRecords | Where-Object { $_.FoundInAD -eq $false })

            [void](Export-DataSet -BasePath "${auditBasePath}_Matched" -Formats @($ExportFormat) -Data $matchedAudit -Title "$ComputerType Matched Targets")
            [void](Export-DataSet -BasePath "${auditBasePath}_Unreachable" -Formats @($ExportFormat) -Data $unreachableAudit -Title "$ComputerType Unreachable Targets")
            [void](Export-DataSet -BasePath "${auditBasePath}_NotFoundInAD" -Formats @($ExportFormat) -Data $notFoundAudit -Title "$ComputerType Not Found In AD")
        }
        catch {
            Write-ErrorAndExit -Message "Failed to export separate targeted status reports: $($_.Exception.Message)" -CodeKey Export
        }
    }
}

if ($computers.Count -eq 0) {
    Write-Log -Message "No $ComputerType objects found matching the current criteria." -Level Warning
}

$preparedRecordCount = 0
$preparationActivity = "Preparing inventory records"
try {
    $resultList = @(
        foreach ($computer in $computers) {
            Get-ComputerRecord -Computer $computer -Type $ComputerType -InactiveThresholdDays $InactiveDays
            $preparedRecordCount++
            Write-ScanProgress -Id 4 -Activity $preparationActivity `
                -Status ("Prepared {0} of {1} device record(s)" -f $preparedRecordCount, $computers.Count) `
                -CompletedCount $preparedRecordCount -TotalCount $computers.Count
        }
    )
}
finally {
    Complete-ScanProgress -Id 4 -Activity $preparationActivity
}
Write-Log -Message ("Record preparation complete: {0} device record(s) prepared." -f $resultList.Count) -Level Info

if ($Mode -eq "Targeted" -and $resultList.Count -gt 0) {
    $auditById = @{}
    foreach ($auditRecord in @($auditRecords | Where-Object { $_.Exported -eq $true -and -not [string]::IsNullOrWhiteSpace([string]$_.IdentityKey) })) {
        if (-not $auditById.ContainsKey($auditRecord.IdentityKey)) {
            $auditById[$auditRecord.IdentityKey] = $auditRecord
        }
    }

    foreach ($record in $resultList) {
        $identityKey = Get-ComputerIdentityKey -InputObject $record
        if (-not [string]::IsNullOrWhiteSpace($identityKey) -and $auditById.ContainsKey($identityKey)) {
            $auditRecord = $auditById[$identityKey]
            $record.ConnectivityMethod = $auditRecord.ConnectivityMethod
            $record.ConnectivityStatus = $auditRecord.ConnectivityStatus
            $record.ConnectivityReachable = $auditRecord.ConnectivityReachable
            $record.ConnectivityDetail = $auditRecord.ConnectivityDetail
        }
    }
}

$performConnectivityEnrichment = $false
if ($Mode -eq "Full" -and $TestMethod -ne "None") {
    $testMethodExplicit = (
        $PSBoundParameters.ContainsKey("TestMethod") -or
        $script:ConfigSourcedParameters.Contains("TestMethod") -or
        $RemoteInventory.IsPresent -or
        $script:ConfigSourcedParameters.Contains("RemoteInventory") -or
        @($TestPorts).Count -gt 0
    )
    if ($testMethodExplicit) {
        $performConnectivityEnrichment = $true
    }
}

try {
    $resultList = Invoke-OperationalEnrichment -Records $resultList -PerformConnectivityCheck $performConnectivityEnrichment
}
catch {
    Write-ErrorAndExit -Message "Operational enrichment failed: $($_.Exception.Message)" -CodeKey Operational
}

$summaryTotals = @(
    [PSCustomObject]@{
        RunStartedAt          = $RunStartedAt
        ComputerType          = $ComputerType
        Mode                  = $Mode
        QueryScope            = $queryScopeDescription
        RequestedCount        = $requestedCount
        AdReturnedCount       = $adReturnedCount
        ExcludedByOUCount     = $excludedByOuCount
        ReachableCount        = $reachableCount
        FoundInAdCount        = $foundInAdCount
        ExportedCount         = $resultList.Count
        IncludeDisabled       = [bool]$IncludeDisabled
        InactiveThresholdDays = if ($InactiveDays -gt 0) { $InactiveDays } else { $null }
        StaleCount            = @($resultList | Where-Object { $_.IsStale -eq $true }).Count
        NonStaleCount         = @($resultList | Where-Object { $_.StaleStatus -eq "Active" }).Count
        UnknownStaleCount     = @($resultList | Where-Object { $_.StaleStatus -eq "Unknown" }).Count
    }
)

$summaryBreakdown = Get-SummaryRecords -Records $resultList

if (-not [string]::IsNullOrWhiteSpace($CompareWithPrevious)) {
    try {
        $previousRecords = Import-PreviousDataSet -Path $CompareWithPrevious
        $deltaRecords = Get-DeltaRecords -PreviousRecords $previousRecords -CurrentRecords $resultList
        $deltaCount = $deltaRecords.Count
        [void](Export-DataSet -BasePath $deltaBasePath -Formats @($ExportFormat) -Data $deltaRecords -Title "$ComputerType Inventory Delta")
    }
    catch {
        Write-ErrorAndExit -Message "Failed to compare current inventory with '$CompareWithPrevious': $($_.Exception.Message)" -CodeKey Compare
    }
}

try {
    if ($SummaryOnly.IsPresent) {
        [void](Export-DataSet -BasePath $summaryBasePath -Formats @($ExportFormat) -Data $summaryTotals -Title "$ComputerType Inventory Summary Totals")
        [void](Export-DataSet -BasePath $summaryBreakdownBasePath -Formats @($ExportFormat) -Data $summaryBreakdown -Title "$ComputerType Inventory Summary Breakdown")
    }
    else {
        [void](Export-DataSet -BasePath $inventoryBasePath -Formats @($ExportFormat) -Data $resultList -Title "$ComputerType Inventory")
    }
}
catch {
    Write-ErrorAndExit -Message "Failed to export final report data: $($_.Exception.Message)" -CodeKey Export
}

Write-Log -Message "Summary:" -Level Info
Write-Log -Message ("  Requested        : {0}" -f $requestedCount) -Level Info
Write-Log -Message ("  AD Returned      : {0}" -f $adReturnedCount) -Level Info
Write-Log -Message ("  Excluded By OU   : {0}" -f $excludedByOuCount) -Level Info
if ($Mode -eq "Targeted") {
    Write-Log -Message ("  Reachable        : {0}" -f $(if ($null -eq $reachableCount) { "Skipped" } else { $reachableCount })) -Level Info
    Write-Log -Message ("  Found In AD      : {0}" -f $foundInAdCount) -Level Info
    Write-Log -Message ("  Skipped          : {0}" -f $skippedCount) -Level Info
}
else {
    Write-Log -Message "  Reachable        : Not evaluated in Full mode unless operational checks were requested." -Level Info
}
Write-Log -Message ("  Exported         : {0}" -f $resultList.Count) -Level Info
if ($InactiveDays -gt 0) {
    Write-Log -Message ("  Stale            : {0}" -f @($resultList | Where-Object { $_.IsStale -eq $true }).Count) -Level Info
}
if (-not [string]::IsNullOrWhiteSpace($CompareWithPrevious)) {
    Write-Log -Message ("  Delta Records    : {0}" -f $deltaCount) -Level Info
}

exit 0
