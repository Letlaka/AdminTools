#requires -Version 5.1
<#
.SYNOPSIS
    Reports Active Directory administrative activity from Domain Controller Security logs.

.DESCRIPTION
    Reads AD-related Security Event IDs from Domain Controllers and extracts:
    - Who performed the action
    - What object/account/group was targeted
    - What changed
    - When and on which Domain Controller it was logged

.NOTES
    Run from an elevated PowerShell session using an account allowed to read
    Security logs on Domain Controllers.

    Optional: ActiveDirectory PowerShell module is used to discover DCs and
    current privileged group members.

.PARAMETER AllowPartialResults
    Export results even if one or more Domain Controllers could not be queried.

.PARAMETER NoClobber
    Fail if the output CSV already exists instead of overwriting it.

.PARAMETER ForceOverwrite
    Overwrite an existing output CSV. Existing files are not overwritten by default.

.PARAMETER AllowNetworkOutputPath
    Allow writing the CSV report to a UNC path. Network output paths are rejected by default.

.PARAMETER AllowUnverifiedDomainController
    Allow manually supplied Domain Controller names without verifying them against AD discovery.

.PARAMETER DisableCsvSanitization
    Export raw string values without Excel formula injection protection.

.PARAMETER Credential
    Credential used for AD discovery, privileged group lookups, and remote Security log reads.
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 3650)]
    [int]$DaysBack = 7,

    [string[]]$DomainControllers,

    [PSCredential]$Credential,

    [switch]$AdminOnly,

    [string[]]$AdminSamAccountNames,

    [string[]]$PrivilegedGroupNames = @(
        "Domain Admins",
        "Enterprise Admins",
        "Schema Admins",
        "Administrators",
        "Account Operators",
        "Server Operators",
        "Backup Operators",
        "Group Policy Creator Owners"
    ),

    [string]$OutputCsv,

    [switch]$IncludeMessage,

    [ValidateRange(1, 10000)]
    [int]$MaxAttributeValueLength = 500,

    [ValidateRange(0, 1000000)]
    [int]$MaxEventsPerDomainController = 0,

    [switch]$AllowPartialResults,

    [switch]$NoClobber,

    [switch]$ForceOverwrite,

    [switch]$AllowNetworkOutputPath,

    [switch]$AllowUnverifiedDomainController,

    [switch]$DisableCsvSanitization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$SecurityAuditProviderName = "Microsoft-Windows-Security-Auditing"
$StartTime = (Get-Date).AddDays(-1 * $DaysBack)

if (-not $PSBoundParameters.ContainsKey("OutputCsv")) {
    $DefaultOutputRoot = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        Join-Path -Path $env:LOCALAPPDATA -ChildPath "AdminTools\ADAdminActivityReports"
    }
    else {
        Join-Path -Path (Get-Location).Path -ChildPath "ADAdminActivityReports"
    }

    $OutputCsv = Join-Path -Path $DefaultOutputRoot -ChildPath ("AD_Admin_Activity_Report_{0}.csv" -f (Get-Date -Format "yyyyMMddHHmmss"))
}

if ($ForceOverwrite -and $NoClobber) {
    throw "ForceOverwrite and NoClobber cannot be used together."
}

# Core AD administrative event IDs.
# Extend this list if you want to include logon activity such as 4624/4625.
$EventIds = @(
    # User account management
    4720, # User created
    4722, # User enabled
    4723, # Password change attempted
    4724, # Password reset attempted
    4725, # User disabled
    4726, # User deleted
    4738, # User changed
    4740, # User locked out
    4767, # User unlocked
    4781, # Account name changed

    # Computer account management
    4741, # Computer created
    4742, # Computer changed
    4743, # Computer deleted

    # Security group management
    4727, # Global security group created
    4728, # Member added to global security group
    4729, # Member removed from global security group
    4730, # Global security group deleted
    4731, # Local security group created
    4732, # Member added to local security group
    4733, # Member removed from local security group
    4734, # Local security group deleted
    4735, # Local security group changed
    4737, # Global security group changed
    4754, # Universal security group created
    4755, # Universal security group changed
    4756, # Member added to universal security group
    4757, # Member removed from universal security group
    4758, # Universal security group deleted

    # Domain / audit policy
    4719, # System audit policy changed
    4739, # Domain policy changed

    # Directory Service object changes
    5136, # Directory object modified
    5137, # Directory object created
    5138, # Directory object undeleted
    5139, # Directory object moved
    5141  # Directory object deleted
)

$EventActionMap = @{
    4719 = "System audit policy changed"
    4720 = "User account created"
    4722 = "User account enabled"
    4723 = "Password change attempted"
    4724 = "Password reset attempted"
    4725 = "User account disabled"
    4726 = "User account deleted"
    4727 = "Global security group created"
    4728 = "Member added to global security group"
    4729 = "Member removed from global security group"
    4730 = "Global security group deleted"
    4731 = "Local security group created"
    4732 = "Member added to local security group"
    4733 = "Member removed from local security group"
    4734 = "Local security group deleted"
    4735 = "Local security group changed"
    4737 = "Global security group changed"
    4738 = "User account changed"
    4739 = "Domain policy changed"
    4740 = "User account locked out"
    4741 = "Computer account created"
    4742 = "Computer account changed"
    4743 = "Computer account deleted"
    4754 = "Universal security group created"
    4755 = "Universal security group changed"
    4756 = "Member added to universal security group"
    4757 = "Member removed from universal security group"
    4758 = "Universal security group deleted"
    4767 = "User account unlocked"
    4781 = "Account name changed"
    5136 = "Directory object modified"
    5137 = "Directory object created"
    5138 = "Directory object undeleted"
    5139 = "Directory object moved"
    5141 = "Directory object deleted"
}

function Test-IsTrustedActiveDirectoryModulePath {
    param(
        [string]$Path
    )

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
    $LoadedModules = @(Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)
    if ($LoadedModules.Count -gt 0) {
        foreach ($LoadedModule in $LoadedModules) {
            if (-not (Test-IsTrustedActiveDirectoryModulePath -Path $LoadedModule.ModuleBase)) {
                Write-Warning "ActiveDirectory module is already loaded from an untrusted path: $($LoadedModule.ModuleBase)"
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
        Write-Warning "Trusted ActiveDirectory module could not be found in the Windows RSAT module path."
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
        Write-Warning "ActiveDirectory module could not be loaded from trusted path '$($TrustedModule[0].ModuleBase)'. Error: $($_.Exception.Message)"
        return $false
    }
}

function Get-DomainControllerNames {
    param(
        [string[]]$ProvidedDomainControllers,
        [switch]$AllowUnverified,
        [PSCredential]$Credential
    )

    $AdCommandParameters = @{}
    if ($Credential) {
        $AdCommandParameters["Credential"] = $Credential
    }

    if ($ProvidedDomainControllers -and $ProvidedDomainControllers.Count -gt 0) {
        if ($AllowUnverified) {
            Write-Warning "Using manually supplied Domain Controller names without AD verification."
            return $ProvidedDomainControllers
        }

        if (-not (Import-ActiveDirectoryModule)) {
            throw "DomainControllers were provided but could not be verified because the trusted ActiveDirectory module is unavailable. Re-run with -AllowUnverifiedDomainController only if these names are trusted."
        }

        $DiscoveredDomainControllers = @(Get-ADDomainController -Filter * @AdCommandParameters |
                Where-Object { -not $_.IsReadOnly } |
                Select-Object -ExpandProperty HostName)

        $DiscoveredSet = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        foreach ($DiscoveredDomainController in $DiscoveredDomainControllers) {
            [void]$DiscoveredSet.Add($DiscoveredDomainController)
            [void]$DiscoveredSet.Add(($DiscoveredDomainController -split '\.')[0])
        }

        foreach ($ProvidedDomainController in $ProvidedDomainControllers) {
            if (-not $DiscoveredSet.Contains($ProvidedDomainController)) {
                throw "Provided Domain Controller '$ProvidedDomainController' was not found in AD discovery. Re-run with -AllowUnverifiedDomainController only if this host is trusted."
            }
        }

        return $ProvidedDomainControllers
    }

    if (-not (Import-ActiveDirectoryModule)) {
        throw "No Domain Controllers were provided and the ActiveDirectory module is unavailable."
    }

    return Get-ADDomainController -Filter * @AdCommandParameters |
        Where-Object { -not $_.IsReadOnly } |
        Select-Object -ExpandProperty HostName
}

function Get-CurrentPrivilegedAdminNames {
    param(
        [string[]]$GroupNames,
        [string[]]$AdditionalAdminNames,
        [PSCredential]$Credential
    )

    $AdCommandParameters = @{}
    if ($Credential) {
        $AdCommandParameters["Credential"] = $Credential
    }

    $AdminNameSet = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($AdminName in $AdditionalAdminNames) {
        if (-not [string]::IsNullOrWhiteSpace($AdminName)) {
            [void]$AdminNameSet.Add($AdminName)

            if ($AdminName -match '\\') {
                [void]$AdminNameSet.Add(($AdminName -split '\\')[-1])
            }
        }
    }

    if (-not (Import-ActiveDirectoryModule)) {
        return $AdminNameSet
    }

    foreach ($GroupName in $GroupNames) {
        try {
            $Group = Get-ADGroup -Identity $GroupName @AdCommandParameters -ErrorAction Stop

            Get-ADGroupMember -Identity $Group.DistinguishedName -Recursive @AdCommandParameters -ErrorAction Stop |
                Where-Object { $_.ObjectClass -eq "user" } |
                ForEach-Object {
                    if (-not [string]::IsNullOrWhiteSpace($_.SamAccountName)) {
                        [void]$AdminNameSet.Add($_.SamAccountName)
                    }

                    if ($_.PSObject.Properties["SID"] -and $null -ne $_.SID) {
                        [void]$AdminNameSet.Add([string]$_.SID)
                    }
                }
        }
        catch {
            Write-Warning "Could not read members of privileged group '$GroupName'. Error: $($_.Exception.Message)"
        }
    }

    return $AdminNameSet
}

function ConvertTo-EventDataMap {
    param(
        [System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord
    )

    [xml]$EventXml = $EventRecord.ToXml()
    $EventDataMap = @{}

    foreach ($DataElement in $EventXml.Event.EventData.Data) {
        if ($null -ne $DataElement.Name) {
            $EventDataMap[$DataElement.Name] = [string]$DataElement.'#text'
        }
    }

    return $EventDataMap
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

function Limit-TextLength {
    param(
        [string]$Text,
        [int]$MaximumLength
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    if ($Text.Length -le $MaximumLength) {
        return $Text
    }

    return $Text.Substring(0, $MaximumLength) + "...[truncated]"
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

function Test-IsCsvPath {
    param(
        [string]$Path
    )

    return ([System.IO.Path]::GetExtension($Path) -ieq ".csv")
}

function Write-SensitiveOutputWarning {
    param(
        [string]$Path,
        [switch]$IncludesRenderedMessage
    )

    Write-Warning "The report contains sensitive audit data. Store it in a restricted location and limit access to authorized administrators."

    if ($IncludesRenderedMessage) {
        Write-Warning "IncludeMessage is enabled. Rendered event messages can contain additional account, object, and attribute details."
    }

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        Write-Verbose "Sensitive report output path: $Path"
    }
}

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

function ConvertTo-AdActivityRecord {
    param(
        [System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord,
        [string]$DomainController,
        [switch]$IncludeRenderedMessage,
        [int]$AttributeValueLimit
    )

    $EventData = ConvertTo-EventDataMap -EventRecord $EventRecord

    $SubjectUserName = Get-MapValue -Map $EventData -Key "SubjectUserName"
    $SubjectDomainName = Get-MapValue -Map $EventData -Key "SubjectDomainName"
    $SubjectLogonId = Get-MapValue -Map $EventData -Key "SubjectLogonId"
    $SubjectUserSid = Get-MapValue -Map $EventData -Key "SubjectUserSid"

    $ActorAccount = Join-DomainAccount `
        -DomainName $SubjectDomainName `
        -AccountName $SubjectUserName

    $TargetAccount = Join-DomainAccount `
        -DomainName (Get-MapValue -Map $EventData -Key "TargetDomainName") `
        -AccountName (Get-MapValue -Map $EventData -Key "TargetUserName")

    $EventId = [int]$EventRecord.Id
    $Action = $EventActionMap[$EventId]

    $ObjectDistinguishedName = Get-MapValue -Map $EventData -Key "ObjectDN"
    $MemberName = Get-MapValue -Map $EventData -Key "MemberName"
    $AttributeName = Get-MapValue -Map $EventData -Key "AttributeLDAPDisplayName"
    $OperationType = Get-MapValue -Map $EventData -Key "OperationType"
    $AttributeValue = Limit-TextLength `
        -Text (Get-MapValue -Map $EventData -Key "AttributeValue") `
        -MaximumLength $AttributeValueLimit

    $TargetObject = $null

    if (-not [string]::IsNullOrWhiteSpace($ObjectDistinguishedName)) {
        $TargetObject = $ObjectDistinguishedName
    }
    elseif (-not [string]::IsNullOrWhiteSpace($TargetAccount)) {
        $TargetObject = $TargetAccount
    }
    elseif (-not [string]::IsNullOrWhiteSpace($MemberName)) {
        $TargetObject = $MemberName
    }
    else {
        $TargetObject = Get-MapValue -Map $EventData -Key "TargetSid"
    }

    $RenderedMessage = $null
    if ($IncludeRenderedMessage) {
        $RenderedMessage = $EventRecord.Message
    }

    [pscustomobject]@{
        TimeCreated                 = $EventRecord.TimeCreated
        DomainController            = $DomainController
        EventId                     = $EventId
        Action                      = $Action
        ActorAccount                = $ActorAccount
        ActorSamAccountName         = $SubjectUserName
        ActorSid                    = $SubjectUserSid
        SubjectLogonId              = $SubjectLogonId
        TargetObject                = $TargetObject
        TargetAccount               = $TargetAccount
        MemberName                  = $MemberName
        ObjectDistinguishedName     = $ObjectDistinguishedName
        AttributeName               = $AttributeName
        OperationType               = $OperationType
        AttributeValue              = $AttributeValue
        EventRecordId               = $EventRecord.RecordId
        RenderedMessage             = $RenderedMessage
    }
}

$ResolvedDomainControllers = Get-DomainControllerNames `
    -ProvidedDomainControllers $DomainControllers `
    -AllowUnverified:$AllowUnverifiedDomainController `
    -Credential $Credential

$CurrentAdminNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

if ($AdminOnly) {
    $CurrentAdminNames = Get-CurrentPrivilegedAdminNames `
        -GroupNames $PrivilegedGroupNames `
        -AdditionalAdminNames $AdminSamAccountNames `
        -Credential $Credential

    if ($CurrentAdminNames.Count -eq 0) {
        throw "AdminOnly was selected, but no admin accounts could be resolved. Provide -AdminSamAccountNames or verify AD module access."
    }
}

$SuccessfulDomainControllers = New-Object System.Collections.Generic.List[string]
$FailedDomainControllers = New-Object System.Collections.Generic.List[object]

$AllRecords = foreach ($DomainController in $ResolvedDomainControllers) {
    Write-Verbose "Reading Security log from $DomainController from $StartTime"

    try {
        $SecurityEvents = Get-SecurityAuditEvents `
            -ComputerName $DomainController `
            -Ids $EventIds `
            -ProviderName $SecurityAuditProviderName `
            -StartDate $StartTime `
            -MaxEvents $MaxEventsPerDomainController `
            -Credential $Credential

        $DomainControllerRecords = @($SecurityEvents | ForEach-Object {
            $ActivityRecord = ConvertTo-AdActivityRecord `
                -EventRecord $_ `
                -DomainController $DomainController `
                -IncludeRenderedMessage:$IncludeMessage `
                -AttributeValueLimit $MaxAttributeValueLength

            if ($AdminOnly) {
                $ActorMatchedAdmin = $false
                foreach ($ActorIdentifier in @(
                        $ActivityRecord.ActorSamAccountName,
                        $ActivityRecord.ActorAccount,
                        $ActivityRecord.ActorSid
                    )) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$ActorIdentifier) -and $CurrentAdminNames.Contains([string]$ActorIdentifier)) {
                        $ActorMatchedAdmin = $true
                        break
                    }
                }

                if (-not $ActorMatchedAdmin) {
                    return
                }
            }

            $ActivityRecord
        })

        [void]$SuccessfulDomainControllers.Add($DomainController)
        $DomainControllerRecords
    }
    catch {
        if ($_.Exception.Message -like "*No events were found*") {
            [void]$SuccessfulDomainControllers.Add($DomainController)
            Write-Verbose "No matching Security log events found on $DomainController."
            continue
        }

        $Failure = [pscustomobject]@{
            DomainController = $DomainController
            Error            = $_.Exception.Message
        }

        [void]$FailedDomainControllers.Add($Failure)
        Write-Warning "Failed to read Security log from $DomainController. Error: $($_.Exception.Message)"
    }
}

if ($FailedDomainControllers.Count -gt 0) {
    $FailedNames = ($FailedDomainControllers | Select-Object -ExpandProperty DomainController) -join ", "
    $PartialMessage = "Failed to read Security logs from $($FailedDomainControllers.Count) Domain Controller(s): $FailedNames"

    if (-not $AllowPartialResults) {
        throw "$PartialMessage. No report was exported. Re-run with -AllowPartialResults if a partial report is acceptable."
    }

    Write-Warning "$PartialMessage. Exporting partial results because AllowPartialResults was specified."
}

$SortedRecords = $AllRecords | Sort-Object TimeCreated -Descending

if ($OutputCsv) {
    $ResolvedOutputCsv = Get-ResolvedOutputPath -Path $OutputCsv

    if (-not (Test-IsCsvPath -Path $ResolvedOutputCsv)) {
        throw "OutputCsv must use a .csv file extension: $ResolvedOutputCsv"
    }

    if ((Test-IsUncPath -Path $ResolvedOutputCsv) -and -not $AllowNetworkOutputPath) {
        throw "Network output paths are not allowed by default: $ResolvedOutputCsv. Re-run with -AllowNetworkOutputPath only if the location is trusted."
    }

    $OutputDirectory = Split-Path -Path $ResolvedOutputCsv -Parent

    if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    if (Test-Path -LiteralPath $ResolvedOutputCsv) {
        if ($NoClobber -or -not $ForceOverwrite) {
            throw "Output file already exists: $ResolvedOutputCsv. Re-run with -ForceOverwrite only if replacing it is intended."
        }

        Write-Warning "Overwriting existing report file: $ResolvedOutputCsv"
    }

    Write-SensitiveOutputWarning -Path $ResolvedOutputCsv -IncludesRenderedMessage:$IncludeMessage
    $CsvRecords = if ($DisableCsvSanitization) {
        Write-Warning "CSV sanitization is disabled. Opening this report in spreadsheet software may evaluate event data as formulas."
        $SortedRecords
    }
    else {
        @($SortedRecords | ConvertTo-SafeCsvRecord)
    }

    $CsvRecords | Export-Csv -LiteralPath $ResolvedOutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Report exported to: $ResolvedOutputCsv"
}

Write-Host ("Summary: queried {0} Domain Controller(s), succeeded {1}, failed {2}, exported {3} event record(s)." -f `
        @($ResolvedDomainControllers).Count,
        $SuccessfulDomainControllers.Count,
        $FailedDomainControllers.Count,
        @($SortedRecords).Count)

if ($FailedDomainControllers.Count -gt 0) {
    Write-Warning "Partial report: at least one Domain Controller could not be queried."
}

$SortedRecords |
    Select-Object -First 50 `
        TimeCreated,
        DomainController,
        EventId,
        Action,
        ActorAccount,
        TargetObject,
        AttributeName,
        OperationType |
    Format-Table -AutoSize
