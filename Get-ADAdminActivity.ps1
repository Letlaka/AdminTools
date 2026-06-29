#requires -Version 5.1
using module ./AdminToolsCommon.psm1
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

    IMPORTANT: -AdminOnly uses current group membership at the time the script
    runs. It does not reconstruct historical membership. Events performed by
    accounts that were privileged at the time but have since been removed from
    privileged groups will be excluded. Events performed by accounts that were
    promoted after the event occurred will be included.

.PARAMETER AllowPartialResults
    Export results even if one or more Domain Controllers could not be queried.

.PARAMETER NoClobber
    Fail if the output CSV already exists instead of overwriting it.

.PARAMETER LogPath
    Optional path to the run log file.

.PARAMETER ForceOverwrite
    Overwrite an existing output CSV. Existing files are not overwritten by default.

.PARAMETER AllowNetworkOutputPath
    Allow writing the CSV report to a UNC path. Network output paths are rejected by default.

.PARAMETER AllowNetworkInputPath
    Allow reading credential files from UNC paths. Network input paths are rejected by default.

.PARAMETER AllowUnverifiedDomainController
    Allow manually supplied Domain Controller names without verifying them against AD discovery.

.PARAMETER DisableCsvSanitization
    Export raw string values without Excel formula injection protection.

.PARAMETER Credential
    Credential used for AD discovery, privileged group lookups, and remote Security log reads.

.PARAMETER CredentialSecretName
    SecretManagement secret name containing a PSCredential.

.PARAMETER CredentialPath
    Path to a PSCredential exported with Export-Clixml. Credential files must be outside the repository directory.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseSingularNouns",
    "",
    Justification = "Get-DomainControllerNames returns a collection and retains its established internal name."
)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSAvoidUsingPlainTextForPassword",
    "",
    Justification = "CredentialSecretName is a SecretManagement lookup key and CredentialPath is a file path; neither parameter carries a password value."
)]
[CmdletBinding()]
param(
    [ValidateRange(1, 3650)]
    [int]$DaysBack = 7,

    [string[]]$DomainControllers,

    [PSCredential]$Credential,

    [string]$CredentialSecretName,

    [string]$CredentialPath,

    [switch]$AdminOnly,

    [string[]]$AdminSamAccountNames,

    [string[]]$PrivilegedGroupNames = (Get-DefaultPrivilegedGroupNames),

    [string]$OutputCsv,

    [string]$LogPath,

    [switch]$IncludeMessage,

    [ValidateRange(1, 10000)]
    [int]$MaxAttributeValueLength = 500,

    [ValidateRange(0, 1000000)]
    [int]$MaxEventsPerDomainController = 0,

    [switch]$AllowPartialResults,

    [switch]$NoClobber,

    [switch]$ForceOverwrite,

    [switch]$AllowNetworkOutputPath,

    [switch]$AllowNetworkInputPath,

    [switch]$AllowUnverifiedDomainController,

    [switch]$DisableCsvSanitization
)

Import-Module (Join-Path $PSScriptRoot "AdminToolsCommon.psm1") -Force -ErrorAction Stop

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$ExitCodes = @{
    Success    = 0
    General    = 1
    Prereq     = 2
    Validation = 4
    ADQuery    = 5
    Export     = 7
}

$ScriptDirectory = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$RunStartedAt = Get-Date
$RunTimestamp = $RunStartedAt.ToString("yyyyMMddHHmmss")
$SecurityAuditProviderName = "Microsoft-Windows-Security-Auditing"
$StartTime = $RunStartedAt.AddDays(-1 * $DaysBack)

if (-not $PSBoundParameters.ContainsKey("OutputCsv")) {
    $DefaultOutputRoot = Join-Path -Path $PSScriptRoot -ChildPath "reports\ad-admin-activity"
    $OutputCsv = Join-Path -Path $DefaultOutputRoot -ChildPath ("AD_Admin_Activity_Report_{0}.csv" -f $RunTimestamp)
}

Assert-SafeTextValues -Purpose "CredentialSecretName" -Values @($CredentialSecretName) -MaximumLength 256
Assert-SafeTextValues -Purpose "CredentialPath" -Values @($CredentialPath)
Assert-SafeTextValues -Purpose "LogPath" -Values @($LogPath)
$Credential = Resolve-AdminToolsCredential -Credential $Credential -CredentialSecretName $CredentialSecretName -CredentialPath $CredentialPath -BaseDirectory $PSScriptRoot -AllowNetworkInputPath:$AllowNetworkInputPath

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

function Get-CurrentPrivilegedAdminName {
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

function Write-RunLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )

    if (-not $script:FileLoggingEnabled -or [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
        return
    }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[{0}] [{1}] {2}" -f $timestamp, $Level.ToUpperInvariant(), $Message
    Add-Content -LiteralPath $script:LogFilePath -Value $entry -Encoding UTF8
}

function Initialize-RunLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ((Test-IsUncPath -Path $Path) -and -not $AllowNetworkOutputPath) {
        throw "Network log paths are not allowed by default: $Path. Re-run with -AllowNetworkOutputPath only if the location is trusted."
    }

    if ((Test-Path -LiteralPath $Path) -and ($NoClobber -or -not $ForceOverwrite)) {
        throw "Log file already exists: $Path. Re-run with -ForceOverwrite only if replacing it is intended."
    }

    $logDirectory = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    Set-Content -LiteralPath $Path -Value @(
        "Run log: Get-ADAdminActivity.ps1",
        "Started: $($RunStartedAt.ToString('o'))"
    ) -Encoding UTF8

    $script:LogFilePath = $Path
    $script:FileLoggingEnabled = $true
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

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path (Join-Path -Path $ScriptDirectory -ChildPath "logs\get-ad-admin-activity") -ChildPath ("Get-ADAdminActivity_{0}.log" -f $RunTimestamp)
}
else {
    $LogPath = Get-ResolvedOutputPath -Path $LogPath
}

Initialize-RunLog -Path $LogPath
Write-Information "Run log: $LogPath"
Write-RunLog -Message "Starting Get-ADAdminActivity. DaysBack=$DaysBack; AdminOnly=$($AdminOnly.IsPresent); OutputCsv=$OutputCsv"
Assert-SafeDomainControllerNames -Names $DomainControllers
Assert-SafeTextValues -Purpose "AdminSamAccountNames" -Values $AdminSamAccountNames -MaximumLength 256
Assert-SafeTextValues -Purpose "PrivilegedGroupNames" -Values $PrivilegedGroupNames -MaximumLength 256

$ResolvedDomainControllers = Get-DomainControllerNames `
    -ProvidedDomainControllers $DomainControllers `
    -AllowUnverified:$AllowUnverifiedDomainController `
    -Credential $Credential

if (@($ResolvedDomainControllers).Count -eq 0) {
    throw "No writable Domain Controllers were found. Verify AD connectivity and that the account has permission to enumerate domain controllers."
}

$CurrentAdminNames = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

# Resolves current group membership. See NOTES in .DESCRIPTION: this is a
# point-in-time snapshot, not a historical reconstruction.
if ($AdminOnly) {
    $CurrentAdminNames = Get-CurrentPrivilegedAdminName `
        -GroupNames $PrivilegedGroupNames `
        -AdditionalAdminNames $AdminSamAccountNames `
        -Credential $Credential

    if ($CurrentAdminNames.Count -eq 0) {
        throw "AdminOnly was selected, but no admin accounts could be resolved. Provide -AdminSamAccountNames or verify AD module access."
    }
}

if ($AdminOnly -and $MaxEventsPerDomainController -gt 0) {
    Write-Warning "AdminOnly filtering is applied after MaxEventsPerDomainController cap. Set MaxEventsPerDomainController to 0 to ensure all relevant events are considered."
}

$SuccessfulDomainControllers = New-Object System.Collections.Generic.List[string]
$FailedDomainControllers = New-Object System.Collections.Generic.List[object]

$AllRecords = foreach ($DomainController in $ResolvedDomainControllers) {
    Write-Verbose "Reading Security log from $DomainController from $StartTime"

    try {
        # NOTE: MaxEventsPerDomainController caps the raw event pull before
        # per-user filtering is applied. If this limit is set very low and
        # the target user's events are not among the first N events returned,
        # they will be silently omitted. Set MaxEventsPerDomainController to
        # 0 (no limit) when querying for a specific user or when using
        # AdminOnly to ensure complete results.
        $SecurityEvents = Get-SecurityAuditEvents `
            -ComputerName $DomainController `
            -Ids $EventIds `
            -ProviderName $SecurityAuditProviderName `
            -StartDate $StartTime `
            -MaxEvents $MaxEventsPerDomainController `
            -Credential $Credential

        $RawRecords = @($SecurityEvents | ForEach-Object {
            ConvertTo-AdActivityRecord `
                -EventRecord $_ `
                -DomainController $DomainController `
                -IncludeRenderedMessage:$IncludeMessage `
                -AttributeValueLimit $MaxAttributeValueLength
        })

        $DomainControllerRecords = if ($AdminOnly) {
            @($RawRecords | Where-Object {
                $Record = $_
                $Matched = $false
                foreach ($ActorIdentifier in @(
                        $Record.ActorSamAccountName,
                        $Record.ActorAccount,
                        $Record.ActorSid
                    )) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$ActorIdentifier) -and
                        $CurrentAdminNames.Contains([string]$ActorIdentifier)) {
                        $Matched = $true
                        break
                    }
                }
                $Matched
            })
        } else {
            $RawRecords
        }

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

    try {
        $CsvRecords | Export-Csv -LiteralPath $ResolvedOutputCsv -NoTypeInformation -Encoding UTF8
        Write-Information "Report exported to: $ResolvedOutputCsv"
        Write-RunLog -Message "Report exported to: $ResolvedOutputCsv"
    }
    catch {
        Write-Error "Failed to export report to '$ResolvedOutputCsv': $($_.Exception.Message)"
        exit $ExitCodes.Export
    }
}

$SummaryMessage = "Summary: queried {0} Domain Controller(s), succeeded {1}, failed {2}, exported {3} event record(s)." -f `
        @($ResolvedDomainControllers).Count,
        $SuccessfulDomainControllers.Count,
        $FailedDomainControllers.Count,
        @($SortedRecords).Count
Write-Information $SummaryMessage
Write-RunLog -Message $SummaryMessage

if ($FailedDomainControllers.Count -gt 0) {
    Write-Warning "Partial report: at least one Domain Controller could not be queried."
    Write-RunLog -Message "Partial report: at least one Domain Controller could not be queried." -Level Warning
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

Write-RunLog -Message "Completed Get-ADAdminActivity. ExitCode=$($ExitCodes.Success)"

exit $ExitCodes.Success




