#requires -Version 5.1
using module ./AdminToolsCommon.psm1
<#
.SYNOPSIS
    Manage and report on Active Directory user accounts.

.DESCRIPTION
    Provides safe read-only reporting for AD user accounts, focused user audit
    lookups from Domain Controller Security logs, locked-out account details,
    and explicit account reset actions such as unlock and password reset.

    Account-changing operations support -WhatIf and confirmation through
    SupportsShouldProcess.

.PARAMETER Mode
    Operation mode. Report generates account reports, UserAudit reports on one
    user and optional Security events, LockedOut lists currently locked accounts
    with optional lockout events, and Reset performs explicit reset actions.

.PARAMETER Identity
    User identity for UserAudit or Reset mode. Accepts values supported by
    Get-ADUser -Identity, such as SamAccountName, DistinguishedName, GUID, or SID.

.PARAMETER UserListPath
    Optional plain text file containing one user identity per line. Blank lines
    and lines beginning with # are ignored.

.PARAMETER ReportType
    One or more reports to generate in Report mode.

.PARAMETER Unlock
    Unlock the selected user in Reset mode.

.PARAMETER Enable
    Enable the selected user account in Reset mode.

.PARAMETER ResetPassword
    Reset the selected user's password in Reset mode.

.PARAMETER NewPassword
    SecureString password used with -ResetPassword.

.PARAMETER GenerateTemporaryPassword
    Generate a temporary password for -ResetPassword. The generated value is
    displayed only when -ShowGeneratedPassword is also supplied and is not
    written to exported reports. The generated value exists briefly as plain
    text so it can be displayed, then is copied character by character into a
    read-only SecureString before it is sent to AD.

.PARAMETER ShowGeneratedPassword
    Display a generated temporary password once in the console. This is required
    with -GenerateTemporaryPassword so password disclosure is explicit.

.PARAMETER ChangePasswordAtLogon
    Require the selected user to change password at next logon.

.NOTES
    Run from an elevated PowerShell session using an account allowed to query AD.
    Event audit features require permission to read Security logs on Domain
    Controllers.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSReviewUnusedParameter",
    "",
    Justification = "Script parameters are consumed by script-scoped helper functions, which this rule does not resolve."
)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseSingularNouns",
    "",
    Justification = "These established internal helpers intentionally return collections; renaming them would add compatibility churn."
)]
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [ValidateSet("Report", "UserAudit", "Reset", "LockedOut")]
    [string]$Mode = "Report",

    [string]$Identity,

    [string]$UserListPath,

    [string]$SearchBase,

    [string[]]$SearchBaseList,

    [string[]]$ExcludeOU,

    [string[]]$DomainControllers,

    [PSCredential]$Credential,

    [ValidateSet("UserSummary", "PasswordAge", "LockedOut", "AuditEvents", "PrivilegedUsers", "DisabledUsers", "StaleUsers")]
    [string[]]$ReportType = @("UserSummary"),

    [ValidateSet("Csv", "Json", "Html")]
    [string[]]$ExportFormat = @("Csv"),

    [ValidateRange(1, 3650)]
    [int]$DaysBack = 7,

    [ValidateRange(1, 3650)]
    [int]$PasswordAgeWarningDays = 90,

    [ValidateRange(1, 3650)]
    [int]$StaleUserDays = 90,

    [ValidateRange(0, 1000000)]
    [int]$MaxEventsPerDomainController = 0,

    [ValidateRange(1, 10000)]
    [int]$MaxAttributeValueLength = 500,

    [string[]]$PrivilegedGroupNames = (Get-DefaultPrivilegedGroupNames),

    [string]$OutputDirectory,

    [string]$OutputPrefix = "ADUsers",

    [switch]$Unlock,

    [switch]$Enable,

    [switch]$ResetPassword,

    [switch]$ChangePasswordAtLogon,

    [securestring]$NewPassword,

    [switch]$GenerateTemporaryPassword,

    [switch]$ShowGeneratedPassword,

    [switch]$IncludeEvents,

    [switch]$IncludeGroupMembership,

    [switch]$IncludeDisabled,

    [switch]$IncludeMessage,

    [switch]$AllowPartialResults,

    [switch]$AllowUnverifiedDomainController,

    [switch]$NoClobber,

    [switch]$ForceOverwrite,

    [switch]$AllowNetworkOutputPath,

    [switch]$AllowNetworkInputPath,

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

$UserEventIds = @(
    4720, # User created
    4722, # User enabled
    4723, # Password change attempted
    4724, # Password reset attempted
    4725, # User disabled
    4726, # User deleted
    4738, # User changed
    4740, # User locked out
    4767, # User unlocked
    4781  # Account name changed
)

$UserEventActionMap = @{
    4720 = "User account created"
    4722 = "User account enabled"
    4723 = "Password change attempted"
    4724 = "Password reset attempted"
    4725 = "User account disabled"
    4726 = "User account deleted"
    4738 = "User account changed"
    4740 = "User account locked out"
    4767 = "User account unlocked"
    4781 = "Account name changed"
}

if ($ForceOverwrite -and $NoClobber) {
    throw "ForceOverwrite and NoClobber cannot be used together."
}

function Get-AdCommandParameters {
    $Parameters = @{}
    if ($Credential) {
        $Parameters["Credential"] = $Credential
    }

    return $Parameters
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

    return Get-ADDomainController -Filter * @AdCommandParameters |
        Where-Object { -not $_.IsReadOnly } |
        Select-Object -ExpandProperty HostName
}

function ConvertTo-UserAuditRecord {
    param(
        [System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord,
        [string]$DomainController,
        [switch]$IncludeRenderedMessage
    )

    $EventData = ConvertTo-EventDataMap -EventRecord $EventRecord
    $EventId = [int]$EventRecord.Id
    $TargetUserName = Get-MapValue -Map $EventData -Key "TargetUserName"
    $TargetDomainName = Get-MapValue -Map $EventData -Key "TargetDomainName"
    $SubjectUserName = Get-MapValue -Map $EventData -Key "SubjectUserName"
    $SubjectDomainName = Get-MapValue -Map $EventData -Key "SubjectDomainName"
    $AttributeName = Get-MapValue -Map $EventData -Key "AttributeLDAPDisplayName"
    $AttributeValue = Limit-TextLength -Text (Get-MapValue -Map $EventData -Key "AttributeValue") -MaximumLength $MaxAttributeValueLength
    $ChangedAttributes = $null
    $RenderedMessage = $null

    if ($AttributeName) {
        $ChangedAttributes = "$AttributeName=$AttributeValue"
    }

    if ($IncludeRenderedMessage) {
        $RenderedMessage = $EventRecord.Message
    }

    [pscustomobject]@{
        TimeCreated        = $EventRecord.TimeCreated
        DomainController   = $DomainController
        EventId            = $EventId
        Action             = $UserEventActionMap[$EventId]
        TargetUser         = $TargetUserName
        TargetDomain       = $TargetDomainName
        TargetAccount      = Join-DomainAccount -DomainName $TargetDomainName -AccountName $TargetUserName
        SubjectUser        = $SubjectUserName
        SubjectDomain      = $SubjectDomainName
        SubjectAccount     = Join-DomainAccount -DomainName $SubjectDomainName -AccountName $SubjectUserName
        SubjectLogonId     = Get-MapValue -Map $EventData -Key "SubjectLogonId"
        CallerComputerName = Get-MapValue -Map $EventData -Key "CallerComputerName"
        TargetSid          = Get-MapValue -Map $EventData -Key "TargetSid"
        AttributeName      = $AttributeName
        AttributeValue     = $AttributeValue
        ChangedAttributes  = $ChangedAttributes
        EventRecordId      = $EventRecord.RecordId
        RenderedMessage    = $RenderedMessage
    }
}

function Test-UserMatchesAuditRecord {
    param(
        [psobject]$Record,
        [psobject]$User
    )

    if ($null -eq $User) {
        return $true
    }

    $Identifiers = @(
        $User.SamAccountName,
        $User.UserPrincipalName,
        $User.SID.Value,
        $User.DistinguishedName
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    foreach ($Identifier in $Identifiers) {
        if ($Record.TargetUser -ieq $Identifier -or
            $Record.TargetAccount -ieq $Identifier -or
            $Record.TargetSid -ieq $Identifier) {
            return $true
        }
    }

    return $false
}

function Get-UserAuditEvents {
    param(
        [AllowNull()]
        [psobject]$User
    )

    $ResolvedDomainControllers = Get-DomainControllerNames `
        -ProvidedDomainControllers $DomainControllers `
        -AllowUnverified:$AllowUnverifiedDomainController `
        -Credential $Credential

    if (@($ResolvedDomainControllers).Count -eq 0) {
        throw "No writable Domain Controllers were found. Verify AD connectivity and that the account has permission to enumerate domain controllers."
    }

    if (-not [string]::IsNullOrWhiteSpace($Identity) -and $MaxEventsPerDomainController -gt 0) {
        Write-Warning "UserAudit event filtering is applied after MaxEventsPerDomainController cap. Set MaxEventsPerDomainController to 0 to ensure all events for the target user are returned."
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
                -Ids $UserEventIds `
                -ProviderName $SecurityAuditProviderName `
                -StartDate $StartTime `
                -MaxEvents $MaxEventsPerDomainController `
                -Credential $Credential

            [void]$SuccessfulDomainControllers.Add($DomainController)
            $SecurityEvents | ForEach-Object {
                $Record = ConvertTo-UserAuditRecord `
                    -EventRecord $_ `
                    -DomainController $DomainController `
                    -IncludeRenderedMessage:$IncludeMessage

                if (Test-UserMatchesAuditRecord -Record $Record -User $User) {
                    $Record
                }
            }
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

    return @($AllRecords | Sort-Object TimeCreated -Descending)
}

function Assert-InputPathAllowed {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Purpose
    )

    if ((Test-IsUncPath -Path $Path) -and -not $AllowNetworkInputPath) {
        throw "Network $Purpose paths are not allowed by default: $Path. Re-run with -AllowNetworkInputPath only if the location is trusted."
    }
}

function Get-SafeFileNamePart {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "domain"
    }

    return ($Value -replace '[^a-zA-Z0-9._-]', '_')
}

function Get-DefaultOutputDirectory {
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
        return Join-Path -Path $env:LOCALAPPDATA -ChildPath "AdminTools\ADUserAccountReports"
    }

    return Join-Path -Path $ScriptDirectory -ChildPath "ADUserAccountReports"
}

function Export-AccountReport {
    param(
        [string]$ReportName,
        [object[]]$Data,
        [string[]]$Formats
    )

    $Rows = @($Data)
    $BaseName = "{0}_{1}_{2}_{3}" -f (Get-SafeFileNamePart -Value $OutputPrefix), $ReportName, (Get-SafeFileNamePart -Value $DomainName), $RunTimestamp
    $BasePath = Join-Path -Path $OutputDirectory -ChildPath $BaseName
    $ExportedPaths = New-Object System.Collections.Generic.List[string]

    foreach ($Format in $Formats) {
        $TargetPath = "{0}.{1}" -f $BasePath, $Format.ToLowerInvariant()

        if ((Test-IsUncPath -Path $TargetPath) -and -not $AllowNetworkOutputPath) {
            throw "Network output paths are not allowed by default: $TargetPath. Re-run with -AllowNetworkOutputPath only if the location is trusted."
        }

        if (Test-Path -LiteralPath $TargetPath) {
            if ($NoClobber -or -not $ForceOverwrite) {
                throw "Output file already exists: $TargetPath. Re-run with -ForceOverwrite only if replacing it is intended."
            }

            Write-Warning "Overwriting existing report file: $TargetPath"
        }

        if ($WhatIfPreference) {
            Write-Information "WhatIf: would export $ReportName -> $TargetPath"
            continue
        }

        try {
            switch ($Format) {
                "Csv" {
                    $CsvRows = if ($DisableCsvSanitization) {
                        Write-Warning "CSV sanitization is disabled. Opening this report in spreadsheet software may evaluate account data as formulas."
                        $Rows
                    }
                    else {
                        @($Rows | ConvertTo-SafeCsvRecord)
                    }

                    if ($CsvRows.Count -gt 0) {
                        $CsvRows | Export-Csv -LiteralPath $TargetPath -NoTypeInformation -Encoding UTF8
                    }
                    else {
                        Set-Content -LiteralPath $TargetPath -Value "" -Encoding UTF8
                    }
                }
                "Json" {
                    $Json = if ($Rows.Count -gt 0) {
                        $Rows | ConvertTo-Json -Depth 8
                    }
                    else {
                        "[]"
                    }

                    Set-Content -LiteralPath $TargetPath -Value $Json -Encoding UTF8
                }
                "Html" {
                    $Title = "AD User $ReportName Report"
                    $Style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }
h1 { margin-bottom: 4px; }
p.meta { color: #6b7280; margin-top: 0; }
table { border-collapse: collapse; width: 100%; font-size: 12px; }
th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; vertical-align: top; }
th { background: #f3f4f6; }
tr:nth-child(even) { background: #f9fafb; }
</style>
"@

                    $Body = if ($Rows.Count -gt 0) {
                        $Rows | ConvertTo-Html -Head $Style -Title $Title -PreContent "<h1>$Title</h1><p class='meta'>Generated $RunStartedAt</p>"
                    }
                    else {
                        @"
<html>
<head>$Style<title>$Title</title></head>
<body>
<h1>$Title</h1>
<p class='meta'>Generated $RunStartedAt</p>
<p>No records found.</p>
</body>
</html>
"@
                    }

                    Set-Content -LiteralPath $TargetPath -Value $Body -Encoding UTF8
                }
            }
        }
        catch {
            Write-Error "Failed to export report to '$TargetPath': $($_.Exception.Message)"
            exit $ExitCodes.Export
        }

        [void]$ExportedPaths.Add($TargetPath)
        Write-Information "Exported $ReportName report: $TargetPath"
    }

    return @($ExportedPaths.ToArray())
}

function Read-UserList {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "UserListPath was not found: $Path"
    }

    Assert-InputPathAllowed -Path $Path -Purpose "user list"

    return @(Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") })
}

function Resolve-ManagedADUser {
    param(
        [string]$UserIdentity
    )

    $Properties = @(
        "CanonicalName",
        "Created",
        "Modified",
        "PasswordLastSet",
        "PasswordNeverExpires",
        "CannotChangePassword",
        "LockedOut",
        "LastLogonDate",
        "LastBadPasswordAttempt",
        "AccountExpirationDate",
        "Enabled",
        "AdminCount",
        "MemberOf",
        "PrimaryGroupID",
        "UserPrincipalName",
        "DisplayName",
        "Description",
        "EmailAddress",
        "SID"
    )

    $AdCommandParameters = Get-AdCommandParameters

    return Get-ADUser -Identity $UserIdentity -Properties $Properties @AdCommandParameters -ErrorAction Stop
}

function Get-UserQueryProperties {
    return @(
        "CanonicalName",
        "Created",
        "Modified",
        "PasswordLastSet",
        "PasswordNeverExpires",
        "CannotChangePassword",
        "LockedOut",
        "LastLogonDate",
        "LastBadPasswordAttempt",
        "AccountExpirationDate",
        "Enabled",
        "AdminCount",
        "MemberOf",
        "PrimaryGroupID",
        "UserPrincipalName",
        "DisplayName",
        "Description",
        "EmailAddress",
        "SID"
    )
}

function Get-ScopedUsers {
    param(
        [string]$Filter = "*"
    )

    $AdCommandParameters = Get-AdCommandParameters
    $Properties = Get-UserQueryProperties
    $SearchBases = @()

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $SearchBases += $SearchBase
    }

    if ($SearchBaseList) {
        $SearchBases += $SearchBaseList
    }

    $SeenGuids = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $UniqueUsers = New-Object System.Collections.Generic.List[object]

    if ($SearchBases.Count -eq 0) {
        $RawUsers = @(Get-ADUser -Filter $Filter -Properties $Properties @AdCommandParameters)
    }
    else {
        $RawUsers = foreach ($Scope in $SearchBases) {
            Get-ADUser -Filter $Filter -SearchBase $Scope -Properties $Properties @AdCommandParameters
        }
    }

    foreach ($User in @($RawUsers)) {
        $Guid = [string]$User.ObjectGUID
        if ($SeenGuids.Add($Guid)) {
            [void]$UniqueUsers.Add($User)
        }
    }

    $Users = $UniqueUsers.ToArray()

    if (-not $IncludeDisabled) {
        $Users = @($Users | Where-Object { $_.Enabled })
    }

    foreach ($ExcludedOu in $ExcludeOU) {
        if (-not [string]::IsNullOrWhiteSpace($ExcludedOu)) {
            $Users = @($Users | Where-Object {
                -not $_.DistinguishedName.EndsWith($ExcludedOu, [System.StringComparison]::OrdinalIgnoreCase)
            })
        }
    }

    return @($Users | Sort-Object SamAccountName)
}

function Get-TargetUsers {
    $Targets = New-Object System.Collections.Generic.List[object]

    if (-not [string]::IsNullOrWhiteSpace($Identity)) {
        [void]$Targets.Add((Resolve-ManagedADUser -UserIdentity $Identity))
    }

    foreach ($UserIdentity in (Read-UserList -Path $UserListPath)) {
        [void]$Targets.Add((Resolve-ManagedADUser -UserIdentity $UserIdentity))
    }

    if ($Targets.Count -gt 0) {
        return @($Targets.ToArray() | Sort-Object SamAccountName -Unique)
    }

    return Get-ScopedUsers
}

function Get-UserAccountRecord {
    param(
        [psobject]$User
    )

    $Now = Get-Date
    $AccountAgeDays = if ($User.Created) { [int]($Now - $User.Created).TotalDays } else { $null }
    $PasswordAgeDays = if ($User.PasswordLastSet) { [int]($Now - $User.PasswordLastSet).TotalDays } else { $null }
    $DaysSinceLastLogon = if ($User.LastLogonDate) { [int]($Now - $User.LastLogonDate).TotalDays } else { $null }
    $PasswordAgeStatus = "Unknown"
    $MemberOf = $null

    if ($null -ne $PasswordAgeDays -and $PasswordAgeDays -ge $PasswordAgeWarningDays) {
        $PasswordAgeStatus = "Warning"
    }
    elseif ($null -ne $PasswordAgeDays) {
        $PasswordAgeStatus = "OK"
    }

    if ($IncludeGroupMembership) {
        $MemberOf = @($User.MemberOf) -join "; "
    }

    [pscustomobject]@{
        SamAccountName           = $User.SamAccountName
        UserPrincipalName        = $User.UserPrincipalName
        DisplayName              = $User.DisplayName
        Enabled                  = $User.Enabled
        LockedOut                = $User.LockedOut
        DistinguishedName        = $User.DistinguishedName
        CanonicalName            = $User.CanonicalName
        Description              = $User.Description
        EmailAddress             = $User.EmailAddress
        Created                  = $User.Created
        AccountAgeDays           = $AccountAgeDays
        Modified                 = $User.Modified
        PasswordLastSet          = $User.PasswordLastSet
        PasswordAgeDays          = $PasswordAgeDays
        PasswordAgeWarningDays   = $PasswordAgeWarningDays
        PasswordAgeStatus        = $PasswordAgeStatus
        PasswordNeverExpires     = $User.PasswordNeverExpires
        CannotChangePassword     = $User.CannotChangePassword
        LastLogonDate            = $User.LastLogonDate
        DaysSinceLastLogon       = $DaysSinceLastLogon
        LastBadPasswordAttempt   = $User.LastBadPasswordAttempt
        AccountExpirationDate    = $User.AccountExpirationDate
        AdminCount               = $User.AdminCount
        MemberOfCount            = @($User.MemberOf).Count
        MemberOf                 = $MemberOf
        ObjectSID                = [string]$User.SID
        ObjectGUID               = [string]$User.ObjectGUID
        QueriedAt                = $RunStartedAt
    }
}

function Get-UserAccountSummaryReport {
    param(
        [object[]]$Users
    )

    return @($Users | ForEach-Object { Get-UserAccountRecord -User $_ })
}

function Get-PasswordAgeReport {
    param(
        [object[]]$Users
    )

    return @(Get-UserAccountSummaryReport -Users $Users |
        Select-Object SamAccountName, DisplayName, Enabled, LockedOut, Created, AccountAgeDays, PasswordLastSet, PasswordAgeDays, PasswordAgeWarningDays, PasswordAgeStatus, PasswordNeverExpires, CannotChangePassword, LastLogonDate, DaysSinceLastLogon, DistinguishedName, QueriedAt)
}

function Get-LockedOutUserReport {
    $AdCommandParameters = Get-AdCommandParameters
    $Properties = Get-UserQueryProperties
    $SearchBases = @()

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $SearchBases += $SearchBase
    }

    if ($SearchBaseList) {
        $SearchBases += $SearchBaseList
    }

    $SeenGuids = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $Users = New-Object System.Collections.Generic.List[object]

    if ($SearchBases.Count -eq 0) {
        $LockedAccounts = @(Search-ADAccount -LockedOut -UsersOnly @AdCommandParameters)
        foreach ($Account in $LockedAccounts) {
            try {
                $User = Get-ADUser -Identity $Account.DistinguishedName -Properties $Properties @AdCommandParameters -ErrorAction Stop
                $Guid = [string]$User.ObjectGUID
                if ($SeenGuids.Add($Guid)) {
                    [void]$Users.Add($User)
                }
            }
            catch {
                Write-Warning "Could not resolve locked-out user '$($Account.DistinguishedName)'. Error: $($_.Exception.Message)"
            }
        }
    }
    else {
        foreach ($Scope in $SearchBases) {
            $LockedAccounts = @(Search-ADAccount -LockedOut -UsersOnly -SearchBase $Scope @AdCommandParameters)
            foreach ($Account in $LockedAccounts) {
                try {
                    $User = Get-ADUser -Identity $Account.DistinguishedName -Properties $Properties @AdCommandParameters -ErrorAction Stop
                    $Guid = [string]$User.ObjectGUID
                    if ($SeenGuids.Add($Guid)) {
                        [void]$Users.Add($User)
                    }
                }
                catch {
                    Write-Warning "Could not resolve locked-out user '$($Account.DistinguishedName)'. Error: $($_.Exception.Message)"
                }
            }
        }
    }

    $FilteredUsers = @($Users | Where-Object {
        $Dn = $_.DistinguishedName
        $Excluded = $false
        foreach ($ExcludedOu in $ExcludeOU) {
            if (-not [string]::IsNullOrWhiteSpace($ExcludedOu) -and
                $Dn.EndsWith($ExcludedOu, [System.StringComparison]::OrdinalIgnoreCase)) {
                $Excluded = $true
                break
            }
        }
        -not $Excluded
    })

    return @(Get-UserAccountSummaryReport -Users $FilteredUsers |
        Select-Object SamAccountName, DisplayName, UserPrincipalName, LockedOut,
            LastBadPasswordAttempt, LastLogonDate, PasswordLastSet, Enabled,
            DistinguishedName, QueriedAt)
}

function Get-StaleUserReport {
    param(
        [object[]]$Users
    )

    return @(Get-UserAccountSummaryReport -Users $Users |
        Where-Object { $null -eq $_.LastLogonDate -or $_.DaysSinceLastLogon -ge $StaleUserDays } |
        Select-Object SamAccountName, DisplayName, Enabled, Created, AccountAgeDays, LastLogonDate, DaysSinceLastLogon, @{ Name = "StaleUserDays"; Expression = { $StaleUserDays } }, PasswordLastSet, DistinguishedName, QueriedAt)
}

function Get-DisabledUserReport {
    $AdCommandParameters = Get-AdCommandParameters
    $Properties = Get-UserQueryProperties
    $SearchBases = @()

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
        $SearchBases += $SearchBase
    }

    if ($SearchBaseList) {
        $SearchBases += $SearchBaseList
    }

    if ($SearchBases.Count -eq 0) {
        $Users = @(Get-ADUser -Filter "Enabled -eq `$false" -Properties $Properties @AdCommandParameters)
    }
    else {
        $Users = foreach ($Scope in $SearchBases) {
            Get-ADUser -Filter "Enabled -eq `$false" -SearchBase $Scope -Properties $Properties @AdCommandParameters
        }
    }

    foreach ($ExcludedOu in $ExcludeOU) {
        if (-not [string]::IsNullOrWhiteSpace($ExcludedOu)) {
            $Users = @($Users | Where-Object { -not $_.DistinguishedName.EndsWith($ExcludedOu, [System.StringComparison]::OrdinalIgnoreCase) })
        }
    }

    return @(Get-UserAccountSummaryReport -Users $Users |
        Where-Object { -not $_.Enabled } |
        Select-Object SamAccountName, DisplayName, Enabled, LockedOut, Created, LastLogonDate, PasswordLastSet, DistinguishedName, QueriedAt)
}

function Get-PrivilegedUserReport {
    $AdCommandParameters = Get-AdCommandParameters
    $Rows = New-Object System.Collections.Generic.List[object]

    foreach ($GroupName in $PrivilegedGroupNames) {
        try {
            $Group = Get-ADGroup -Identity $GroupName @AdCommandParameters -ErrorAction Stop
            $Members = @(Get-ADGroupMember -Identity $Group.DistinguishedName -Recursive @AdCommandParameters -ErrorAction Stop |
                    Where-Object { $_.ObjectClass -eq "user" })

            foreach ($Member in $Members) {
                try {
                    $User = Resolve-ManagedADUser -UserIdentity $Member.DistinguishedName
                    [void]$Rows.Add([pscustomobject]@{
                            PrivilegedGroup    = $Group.Name
                            SamAccountName     = $User.SamAccountName
                            DisplayName        = $User.DisplayName
                            Enabled            = $User.Enabled
                            LockedOut          = $User.LockedOut
                            PasswordLastSet    = $User.PasswordLastSet
                            LastLogonDate      = $User.LastLogonDate
                            AdminCount         = $User.AdminCount
                            DistinguishedName  = $User.DistinguishedName
                            QueriedAt          = $RunStartedAt
                        })
                }
                catch {
                    Write-Warning "Could not resolve privileged member '$($Member.DistinguishedName)'. Error: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Warning "Could not read privileged group '$GroupName'. Error: $($_.Exception.Message)"
        }
    }

    return @($Rows.ToArray() | Sort-Object PrivilegedGroup, SamAccountName)
}

function New-TemporaryPassword {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "This pure helper returns a random string and does not change system state."
    )]
    param()

    $Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"
    $Symbols = "!#$%&*+-=?@"
    $Bytes = New-Object byte[] 18
    $Rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $Rng.GetBytes($Bytes)
    }
    finally {
        $Rng.Dispose()
    }

    $Chars = for ($Index = 0; $Index -lt $Bytes.Length; $Index++) {
        if ($Index % 6 -eq 5) {
            $Symbols[$Bytes[$Index] % $Symbols.Length]
        }
        else {
            $Alphabet[$Bytes[$Index] % $Alphabet.Length]
        }
    }

    return -join $Chars
}

function Invoke-UserAccountResetAction {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$User,

        [switch]$UnlockAccount,

        [switch]$EnableAccount,

        [switch]$ResetUserPassword,

        [switch]$RequireChangePasswordAtLogon,

        [securestring]$Password,

        [switch]$GeneratePassword,

        [switch]$RevealGeneratedPassword
    )

    if (-not ($UnlockAccount -or $EnableAccount -or $ResetUserPassword -or $RequireChangePasswordAtLogon)) {
        throw "Mode Reset requires at least one action: -Unlock, -Enable, -ResetPassword, or -ChangePasswordAtLogon."
    }

    $AdCommandParameters = Get-AdCommandParameters
    $ActionResults = New-Object System.Collections.Generic.List[object]
    $TemporaryPassword = $null

    if ($ResetUserPassword) {
        if ($GeneratePassword -and $Password) {
            throw "Use either -NewPassword or -GenerateTemporaryPassword, not both."
        }

        if (-not $Password -and -not $GeneratePassword) {
            throw "ResetPassword requires -NewPassword or -GenerateTemporaryPassword."
        }

        if ($GeneratePassword -and -not $RevealGeneratedPassword) {
            throw "GenerateTemporaryPassword requires -ShowGeneratedPassword so console password disclosure is explicit."
        }
    }

    if ($UnlockAccount) {
        if ($PSCmdlet.ShouldProcess($User.SamAccountName, "Unlock AD user account")) {
            Unlock-ADAccount -Identity $User.DistinguishedName @AdCommandParameters
            [void]$ActionResults.Add([pscustomobject]@{
                    SamAccountName = $User.SamAccountName
                    Action         = "Unlock"
                    Status         = "Completed"
                    Timestamp      = Get-Date
                })
        }
    }

    if ($EnableAccount) {
        if ($PSCmdlet.ShouldProcess($User.SamAccountName, "Enable AD user account")) {
            Enable-ADAccount -Identity $User.DistinguishedName @AdCommandParameters
            [void]$ActionResults.Add([pscustomobject]@{
                    SamAccountName = $User.SamAccountName
                    Action         = "Enable"
                    Status         = "Completed"
                    Timestamp      = Get-Date
                })
        }
    }

    if ($ResetUserPassword) {
        if ($PSCmdlet.ShouldProcess($User.SamAccountName, "Reset AD user password")) {
            $PasswordToSet = $Password
            if ($GeneratePassword) {
                $TemporaryPassword = New-TemporaryPassword

                # The generated string exists only for the explicitly requested
                # one-time display. Copy it directly into a read-only SecureString
                # without using ConvertTo-SecureString -AsPlainText.
                $PasswordToSet = [securestring]::new()
                foreach ($Character in $TemporaryPassword.ToCharArray()) {
                    $PasswordToSet.AppendChar($Character)
                }
                $PasswordToSet.MakeReadOnly()
            }

            Set-ADAccountPassword -Identity $User.DistinguishedName -Reset -NewPassword $PasswordToSet @AdCommandParameters
            [void]$ActionResults.Add([pscustomobject]@{
                    SamAccountName = $User.SamAccountName
                    Action         = "ResetPassword"
                    Status         = "Completed"
                    Timestamp      = Get-Date
                })
        }
    }

    if ($RequireChangePasswordAtLogon) {
        if ($PSCmdlet.ShouldProcess($User.SamAccountName, "Require password change at next logon")) {
            Set-ADUser -Identity $User.DistinguishedName -ChangePasswordAtLogon $true @AdCommandParameters
            [void]$ActionResults.Add([pscustomobject]@{
                    SamAccountName = $User.SamAccountName
                    Action         = "ChangePasswordAtLogon"
                    Status         = "Completed"
                    Timestamp      = Get-Date
                })
        }
    }

    if ($TemporaryPassword) {
        Write-Warning "Generated temporary password for $($User.SamAccountName): $TemporaryPassword"
        Write-Warning "The temporary password is displayed once and is not exported."
    }

    return @($ActionResults.ToArray())
}

if (-not (Import-ActiveDirectoryModule)) {
    throw "The trusted ActiveDirectory module is required. Install RSAT Active Directory tools and retry from a Windows admin workstation."
}

Assert-SafeDomainControllerNames -Names $DomainControllers
Assert-SafeTextValues -Purpose "Identity" -Values @($Identity) -MaximumLength 1024
Assert-SafeTextValues -Purpose "SearchBase" -Values @($SearchBase)
Assert-SafeTextValues -Purpose "SearchBaseList" -Values $SearchBaseList
Assert-SafeTextValues -Purpose "ExcludeOU" -Values $ExcludeOU
Assert-SafeTextValues -Purpose "PrivilegedGroupNames" -Values $PrivilegedGroupNames -MaximumLength 256
Assert-SafeTextValues -Purpose "OutputPrefix" -Values @($OutputPrefix) -MaximumLength 128

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Get-DefaultOutputDirectory
}

if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path -Path (Get-Location).Path -ChildPath $OutputDirectory
}

if ((Test-IsUncPath -Path $OutputDirectory) -and -not $AllowNetworkOutputPath) {
    throw "Network output paths are not allowed by default: $OutputDirectory. Re-run with -AllowNetworkOutputPath only if the location is trusted."
}

if (-not $WhatIfPreference -and -not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$AdCommandParameters = Get-AdCommandParameters
$AdDomain = Get-ADDomain @AdCommandParameters
$DomainName = $AdDomain.DNSRoot

if (-not (Test-IsSafeDnsName -Name $DomainName)) {
    throw "Discovered AD DNS root is not a valid DNS name: $DomainName"
}

$ExportedPaths = New-Object System.Collections.Generic.List[string]

switch ($Mode) {
    "Report" {
        $UserDependentReportTypes = @("UserSummary", "PasswordAge", "StaleUsers")
        $NeedsUsers = $ReportType | Where-Object { $UserDependentReportTypes -contains $_ }
        $Users = if ($NeedsUsers) { Get-TargetUsers } else { @() }

        foreach ($CurrentReportType in $ReportType) {
            $ReportData = switch ($CurrentReportType) {
                "UserSummary"    { Get-UserAccountSummaryReport -Users $Users }
                "PasswordAge"    { Get-PasswordAgeReport -Users $Users }
                "LockedOut"      { Get-LockedOutUserReport }
                "AuditEvents"    { Get-UserAuditEvents -User $null }
                "PrivilegedUsers" { Get-PrivilegedUserReport }
                "DisabledUsers"  { Get-DisabledUserReport }
                "StaleUsers"     { Get-StaleUserReport -Users $Users }
            }

            $Paths = Export-AccountReport -ReportName $CurrentReportType -Data $ReportData -Formats $ExportFormat
            foreach ($Path in $Paths) {
                [void]$ExportedPaths.Add($Path)
            }

            Write-Information ("{0}: {1} record(s)" -f $CurrentReportType, @($ReportData).Count)
        }
    }
    "UserAudit" {
        if ([string]::IsNullOrWhiteSpace($Identity)) {
            throw "Mode UserAudit requires -Identity."
        }

        $User = Resolve-ManagedADUser -UserIdentity $Identity
        $Summary = @(Get-UserAccountSummaryReport -Users @($User))
        foreach ($Path in (Export-AccountReport -ReportName "UserAuditSummary" -Data $Summary -Formats $ExportFormat)) {
            [void]$ExportedPaths.Add($Path)
        }

        if ($IncludeEvents) {
            $Events = Get-UserAuditEvents -User $User
            foreach ($Path in (Export-AccountReport -ReportName "UserAuditEvents" -Data $Events -Formats $ExportFormat)) {
                [void]$ExportedPaths.Add($Path)
            }
            Write-Information ("UserAuditEvents: {0} record(s)" -f @($Events).Count)
        }

        Write-Information ("UserAuditSummary: {0} record(s)" -f $Summary.Count)
    }
    "LockedOut" {
        $LockedUsers = Get-LockedOutUserReport
        foreach ($Path in (Export-AccountReport -ReportName "LockedOut" -Data $LockedUsers -Formats $ExportFormat)) {
            [void]$ExportedPaths.Add($Path)
        }

        if ($IncludeEvents) {
            # Build a set of identifiers for currently-locked accounts so that
            # lockout events are restricted to those accounts only.
            $LockedIdentifiers = [System.Collections.Generic.HashSet[string]]::new(
                [System.StringComparer]::OrdinalIgnoreCase
            )
            foreach ($LockedUser in $LockedUsers) {
                if (-not [string]::IsNullOrWhiteSpace($LockedUser.SamAccountName)) {
                    [void]$LockedIdentifiers.Add($LockedUser.SamAccountName)
                }
                if (-not [string]::IsNullOrWhiteSpace($LockedUser.UserPrincipalName)) {
                    [void]$LockedIdentifiers.Add($LockedUser.UserPrincipalName)
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$LockedUser.ObjectSID)) {
                    [void]$LockedIdentifiers.Add([string]$LockedUser.ObjectSID)
                }
            }

            $AllLockoutEvents = @(Get-UserAuditEvents -User $null | Where-Object { $_.EventId -eq 4740 })
            $LockoutEvents = if ($LockedIdentifiers.Count -gt 0) {
                @($AllLockoutEvents | Where-Object {
                    $LockedIdentifiers.Contains($_.TargetUser) -or
                    $LockedIdentifiers.Contains($_.TargetAccount) -or
                    $LockedIdentifiers.Contains($_.TargetSid)
                })
            } else {
                @()
            }

            foreach ($Path in (Export-AccountReport -ReportName "LockedOutEvents" -Data $LockoutEvents -Formats $ExportFormat)) {
                [void]$ExportedPaths.Add($Path)
            }
            Write-Information ("LockedOutEvents: {0} record(s) (filtered to currently-locked accounts)" -f $LockoutEvents.Count) -InformationAction Continue
        }

        Write-Information ("LockedOut: {0} record(s)" -f @($LockedUsers).Count)
    }
    "Reset" {
        if ([string]::IsNullOrWhiteSpace($Identity)) {
            throw "Mode Reset requires -Identity."
        }

        if (-not [string]::IsNullOrWhiteSpace($UserListPath)) {
            throw "Mode Reset supports one -Identity at a time. Bulk reset from UserListPath is intentionally not enabled."
        }

        $User = Resolve-ManagedADUser -UserIdentity $Identity
        $ActionResults = Invoke-UserAccountResetAction `
            -User $User `
            -UnlockAccount:$Unlock `
            -EnableAccount:$Enable `
            -ResetUserPassword:$ResetPassword `
            -RequireChangePasswordAtLogon:$ChangePasswordAtLogon `
            -Password $NewPassword `
            -GeneratePassword:$GenerateTemporaryPassword `
            -RevealGeneratedPassword:$ShowGeneratedPassword `
            -WhatIf:$WhatIfPreference

        foreach ($Path in (Export-AccountReport -ReportName "ResetActions" -Data $ActionResults -Formats $ExportFormat)) {
            [void]$ExportedPaths.Add($Path)
        }

        Write-Information ("ResetActions: {0} completed action(s)" -f @($ActionResults).Count)
    }
}

if ($ExportedPaths.Count -gt 0) {
    Write-Information "Exported file(s):"
    $ExportedPaths | ForEach-Object { Write-Information "  $_" }
}

exit $ExitCodes.Success
