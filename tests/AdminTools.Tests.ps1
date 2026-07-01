#requires -Version 5.1
#requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSUseDeclaredVarsMoreThanAssignments",
    "",
    Justification = "Pester BeforeAll variables are consumed in It blocks; ScriptAnalyzer does not model that scope flow."
)]
param()
BeforeAll {
    $ScriptRoot = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $ScriptRoot "AdminToolsCommon.psm1") -Force

    function Get-TestCredential {
        param(
            [Parameter(Mandatory = $true)]
            [string]$UserName
        )

        $securePassword = [securestring]::new()
        foreach ($character in @("p", "a", "s", "s")) {
            $securePassword.AppendChar($character)
        }
        $securePassword.MakeReadOnly()

        [PSCredential]::new($UserName, $securePassword)
    }
}

Describe "Test-IsSafeDnsName" {
    It "Returns true for a valid hostname" {
        Test-IsSafeDnsName -Name "dc01.domain.local" | Should -BeTrue
    }
    It "Returns false for a name with control characters" {
        Test-IsSafeDnsName -Name "dc01`n.domain.local" | Should -BeFalse
    }
    It "Returns false for a name exceeding 253 characters" {
        Test-IsSafeDnsName -Name ("a" * 254) | Should -BeFalse
    }
    It "Returns false for a label exceeding 63 characters" {
        Test-IsSafeDnsName -Name (("a" * 64) + ".domain.local") | Should -BeFalse
    }
}

Describe "ConvertTo-SafeCsvValue" {
    It "Prefixes values starting with equals sign" {
        ConvertTo-SafeCsvValue -Value "=cmd" | Should -Be "'=cmd"
    }
    It "Prefixes values starting with plus sign" {
        ConvertTo-SafeCsvValue -Value "+1234" | Should -Be "'+1234"
    }
    It "Does not modify safe string values" {
        ConvertTo-SafeCsvValue -Value "NormalValue" | Should -Be "NormalValue"
    }
    It "Returns non-string values unchanged" {
        ConvertTo-SafeCsvValue -Value 42 | Should -Be 42
    }
}

Describe "Limit-TextLength" {
    It "Returns text unchanged when within limit" {
        Limit-TextLength -Text "short" -MaximumLength 100 | Should -Be "short"
    }
    It "Truncates and appends marker when over limit" {
        $Result = Limit-TextLength -Text ("a" * 20) -MaximumLength 10
        $Result | Should -BeLike "*[truncated]*"
        $Result.Length | Should -BeGreaterThan 10
    }
}

Describe "Assert-SafeTextValues" {
    It "Throws when value contains control characters" {
        { Assert-SafeTextValues -Purpose "Test" -Values @("bad`0value") } | Should -Throw
    }
    It "Throws when value exceeds maximum length" {
        { Assert-SafeTextValues -Purpose "Test" -Values @("a" * 5000) -MaximumLength 100 } | Should -Throw
    }
    It "Does not throw for valid values" {
        { Assert-SafeTextValues -Purpose "Test" -Values @("valid") } | Should -Not -Throw
    }
}

Describe "ConvertTo-EventDataMap" {
    It "Returns an empty hashtable for malformed XML without throwing" {
        $FakeRecord = [PSCustomObject]@{}
        $FakeRecord | Add-Member -MemberType ScriptMethod -Name "ToXml" -Value { return "<<<not xml>>>" }
        $FakeRecord | Add-Member -MemberType ScriptProperty -Name "RecordId" -Value { 999 }
        $FakeRecord | Add-Member -MemberType ScriptProperty -Name "Id" -Value { 4720 }
        $Result = ConvertTo-EventDataMap -EventRecord $FakeRecord 3>$null
        $Result | Should -BeOfType [hashtable]
        $Result.Count | Should -Be 0
    }
}

Describe "Get-ResolvedOutputPath" {
    It "resolves relative output paths against the current location" {
        Push-Location -LiteralPath $TestDrive
        try {
            Get-ResolvedOutputPath -Path "logs\custom.log" |
                Should -Be (Join-Path -Path $TestDrive -ChildPath "logs\custom.log")
        }
        finally {
            Pop-Location
        }
    }

    It "returns null for blank paths and leaves rooted paths unchanged" {
        Get-ResolvedOutputPath -Path "" | Should -BeNullOrEmpty
        Get-ResolvedOutputPath -Path "   " | Should -BeNullOrEmpty

        $rootedPath = Join-Path -Path $TestDrive -ChildPath "custom.log"
        Get-ResolvedOutputPath -Path $rootedPath | Should -Be $rootedPath
    }
}

Describe "Audit partial report completeness metadata" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $GetAdminActivityScript = Get-Content -Raw (Join-Path $ScriptRoot "Get-ADAdminActivity.ps1")
        $ManageUsersScript = Get-Content -Raw (Join-Path $ScriptRoot "Manage-ADUserAccounts.ps1")
    }

    It "marks partial report paths before the file extension" {
        $partialPath = Get-PartialReportPath -Path (Join-Path $TestDrive "audit.csv")
        Split-Path -Leaf $partialPath | Should -Be "audit_PARTIAL.csv"
    }

    It "builds a manifest with DC coverage, query window, cap settings, and script version" {
        $start = [datetime]"2026-07-01T08:00:00Z"
        $end = [datetime]"2026-07-01T09:00:00Z"
        $manifest = Get-AuditCompletenessManifest `
            -ScriptName "Get-ADAdminActivity.ps1" `
            -ScriptVersion "1.0.0" `
            -QueryStartTime $start `
            -QueryEndTime $end `
            -DomainControllers @("dc01", "dc02") `
            -FailedDomainControllers @([pscustomobject]@{ DomainController = "dc02"; Error = "Access denied"; ErrorCategory = "PermissionDenied" }) `
            -MaxEventsPerDomainController 100000 `
            -UnlimitedEvents:$false `
            -OutputPaths @("audit_PARTIAL.csv") `
            -Partial:$true

        $manifest.IsPartial | Should -BeTrue
        $manifest.ScriptName | Should -Be "Get-ADAdminActivity.ps1"
        $manifest.ScriptVersion | Should -Be "1.0.0"
        $manifest.DomainControllersQueried | Should -Be @("dc01", "dc02")
        $manifest.DomainControllersFailed[0].DomainController | Should -Be "dc02"
        $manifest.DomainControllersFailed[0].ErrorCategory | Should -Be "PermissionDenied"
        $manifest.QueryWindow.StartTimeUtc | Should -Be $start.ToUniversalTime().ToString("o")
        $manifest.QueryWindow.EndTimeUtc | Should -Be $end.ToUniversalTime().ToString("o")
        $manifest.CapSettings.MaxEventsPerDomainController | Should -Be 100000
        $manifest.CapSettings.UnlimitedEvents | Should -BeFalse
    }

    It "exports a JSON sidecar manifest beside a report path" {
        $reportPath = Join-Path $TestDrive "audit_PARTIAL.csv"
        Set-Content -LiteralPath $reportPath -Value "" -Encoding UTF8
        $manifest = Get-AuditCompletenessManifest `
            -ScriptName "Get-ADAdminActivity.ps1" `
            -ScriptVersion "1.0.0" `
            -QueryStartTime ([datetime]"2026-07-01T08:00:00Z") `
            -QueryEndTime ([datetime]"2026-07-01T09:00:00Z") `
            -DomainControllers @("dc01") `
            -FailedDomainControllers @() `
            -MaxEventsPerDomainController 100000 `
            -UnlimitedEvents:$false `
            -OutputPaths @($reportPath) `
            -Partial:$true

        $manifestPath = Export-AuditCompletenessManifest -Manifest $manifest -ReportPath $reportPath
        Split-Path -Leaf $manifestPath | Should -Be "audit_PARTIAL.manifest.json"
        $json = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
        $json.ScriptName | Should -Be "Get-ADAdminActivity.ps1"
        $json.IsPartial | Should -BeTrue
    }

    It "wires partial manifest metadata, partial filenames, and partial-success exit codes in both audit scripts" {
        foreach ($ScriptText in @($GetAdminActivityScript, $ManageUsersScript)) {
            $ScriptText | Should -Match 'PartialSuccess\s*=\s*8'
            $ScriptText | Should -Match 'Get-PartialReportPath'
            $ScriptText | Should -Match 'Get-AuditCompletenessManifest'
            $ScriptText | Should -Match 'Export-AuditCompletenessManifest'
            $ScriptText | Should -Match 'exit \$ExitCodes\.PartialSuccess'
            $ScriptText | Should -Match 'ScriptVersion'
        }

        $ManageUsersScript | Should -Match '\$script:LastAuditFailedDomainControllers'
    }
}

Describe "Security audit outcome handling" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $GetAdminActivityScript = Get-Content -Raw (Join-Path $ScriptRoot "Get-ADAdminActivity.ps1")
        $ManageUsersScript = Get-Content -Raw (Join-Path $ScriptRoot "Manage-ADUserAccounts.ps1")
    }

    It "derives AuditOutcome from event keyword display names" {
        Get-AuditOutcome -EventRecord ([pscustomobject]@{ KeywordsDisplayNames = @("Audit Success") }) | Should -Be "Success"
        Get-AuditOutcome -EventRecord ([pscustomobject]@{ KeywordsDisplayNames = @("Audit Failure") }) | Should -Be "Failure"
        Get-AuditOutcome -EventRecord ([pscustomobject]@{ KeywordsDisplayNames = @("Classic") }) | Should -Be "Unknown"
    }

    It "labels attempted audit actions with the resolved outcome" {
        Format-AuditAction -Action "Password reset attempted" -AuditOutcome "Failure" |
            Should -Be "Password reset attempted - outcome: Failure"
        Format-AuditAction -Action "User account created" -AuditOutcome "Success" |
            Should -Be "User account created"
    }

    It "adds AuditOutcome export and filter support to admin and user audit reports" {
        foreach ($ScriptText in @($GetAdminActivityScript, $ManageUsersScript)) {
            $ScriptText | Should -Match '\[ValidateSet\("Success", "Failure", "Unknown"\)\]'
            $ScriptText | Should -Match '\[string\[\]\]\$AuditOutcome'
            $ScriptText | Should -Match 'AuditOutcome\s*=\s*\$AuditOutcomeValue'
            $ScriptText | Should -Match 'Action\s*=\s*Format-AuditAction -Action .+ -AuditOutcome \$AuditOutcomeValue'
            $ScriptText | Should -Match '\$AuditOutcome -contains \$_\.AuditOutcome'
        }

        $GetAdminActivityScript | Should -Match 'AuditOutcome,'
    }
}
Describe "Get-WithRetry (via Scan-ADComputers)" {

    # Integration tests for retry behavior require mocking.
    # Placeholder: verify the function exists in the module after refactor.
    It "Get-WithRetry is available" {
        Get-Command -Name "Get-WithRetry" -ErrorAction SilentlyContinue |
            Should -Not -BeNullOrEmpty
    }
}
Describe "Default report output locations" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $GetAdminActivityScript = Get-Content -Raw (Join-Path $ScriptRoot "Get-ADAdminActivity.ps1")
        $ManageUsersScript = Get-Content -Raw (Join-Path $ScriptRoot "Manage-ADUserAccounts.ps1")
        $ScanComputersScript = Get-Content -Raw (Join-Path $ScriptRoot "Scan-ADComputers.ps1")
    }

    It "Get-ADAdminActivity defaults to reports/ad-admin-activity and logs/get-ad-admin-activity" {
        $GetAdminActivityScript | Should -Match 'reports\\ad-admin-activity'
        $GetAdminActivityScript | Should -Match 'logs\\get-ad-admin-activity'
        $GetAdminActivityScript | Should -Match '\$LogPath'
        $GetAdminActivityScript | Should -Not -Match 'LOCALAPPDATA'
        $GetAdminActivityScript | Should -Not -Match 'ADAdminActivityReports'
    }

    It "Manage-ADUserAccounts defaults to reports/ad-user-accounts and logs/manage-ad-user-accounts" {
        $ManageUsersScript | Should -Match 'reports\\ad-user-accounts'
        $ManageUsersScript | Should -Match 'logs\\manage-ad-user-accounts'
        $ManageUsersScript | Should -Match '\$LogPath'
        $ManageUsersScript | Should -Not -Match 'LOCALAPPDATA'
        $ManageUsersScript | Should -Not -Match 'ADUserAccountReports'
    }
    It "uses the shared output path resolver for custom log and output paths" {
        $CommonModuleScript = Get-Content -Raw (Join-Path $ScriptRoot "AdminToolsCommon.psm1")
        $CommonModuleScript | Should -Match 'function Get-ResolvedOutputPath'
        $GetAdminActivityScript | Should -Not -Match 'function Get-ResolvedOutputPath'
        $GetAdminActivityScript | Should -Match 'Get-ResolvedOutputPath -Path \$LogPath'
        $GetAdminActivityScript | Should -Match 'Get-ResolvedOutputPath -Path \$OutputCsv'
        $ManageUsersScript | Should -Not -Match 'function Get-ResolvedOutputPath'
        $ManageUsersScript | Should -Match 'Get-ResolvedOutputPath -Path \$LogPath'
    }

    It "Scan-ADComputers defaults to reports/ad-computers" {
        $ScanComputersScript | Should -Match 'reports\\ad-computers'
        $ScanComputersScript | Should -Match 'logs\\scan-ad-computers'
        $ScanComputersScript | Should -Not -Match '\$OutputDirectory = \$ScriptDirectory'
    }
}
Describe "Scan-ADComputers performance pipeline" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $ScanComputersScript = Get-Content -Raw (Join-Path $ScriptRoot "Scan-ADComputers.ps1")
    }

    It "exposes performance tuning parameters" {
        foreach ($parameterName in @(
            "PerformanceSummary",
            "AdResultPageSize",
            "AdSearchScope",
            "TargetedQueryChunkSize",
            "ConnectivityThrottleLimit",
            "DnsThrottleLimit",
            "PortThrottleLimit",
            "RemoteInventoryThrottleLimit"
        )) {
            $ScanComputersScript | Should -Match "\`$$parameterName"
        }
    }

    It "records pipeline stage timings" {
        $ScanComputersScript | Should -Match 'function Measure-PerformanceStage'
        $ScanComputersScript | Should -Match 'function Complete-PerformanceStage'
        $ScanComputersScript | Should -Match 'function Export-PerformanceSummary'
        foreach ($stageName in @(
            "ConfigAndValidation",
            "DomainDiscovery",
            "AdDiscovery",
            "TargetedConnectivity",
            "TargetedMatching",
            "InventoryPreparation",
            "OperationalConnectivity",
            "DnsResolution",
            "PortChecks",
            "RemoteInventory",
            "Export"
        )) {
            $ScanComputersScript | Should -Match $stageName
        }
    }

    It "passes AD page size and search scope into Get-ADComputer" {
        $ScanComputersScript | Should -Match 'ResultPageSize\s*=\s*\$AdResultPageSize'
        $ScanComputersScript | Should -Match 'SearchScope\s*=\s*\$AdSearchScope'
        $ScanComputersScript | Should -Match 'TargetedQueryChunkSize'
        $ScanComputersScript | Should -Match 'ProgressUpdateInterval'
    }
    It "accepts safe legacy short computer names with underscores in targeted lists" {
        $ScanComputersScript | Should -Match 'function Test-IsSafeShortComputerName'
        $ScanComputersScript | Should -Match '\[A-Za-z0-9_-'
        $ScanComputersScript | Should -Match 'Test-IsSafeShortComputerName -Name \$inputName'
    }
    It "reports ComputerListPath validation failures with line numbers" {
        $ScanComputersScript | Should -Match '\$lineNumber\+\+'
        $ScanComputersScript | Should -Match 'ComputerListPath line \$lineNumber contains an invalid short computer name'
        $ScanComputersScript | Should -Match 'ComputerListPath line \$lineNumber contains an invalid DNS name'
    }

    It "uses fixed timestamps when preparing inventory records" {
        $ScanComputersScript | Should -Match '\[datetime\]\$ScanTimestamp'
        $ScanComputersScript | Should -Match '\[datetime\]\$StaleComparisonTimestamp'
        $ScanComputersScript | Should -Match 'Get-StaleStatus -LastSeenDate \$lastSeenDate -ThresholdDays \$InactiveThresholdDays -ComparisonTimestamp \$StaleComparisonTimestamp'
        $ScanComputersScript | Should -Match 'ProgressUpdateInterval'
    }

    It "splits operational enrichment into separately throttled phases" {
        foreach ($functionName in @(
            "Invoke-ConnectivityEnrichment",
            "Invoke-DnsEnrichment",
            "Invoke-PortEnrichment",
            "Invoke-RemoteInventoryEnrichment"
        )) {
            $ScanComputersScript | Should -Match "function $functionName"
        }

        $ScanComputersScript | Should -Match 'Remote inventory requires connectivity gating'
        $ScanComputersScript | Should -Match '\$ConnectivityThrottleLimitEffective'
        $ScanComputersScript | Should -Match '\$DnsThrottleLimitEffective'
        $ScanComputersScript | Should -Match '\$PortThrottleLimitEffective'
        $ScanComputersScript | Should -Match '\$RemoteInventoryThrottleLimitEffective'
    }

    It "allows empty targeted result sets through enrichment and summary generation" {
        $ScanComputersScript | Should -Match 'function Invoke-OperationalEnrichment[\s\S]*?\[AllowEmptyCollection\(\)\]\s*\[object\[\]\]\$Records'
        $ScanComputersScript | Should -Match 'function Get-SummaryRecords[\s\S]*?\[AllowEmptyCollection\(\)\]\s*\[object\[\]\]\$Records'
        $ScanComputersScript | Should -Match 'if \(@\(\$Records\)\.Count -eq 0\) \{ return @\(\) \}'
        $ScanComputersScript | Should -Match '\$resultList = @\(Invoke-OperationalEnrichment -Records @\(\$resultList\) -PerformConnectivityCheck \$performConnectivityEnrichment\)'
    }
}


Describe "Stored credential loading" {
    It "exposes stored credential parameters on AD scripts" {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        foreach ($scriptName in @("Scan-ADComputers.ps1", "Manage-ADUserAccounts.ps1", "Get-ADAdminActivity.ps1")) {
            $scriptText = Get-Content -Raw (Join-Path $ScriptRoot $scriptName)
            $scriptText | Should -Match '\$CredentialSecretName'
            $scriptText | Should -Match '\$CredentialPath'
            $scriptText | Should -Match 'Resolve-AdminToolsCredential'
        }
    }

    It "rejects using multiple credential sources" {
        $credential = Get-TestCredential -UserName "DOMAIN\User"

        { Resolve-AdminToolsCredential -Credential $credential -CredentialSecretName "ADScanCredential" -CredentialPath $null -BaseDirectory $TestDrive } |
            Should -Throw '*Only one credential source*'
    }

    It "rejects credential files inside the repository" {
        $repoCredentialPath = Join-Path $ScriptRoot "credential.xml"

        { Resolve-AdminToolsCredential -CredentialPath $repoCredentialPath -BaseDirectory $ScriptRoot } |
            Should -Throw '*outside the repository*'
    }

    It "imports a PSCredential from an allowed CLIXML path" {
        $credential = Get-TestCredential -UserName "DOMAIN\User"
        $credentialPath = Join-Path $TestDrive "credential.xml"
        $credential | Export-Clixml -LiteralPath $credentialPath

        $result = Resolve-AdminToolsCredential -CredentialPath $credentialPath -BaseDirectory $ScriptRoot

        $result.UserName | Should -Be "DOMAIN\User"
    }

    It "uses SecretManagement when CredentialSecretName is supplied" {
        Mock -ModuleName AdminToolsCommon Get-Module { [PSCustomObject]@{ Name = "Microsoft.PowerShell.SecretManagement" } } -ParameterFilter { $ListAvailable -and $Name -eq "Microsoft.PowerShell.SecretManagement" }
        Mock -ModuleName AdminToolsCommon Import-Module { }
        Mock -ModuleName AdminToolsCommon Get-AdminToolsSecret {
            Get-TestCredential -UserName "DOMAIN\SecretUser"
        }

        $result = Resolve-AdminToolsCredential -CredentialSecretName "ADScanCredential" -BaseDirectory $ScriptRoot

        $result.UserName | Should -Be "DOMAIN\SecretUser"
        Should -Invoke -ModuleName AdminToolsCommon Get-AdminToolsSecret -Times 1 -ParameterFilter { $Name -eq "ADScanCredential" }
    }
}
Describe "Scan-ADComputers domain trust boundaries" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $ScanComputersScript = Get-Content -Raw (Join-Path $ScriptRoot "Scan-ADComputers.ps1")
    }

    It "only uses an operator-supplied DomainController for bootstrap after DomainName suffix validation" {
        $ScanComputersScript | Should -Match '\$domainDiscoveryParameters\s*=\s*@\{\s*Credential\s*=\s*\$Credential\s*\}'
        $ScanComputersScript | Should -Match '\$DomainControllerBootstrapServer\s*=\s*\$null'
        $ScanComputersScript | Should -Match 'Test-IsInDnsSuffix -Name \$DomainController -DnsSuffix \$DomainName'
        $ScanComputersScript | Should -Match '\$domainDiscoveryParameters\["Server"\]\s*=\s*\$DomainController'
        $ScanComputersScript | Should -Match '\$DomainControllerBootstrapServer\s*=\s*\$DomainController'
        $ScanComputersScript | Should -Match 'Get-VerifiedDomainControllerName'
    }
    It "uses explicit DomainName for VPN bootstrap identity and verifies the selected DC after discovery" {
        $ScanComputersScript | Should -Match '\$domainDiscoveryParameters\["Identity"\]\s*=\s*\$DomainName'
        $ScanComputersScript | Should -Match '\$domainDiscoveryParameters\["Server"\]\s*=\s*\$DomainName'
        $ScanComputersScript | Should -Match '\$domainControllerVerificationServer\s*=\s*if \(\$DomainControllerBootstrapServer\) \{ \$DomainControllerBootstrapServer \} else \{ \$adDomain\.DNSRoot \}'
        $ScanComputersScript | Should -Match 'Get-VerifiedDomainControllerName -DomainController \$DomainController -Credential \$Credential -DiscoveryServer \$domainControllerVerificationServer'
    }

    It "rejects DomainName values that do not match the discovered AD DNS root" {
        $ScanComputersScript | Should -Match '\$DomainName\.Equals\(\$adDomain\.DNSRoot,\s*\[System\.StringComparison\]::OrdinalIgnoreCase\)'
        $ScanComputersScript | Should -Match 'does not match discovered AD DNS root'
    }

    It "uses a separate credential path for remote inventory" {
        foreach ($parameterName in @(
            "RemoteInventoryCredential",
            "RemoteInventoryCredentialSecretName",
            "RemoteInventoryCredentialPath"
        )) {
            $ScanComputersScript | Should -Match "\`$$parameterName"
        }

        $ScanComputersScript | Should -Match '\$using:RemoteInventoryCredential'
        $ScanComputersScript | Should -Not -Match 'New-CimSession -ComputerName \$targetName -Credential \$using:Credential'
    }

    It "requires AD-backed target identity and SPN evidence before CIM inventory" {
        $ScanComputersScript | Should -Match 'servicePrincipalName'
        $ScanComputersScript | Should -Match 'Test-ComputerHasExpectedSpn'
        $ScanComputersScript | Should -Match 'SkippedUntrustedTarget: expected SPN missing'
    }
}
Describe "Security event query safety limits" {
    BeforeAll {
        $ScriptRoot = Split-Path -Parent $PSScriptRoot
        $CommonModuleScript = Get-Content -Raw (Join-Path $ScriptRoot "AdminToolsCommon.psm1")
        $AdminActivityScript = Get-Content -Raw (Join-Path $ScriptRoot "Get-ADAdminActivity.ps1")
        $ManageUsersScript = Get-Content -Raw (Join-Path $ScriptRoot "Manage-ADUserAccounts.ps1")
    }

    It "sets a conservative default event cap and requires UnlimitedEvents for uncapped queries" {
        foreach ($scriptText in @($AdminActivityScript, $ManageUsersScript)) {
            $scriptText | Should -Match '\[ValidateRange\(1, 1000000\)\]\s*\[int\]\$MaxEventsPerDomainController = 100000'
            $scriptText | Should -Match '\[switch\]\$UnlimitedEvents'
            $scriptText | Should -Match '\$EffectiveMaxEventsPerDomainController = if \(\$UnlimitedEvents\.IsPresent\) \{ 0 \} else \{ \$MaxEventsPerDomainController \}'
        }
    }

    It "processes and disposes EventRecord instances in the shared reader instead of returning raw records" {
        $CommonModuleScript | Should -Match '\[scriptblock\]\$ProcessEvent'
        $CommonModuleScript | Should -Match 'finally \{\s*if \(\$null -ne \$EventRecord\) \{\s*\$EventRecord\.Dispose\(\)'
        $CommonModuleScript | Should -Not -Match '\$Events = New-Object System\.Collections\.Generic\.List\[object\]'
        $CommonModuleScript | Should -Not -Match '\[void\]\$Events\.Add\(\$EventRecord\)'
    }

    It "adds truncation metadata to event output records" {
        foreach ($metadataName in @(
            'EventQueryMaxEventsPerDomainController',
            'EventQueryLimitReached',
            'EventQueryEventsReadFromDomainController',
            'EventQueryUnlimitedEvents'
        )) {
            $CommonModuleScript | Should -Match $metadataName
            $AdminActivityScript | Should -Match $metadataName
            $ManageUsersScript | Should -Match $metadataName
        }
    }

    It "warns for long lookback windows without an explicit event cap override" {
        foreach ($scriptText in @($AdminActivityScript, $ManageUsersScript)) {
            $scriptText | Should -Match '\$DaysBack -gt 90'
            $scriptText | Should -Match '\$PSBoundParameters\.ContainsKey\("MaxEventsPerDomainController"\)'
            $scriptText | Should -Match 'longer than 90 days'
        }
    }
}

