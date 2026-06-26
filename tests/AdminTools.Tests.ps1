#requires -Version 5.1
#requires -Modules Pester

BeforeAll {
    $ScriptRoot = Split-Path -Parent $PSScriptRoot
    Import-Module (Join-Path $ScriptRoot "AdminToolsCommon.psm1") -Force
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

    It "Get-ADAdminActivity defaults to reports/ad-admin-activity" {
        $GetAdminActivityScript | Should -Match 'reports\\ad-admin-activity'
        $GetAdminActivityScript | Should -Not -Match 'LOCALAPPDATA'
        $GetAdminActivityScript | Should -Not -Match 'ADAdminActivityReports'
    }

    It "Manage-ADUserAccounts defaults to reports/ad-user-accounts" {
        $ManageUsersScript | Should -Match 'reports\\ad-user-accounts'
        $ManageUsersScript | Should -Not -Match 'LOCALAPPDATA'
        $ManageUsersScript | Should -Not -Match 'ADUserAccountReports'
    }

    It "Scan-ADComputers defaults to reports/ad-computers" {
        $ScanComputersScript | Should -Match 'reports\\ad-computers'
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
        $ScanComputersScript | Should -Match 'function Start-PerformanceStage'
        $ScanComputersScript | Should -Match 'function Stop-PerformanceStage'
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
        $securePassword = ConvertTo-SecureString "pass" -AsPlainText -Force
        $credential = [PSCredential]::new("DOMAIN\User", $securePassword)

        { Resolve-AdminToolsCredential -Credential $credential -CredentialSecretName "ADScanCredential" -CredentialPath $null -BaseDirectory $TestDrive } |
            Should -Throw '*Only one credential source*'
    }

    It "rejects credential files inside the repository" {
        $repoCredentialPath = Join-Path $ScriptRoot "credential.xml"

        { Resolve-AdminToolsCredential -CredentialPath $repoCredentialPath -BaseDirectory $ScriptRoot } |
            Should -Throw '*outside the repository*'
    }

    It "imports a PSCredential from an allowed CLIXML path" {
        $securePassword = ConvertTo-SecureString "pass" -AsPlainText -Force
        $credential = [PSCredential]::new("DOMAIN\User", $securePassword)
        $credentialPath = Join-Path $TestDrive "credential.xml"
        $credential | Export-Clixml -LiteralPath $credentialPath

        $result = Resolve-AdminToolsCredential -CredentialPath $credentialPath -BaseDirectory $ScriptRoot

        $result.UserName | Should -Be "DOMAIN\User"
    }

    It "uses SecretManagement when CredentialSecretName is supplied" {
        Mock -ModuleName AdminToolsCommon Get-Module { [PSCustomObject]@{ Name = "Microsoft.PowerShell.SecretManagement" } } -ParameterFilter { $ListAvailable -and $Name -eq "Microsoft.PowerShell.SecretManagement" }
        Mock -ModuleName AdminToolsCommon Import-Module { }
        Mock -ModuleName AdminToolsCommon Get-AdminToolsSecret {
            $securePassword = ConvertTo-SecureString "pass" -AsPlainText -Force
            [PSCredential]::new("DOMAIN\SecretUser", $securePassword)
        }

        $result = Resolve-AdminToolsCredential -CredentialSecretName "ADScanCredential" -BaseDirectory $ScriptRoot

        $result.UserName | Should -Be "DOMAIN\SecretUser"
        Should -Invoke -ModuleName AdminToolsCommon Get-AdminToolsSecret -Times 1 -ParameterFilter { $Name -eq "ADScanCredential" }
    }
}

