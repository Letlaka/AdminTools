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
