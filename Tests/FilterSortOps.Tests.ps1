# Tests/FilterSortOps.Tests.ps1
# Parameter-validation tests for FilterSortOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelAutoFilter' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelAutoFilter).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Column parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['Column'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Criteria1 parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['Criteria1'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Operator parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['Operator'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Criteria2 parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['Criteria2'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelAutoFilter).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelAutoFilter' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelAutoFilter).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelAutoFilter).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelAutoFilter).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelAutoFilter).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Sort-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Sort-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Sort-ExcelRange).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Sort-ExcelRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Sort-ExcelRange).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SortKey1 parameter' {
        (Get-Command Sort-ExcelRange).Parameters['SortKey1'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Order1 parameter' {
        (Get-Command Sort-ExcelRange).Parameters['Order1'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SortKey2 parameter' {
        (Get-Command Sort-ExcelRange).Parameters['SortKey2'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Order2 parameter' {
        (Get-Command Sort-ExcelRange).Parameters['Order2'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Header parameter' {
        (Get-Command Sort-ExcelRange).Parameters['Header'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Sort-ExcelRange).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelAutoFilter' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelAutoFilter).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelAutoFilter).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelAutoFilter).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelAutoFilter).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelDuplicates' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelDuplicates).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelDuplicates).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Range parameter' {
        $p = (Get-Command Remove-ExcelDuplicates).Parameters['Range']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Columns parameter (int array)' {
        $p = (Get-Command Remove-ExcelDuplicates).Parameters['Columns']
        $p | Should -Not -BeNullOrEmpty
        $p.ParameterType.Name | Should -Be 'Int32[]'
    }
    It 'Has HasHeader switch' {
        (Get-Command Remove-ExcelDuplicates).Parameters['HasHeader'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelDuplicates).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Invoke-ExcelAdvancedFilter' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelAdvancedFilter).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Invoke-ExcelAdvancedFilter).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Range parameter' {
        $p = (Get-Command Invoke-ExcelAdvancedFilter).Parameters['Range']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Action parameter with ValidateSet' {
        $p = (Get-Command Invoke-ExcelAdvancedFilter).Parameters['Action']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'FilterInPlace'
        $vs.ValidValues | Should -Contain 'CopyToRange'
    }
    It 'Has CriteriaRange parameter' {
        (Get-Command Invoke-ExcelAdvancedFilter).Parameters['CriteriaRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has UniqueOnly switch' {
        (Get-Command Invoke-ExcelAdvancedFilter).Parameters['UniqueOnly'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Invoke-ExcelAdvancedFilter).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
