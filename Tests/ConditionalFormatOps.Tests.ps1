# Tests/ConditionalFormatOps.Tests.ps1
# Parameter-validation tests for ConditionalFormatOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Add-ExcelConditionalFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelConditionalFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has RuleType parameter with validation' {
        $param = (Get-Command Add-ExcelConditionalFormat).Parameters['RuleType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Operator parameter with validation' {
        $param = (Get-Command Add-ExcelConditionalFormat).Parameters['Operator']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Value1 parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['Value1'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Value2 parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['Value2'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Formula parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['Formula'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FontColor parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['FontColor'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FillColor parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['FillColor'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Add-ExcelConditionalFormat).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelConditionalFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelConditionalFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelConditionalFormat).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelConditionalFormat).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Get-ExcelConditionalFormat).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelConditionalFormat).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelConditionalFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelConditionalFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has RuleIndex parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['RuleIndex'] | Should -Not -BeNullOrEmpty
    }
    It 'Has All parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['All'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelConditionalFormat).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Clear-ExcelConditionalFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Clear-ExcelConditionalFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Clear-ExcelConditionalFormat).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Clear-ExcelConditionalFormat).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Clear-ExcelConditionalFormat).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
