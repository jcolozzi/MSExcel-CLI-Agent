# Tests/ImportOps.Tests.ps1
# Parameter-validation tests for ImportOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Import-ExcelCsv' {
    It 'Has CmdletBinding' {
        (Get-Command Import-ExcelCsv).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Import-ExcelCsv).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Import-ExcelCsv).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has CsvPath parameter' {
        (Get-Command Import-ExcelCsv).Parameters['CsvPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has StartCell parameter' {
        (Get-Command Import-ExcelCsv).Parameters['StartCell'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Delimiter parameter with validation' {
        $param = (Get-Command Import-ExcelCsv).Parameters['Delimiter']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has HasHeaders parameter' {
        (Get-Command Import-ExcelCsv).Parameters['HasHeaders'] | Should -Not -BeNullOrEmpty
    }
    It 'Has TextQualifier parameter with validation' {
        $param = (Get-Command Import-ExcelCsv).Parameters['TextQualifier']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Import-ExcelCsv).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Import-ExcelText' {
    It 'Has CmdletBinding' {
        (Get-Command Import-ExcelText).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Import-ExcelText).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Import-ExcelText).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has TextPath parameter' {
        (Get-Command Import-ExcelText).Parameters['TextPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has StartCell parameter' {
        (Get-Command Import-ExcelText).Parameters['StartCell'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ParseType parameter with validation' {
        $param = (Get-Command Import-ExcelText).Parameters['ParseType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Delimiter parameter' {
        (Get-Command Import-ExcelText).Parameters['Delimiter'] | Should -Not -BeNullOrEmpty
    }
    It 'Has TextQualifier parameter with validation' {
        $param = (Get-Command Import-ExcelText).Parameters['TextQualifier']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Import-ExcelText).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
