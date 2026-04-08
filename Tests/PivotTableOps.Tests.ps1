# Tests/PivotTableOps.Tests.ps1
# Parameter-validation tests for PivotTableOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-ExcelPivotTable' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelPivotTable).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PivotTableName parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['PivotTableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationSheet parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['DestinationSheet'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationCell parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['DestinationCell'] | Should -Not -BeNullOrEmpty
    }
    It 'Has RowFields parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['RowFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ColumnFields parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['ColumnFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DataFields parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['DataFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FilterFields parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['FilterFields'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command New-ExcelPivotTable).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelPivotTable' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelPivotTable).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelPivotTable).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelPivotTable).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PivotTableName parameter' {
        (Get-Command Get-ExcelPivotTable).Parameters['PivotTableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelPivotTable).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Update-ExcelPivotTable' {
    It 'Has CmdletBinding' {
        (Get-Command Update-ExcelPivotTable).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Update-ExcelPivotTable).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Update-ExcelPivotTable).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PivotTableName parameter' {
        (Get-Command Update-ExcelPivotTable).Parameters['PivotTableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Update-ExcelPivotTable).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
