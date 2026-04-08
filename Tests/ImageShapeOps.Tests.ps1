# Tests/ImageShapeOps.Tests.ps1
# Parameter-validation tests for ImageShapeOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Add-ExcelImage' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelImage).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Add-ExcelImage).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Add-ExcelImage).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ImagePath parameter' {
        (Get-Command Add-ExcelImage).Parameters['ImagePath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has CellAddress parameter' {
        (Get-Command Add-ExcelImage).Parameters['CellAddress'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Add-ExcelImage).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Add-ExcelImage).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LinkToFile parameter' {
        (Get-Command Add-ExcelImage).Parameters['LinkToFile'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Add-ExcelImage).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelShape' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelShape).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelShape).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelShape).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShapeName parameter' {
        (Get-Command Get-ExcelShape).Parameters['ShapeName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelShape' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelShape).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelShape).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelShape).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShapeName parameter' {
        (Get-Command Remove-ExcelShape).Parameters['ShapeName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
