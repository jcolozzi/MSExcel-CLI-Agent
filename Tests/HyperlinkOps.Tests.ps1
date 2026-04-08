# Tests/HyperlinkOps.Tests.ps1
# Parameter-validation tests for HyperlinkOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelHyperlink' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelHyperlink).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Address parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['Address'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SubAddress parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['SubAddress'] | Should -Not -BeNullOrEmpty
    }
    It 'Has TextToDisplay parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['TextToDisplay'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ScreenTip parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['ScreenTip'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelHyperlink' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelHyperlink).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelHyperlink).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelHyperlink).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Get-ExcelHyperlink).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelHyperlink' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelHyperlink).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelHyperlink).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelHyperlink).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Remove-ExcelHyperlink).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
