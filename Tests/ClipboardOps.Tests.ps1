# Tests/ClipboardOps.Tests.ps1
# Parameter-validation tests for ClipboardOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Copy-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Copy-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Copy-ExcelRange).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Copy-ExcelRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command Copy-ExcelRange).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationSheet parameter' {
        (Get-Command Copy-ExcelRange).Parameters['DestinationSheet'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationRange parameter' {
        (Get-Command Copy-ExcelRange).Parameters['DestinationRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PasteType parameter with validation' {
        $param = (Get-Command Copy-ExcelRange).Parameters['PasteType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Copy-ExcelRange).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Update-ExcelValue' {
    It 'Has CmdletBinding' {
        (Get-Command Update-ExcelValue).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Update-ExcelValue).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Update-ExcelValue).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SearchText parameter' {
        (Get-Command Update-ExcelValue).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ReplaceText parameter' {
        (Get-Command Update-ExcelValue).Parameters['ReplaceText'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Update-ExcelValue).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has MatchCase parameter' {
        (Get-Command Update-ExcelValue).Parameters['MatchCase'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Update-ExcelValue).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Move-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Move-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Move-ExcelRange).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Move-ExcelRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command Move-ExcelRange).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationSheet parameter' {
        (Get-Command Move-ExcelRange).Parameters['DestinationSheet'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationRange parameter' {
        (Get-Command Move-ExcelRange).Parameters['DestinationRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Move-ExcelRange).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
