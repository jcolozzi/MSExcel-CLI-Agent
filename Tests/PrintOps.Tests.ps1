# Tests/PrintOps.Tests.ps1
# Parameter-validation tests for PrintOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelPageSetup' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelPageSetup).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Orientation parameter with validation' {
        $param = (Get-Command Set-ExcelPageSetup).Parameters['Orientation']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has PaperSize parameter with validation' {
        $param = (Get-Command Set-ExcelPageSetup).Parameters['PaperSize']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has LeftMargin parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['LeftMargin'] | Should -Not -BeNullOrEmpty
    }
    It 'Has RightMargin parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['RightMargin'] | Should -Not -BeNullOrEmpty
    }
    It 'Has TopMargin parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['TopMargin'] | Should -Not -BeNullOrEmpty
    }
    It 'Has BottomMargin parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['BottomMargin'] | Should -Not -BeNullOrEmpty
    }
    It 'Has HeaderLeft parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['HeaderLeft'] | Should -Not -BeNullOrEmpty
    }
    It 'Has HeaderCenter parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['HeaderCenter'] | Should -Not -BeNullOrEmpty
    }
    It 'Has HeaderRight parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['HeaderRight'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FooterLeft parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['FooterLeft'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FooterCenter parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['FooterCenter'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FooterRight parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['FooterRight'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FitToPagesWide parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['FitToPagesWide'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FitToPagesTall parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['FitToPagesTall'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PrintArea parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['PrintArea'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelPageSetup).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelPageSetup' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelPageSetup).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelPageSetup).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelPageSetup).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelPageSetup).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-ExcelToPdf' {
    It 'Has CmdletBinding' {
        (Get-Command Export-ExcelToPdf).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Export-ExcelToPdf).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-ExcelToPdf).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Export-ExcelToPdf).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Quality parameter with validation' {
        $param = (Get-Command Export-ExcelToPdf).Parameters['Quality']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Export-ExcelToPdf).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
