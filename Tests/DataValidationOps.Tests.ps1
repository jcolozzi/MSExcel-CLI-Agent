# Tests/DataValidationOps.Tests.ps1
# Parameter-validation tests for DataValidationOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelDataValidation' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelDataValidation).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ValidationType parameter with validation' {
        $param = (Get-Command Set-ExcelDataValidation).Parameters['ValidationType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Operator parameter with validation' {
        $param = (Get-Command Set-ExcelDataValidation).Parameters['Operator']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Formula1 parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['Formula1'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Formula2 parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['Formula2'] | Should -Not -BeNullOrEmpty
    }
    It 'Has InputTitle parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['InputTitle'] | Should -Not -BeNullOrEmpty
    }
    It 'Has InputMessage parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['InputMessage'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ErrorTitle parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['ErrorTitle'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ErrorMessage parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['ErrorMessage'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ErrorStyle parameter with validation' {
        $param = (Get-Command Set-ExcelDataValidation).Parameters['ErrorStyle']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelDataValidation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelDataValidation' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelDataValidation).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelDataValidation).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelDataValidation).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Get-ExcelDataValidation).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelDataValidation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelDataValidation' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelDataValidation).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelDataValidation).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelDataValidation).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Remove-ExcelDataValidation).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelDataValidation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
