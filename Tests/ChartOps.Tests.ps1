# Tests/ChartOps.Tests.ps1
# Parameter-validation tests for ChartOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command New-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command New-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command New-ExcelChart).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartType parameter with validation' {
        $param = (Get-Command New-ExcelChart).Parameters['ChartType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command New-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Title parameter' {
        (Get-Command New-ExcelChart).Parameters['Title'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Left parameter' {
        (Get-Command New-ExcelChart).Parameters['Left'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Top parameter' {
        (Get-Command New-ExcelChart).Parameters['Top'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command New-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command New-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command New-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Get-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Set-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartType parameter with validation' {
        $param = (Get-Command Set-ExcelChart).Parameters['ChartType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Title parameter' {
        (Get-Command Set-ExcelChart).Parameters['Title'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Left parameter' {
        (Get-Command Set-ExcelChart).Parameters['Left'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Top parameter' {
        (Get-Command Set-ExcelChart).Parameters['Top'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Set-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Set-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command Set-ExcelChart).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Export-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Export-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Export-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Export-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-ExcelChart).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Export-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Export-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Export-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
