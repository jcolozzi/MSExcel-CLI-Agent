# Tests/FormattingOps.Tests.ps1
# Parameter-validation tests for FormattingOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelCellFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelCellFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Bold parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['Bold'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Italic parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['Italic'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FontSize parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['FontSize'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FontName parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['FontName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FontColor parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['FontColor'] | Should -Not -BeNullOrEmpty
    }
    It 'Has FillColor parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['FillColor'] | Should -Not -BeNullOrEmpty
    }
    It 'Has BorderStyle parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['BorderStyle'] | Should -Not -BeNullOrEmpty
    }
    It 'Has BorderWeight parameter' {
        (Get-Command Set-ExcelCellFormat).Parameters['BorderWeight'] | Should -Not -BeNullOrEmpty
    }
    It 'Has BorderEdges parameter with validation' {
        $param = (Get-Command Set-ExcelCellFormat).Parameters['BorderEdges']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelNumberFormat' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelNumberFormat).CmdletBinding | Should -BeTrue
    }
    It 'Has NumberFormat parameter' {
        (Get-Command Set-ExcelNumberFormat).Parameters['NumberFormat'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelColumnWidth' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelColumnWidth).CmdletBinding | Should -BeTrue
    }
    It 'Has Column parameter' {
        (Get-Command Set-ExcelColumnWidth).Parameters['Column'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Set-ExcelColumnWidth).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelRowHeight' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelRowHeight).CmdletBinding | Should -BeTrue
    }
    It 'Has Row parameter' {
        (Get-Command Set-ExcelRowHeight).Parameters['Row'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Set-ExcelRowHeight).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelAlignment' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelAlignment).CmdletBinding | Should -BeTrue
    }
    It 'Has Horizontal parameter' {
        (Get-Command Set-ExcelAlignment).Parameters['Horizontal'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Vertical parameter' {
        (Get-Command Set-ExcelAlignment).Parameters['Vertical'] | Should -Not -BeNullOrEmpty
    }
    It 'Has WrapText parameter' {
        (Get-Command Set-ExcelAlignment).Parameters['WrapText'] | Should -Not -BeNullOrEmpty
    }
    It 'Has MergeCells parameter' {
        (Get-Command Set-ExcelAlignment).Parameters['MergeCells'] | Should -Not -BeNullOrEmpty
    }
}
