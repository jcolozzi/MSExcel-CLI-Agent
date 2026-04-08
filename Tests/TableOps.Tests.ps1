# Tests/TableOps.Tests.ps1
# Parameter-validation tests for TableOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-ExcelTable' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelTable).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelTable).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter (optional)' {
        (Get-Command Get-ExcelTable).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'New-ExcelTable' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command New-ExcelTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command New-ExcelTable).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has HasHeaders parameter' {
        (Get-Command New-ExcelTable).Parameters['HasHeaders'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelTable' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Remove-ExcelTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has KeepData switch' {
        (Get-Command Remove-ExcelTable).Parameters['KeepData'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Resize-ExcelTable' {
    It 'Has CmdletBinding' {
        (Get-Command Resize-ExcelTable).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Resize-ExcelTable).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewRange parameter' {
        (Get-Command Resize-ExcelTable).Parameters['NewRange'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelTableData' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelTableData).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Get-ExcelTableData).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Limit parameter' {
        (Get-Command Get-ExcelTableData).Parameters['Limit'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Add-ExcelTableRow' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelTableRow).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Add-ExcelTableRow).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Rows parameter' {
        (Get-Command Add-ExcelTableRow).Parameters['Rows'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelTableRow' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelTableRow).CmdletBinding | Should -BeTrue
    }
    It 'Has RowIndex parameter' {
        (Get-Command Remove-ExcelTableRow).Parameters['RowIndex'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelTableColumn' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelTableColumn).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Get-ExcelTableColumn).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ColumnName parameter (optional)' {
        (Get-Command Get-ExcelTableColumn).Parameters['ColumnName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Limit parameter' {
        (Get-Command Get-ExcelTableColumn).Parameters['Limit'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelTableTotals' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelTableTotals).CmdletBinding | Should -BeTrue
    }
    It 'Has TableName parameter' {
        (Get-Command Set-ExcelTableTotals).Parameters['TableName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowTotals parameter' {
        (Get-Command Set-ExcelTableTotals).Parameters['ShowTotals'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Calculations parameter' {
        (Get-Command Set-ExcelTableTotals).Parameters['Calculations'] | Should -Not -BeNullOrEmpty
    }
}
