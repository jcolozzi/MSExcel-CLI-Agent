# Tests/ViewOps.Tests.ps1
# Parameter-validation tests for ViewOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Set-ExcelFreezePane' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelFreezePane).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Row parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['Row'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Column parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['Column'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Unfreeze parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['Unfreeze'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelFreezePane).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelFreezePane' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelFreezePane).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelFreezePane).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelFreezePane).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelFreezePane).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelSheetVisibility' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelSheetVisibility).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelSheetVisibility).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelSheetVisibility).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Visibility parameter with validation' {
        $param = (Get-Command Set-ExcelSheetVisibility).Parameters['Visibility']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelSheetVisibility).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelSheetVisibility' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelSheetVisibility).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelSheetVisibility).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelSheetVisibility).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelGrouping' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelGrouping).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelGrouping).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelGrouping).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelGrouping).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Ungroup parameter' {
        (Get-Command Set-ExcelGrouping).Parameters['Ungroup'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelGrouping).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelOutlineLevel' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelOutlineLevel).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelOutlineLevel).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelOutlineLevel).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has RowLevel parameter' {
        (Get-Command Set-ExcelOutlineLevel).Parameters['RowLevel'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ColumnLevel parameter' {
        (Get-Command Set-ExcelOutlineLevel).Parameters['ColumnLevel'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelOutlineLevel).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
