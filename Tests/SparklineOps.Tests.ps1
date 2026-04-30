# Tests/SparklineOps.Tests.ps1
# Parameter-validation tests for SparklineOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Add-ExcelSparkline' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelSparkline).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LocationRange parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['LocationRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DataRange parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['DataRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SparklineType parameter with validation' {
        $param = (Get-Command Add-ExcelSparkline).Parameters['SparklineType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowHighPoint parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowHighPoint'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowLowPoint parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowLowPoint'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowFirstPoint parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowFirstPoint'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowLastPoint parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowLastPoint'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowNegativePoints parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowNegativePoints'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ShowMarkers parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['ShowMarkers'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Add-ExcelSparkline).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelSparkline' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelSparkline).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelSparkline).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelSparkline).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has CellAddress parameter' {
        (Get-Command Get-ExcelSparkline).Parameters['CellAddress'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelSparkline).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelSparkline' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelSparkline).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Remove-ExcelSparkline).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelSparkline).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has CellAddress parameter' {
        (Get-Command Remove-ExcelSparkline).Parameters['CellAddress'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Remove-ExcelSparkline).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}
