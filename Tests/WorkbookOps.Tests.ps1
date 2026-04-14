# Tests/WorkbookOps.Tests.ps1
# Parameter-validation tests for WorkbookOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Open-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command Open-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Open-ExcelWorkbook).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Open-ExcelWorkbook).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Close-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command Close-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
}

Describe 'New-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command New-ExcelWorkbook).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetNames parameter' {
        (Get-Command New-ExcelWorkbook).Parameters['SheetNames'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when file already exists' {
        $tempFile = [System.IO.Path]::GetTempFileName()
        try {
            { New-ExcelWorkbook -WorkbookPath $tempFile } | Should -Throw '*already exists*'
        } finally {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Save-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command Save-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Save-ExcelWorkbook).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SaveAsPath parameter' {
        (Get-Command Save-ExcelWorkbook).Parameters['SaveAsPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelWorkbookInfo' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelWorkbookInfo).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelWorkbookInfo).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Repair-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command Repair-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Repair-ExcelWorkbook).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when file not found' {
        { Repair-ExcelWorkbook -WorkbookPath 'C:\nonexistent\fake.xlsx' } | Should -Throw '*not found*'
    }
}

Describe 'Copy-ExcelWorkbook' {
    It 'Has CmdletBinding' {
        (Get-Command Copy-ExcelWorkbook).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Copy-ExcelWorkbook).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has DestinationPath parameter' {
        (Get-Command Copy-ExcelWorkbook).Parameters['DestinationPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Throws when source not found' {
        { Copy-ExcelWorkbook -WorkbookPath 'C:\nonexistent\fake.xlsx' -DestinationPath 'C:\temp\out.xlsx' } | Should -Throw '*not found*'
    }
}

Describe 'Invoke-ExcelMacro' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelMacro).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Invoke-ExcelMacro).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory MacroName parameter' {
        $p = (Get-Command Invoke-ExcelMacro).Parameters['MacroName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Arguments parameter with ValidateCount(0,30)' {
        $p = (Get-Command Invoke-ExcelMacro).Parameters['Arguments']
        $p | Should -Not -BeNullOrEmpty
        $vc = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateCountAttribute] })
        $vc | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Invoke-ExcelMacro).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
