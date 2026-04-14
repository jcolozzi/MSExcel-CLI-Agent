# Tests/WorksheetOps.Tests.ps1
# Parameter-validation tests for WorksheetOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelWorksheet).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'New-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command New-ExcelWorksheet).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has After parameter' {
        (Get-Command New-ExcelWorksheet).Parameters['After'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Remove-ExcelWorksheet).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Rename-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command Rename-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Rename-ExcelWorksheet).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewName parameter' {
        (Get-Command Rename-ExcelWorksheet).Parameters['NewName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Copy-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command Copy-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Copy-ExcelWorksheet).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has NewName parameter' {
        (Get-Command Copy-ExcelWorksheet).Parameters['NewName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Move-ExcelWorksheet' {
    It 'Has CmdletBinding' {
        (Get-Command Move-ExcelWorksheet).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Move-ExcelWorksheet).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Before parameter' {
        (Get-Command Move-ExcelWorksheet).Parameters['Before'] | Should -Not -BeNullOrEmpty
    }
    It 'Has After parameter' {
        (Get-Command Move-ExcelWorksheet).Parameters['After'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelRange).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command Get-ExcelRange).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has Value parameter' {
        (Get-Command Set-ExcelRange).Parameters['Value'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Formula parameter' {
        (Get-Command Set-ExcelRange).Parameters['Formula'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Clear-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Clear-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has ClearType parameter with validation' {
        $param = (Get-Command Clear-ExcelRange).Parameters['ClearType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelUsedRange' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelUsedRange).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelUsedRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Find-ExcelValue' {
    It 'Has CmdletBinding' {
        (Get-Command Find-ExcelValue).CmdletBinding | Should -BeTrue
    }
    It 'Has SearchText parameter' {
        (Get-Command Find-ExcelValue).Parameters['SearchText'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LookIn parameter with validation' {
        $param = (Get-Command Find-ExcelValue).Parameters['LookIn']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has MaxResults parameter' {
        (Get-Command Find-ExcelValue).Parameters['MaxResults'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelNamedRange' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelNamedRange).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter (optional)' {
        (Get-Command Get-ExcelNamedRange).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'New-ExcelNamedRange' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelNamedRange).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command New-ExcelNamedRange).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter' {
        (Get-Command New-ExcelNamedRange).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Remove-ExcelNamedRange' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelNamedRange).CmdletBinding | Should -BeTrue
    }
    It 'Has Name parameter' {
        (Get-Command Remove-ExcelNamedRange).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelSpecialCells' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelSpecialCells).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory CellType parameter with ValidateSet' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['CellType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Blanks'
        $vs.ValidValues | Should -Contain 'Formulas'
        $vs.ValidValues | Should -Contain 'LastCell'
        $vs.ValidValues | Should -Contain 'Constants'
        $vs.ValidValues | Should -Contain 'Visible'
        $vs.ValidValues | Should -Contain 'Comments'
        $vs.ValidValues | Should -Contain 'Errors'
    }
    It 'Range is optional' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['Range']
        $p | Should -Not -BeNullOrEmpty
        $mandatory = $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory })
        $mandatory | Should -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Get-ExcelSpecialCells).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
