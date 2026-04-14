# Tests/StructuralOps.Tests.ps1
# Pester 5+ tests — StructuralOps parameter validation

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ═══════════════════════════════════════════════════════════════════════════════
# Add-ExcelRow
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Add-ExcelRow' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelRow).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Add-ExcelRow).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Add-ExcelRow).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Row parameter with ValidateRange' {
        $p = (Get-Command Add-ExcelRow).Parameters['Row']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has Count parameter with ValidateRange' {
        $p = (Get-Command Add-ExcelRow).Parameters['Count']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelRow).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelRow
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelRow' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelRow).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelRow).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Remove-ExcelRow).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Row parameter with ValidateRange' {
        $p = (Get-Command Remove-ExcelRow).Parameters['Row']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has Count parameter with ValidateRange' {
        $p = (Get-Command Remove-ExcelRow).Parameters['Count']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelRow).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Add-ExcelColumn
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Add-ExcelColumn' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelColumn).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Add-ExcelColumn).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Add-ExcelColumn).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Column parameter' {
        $p = (Get-Command Add-ExcelColumn).Parameters['Column']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Count parameter with ValidateRange' {
        $p = (Get-Command Add-ExcelColumn).Parameters['Count']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelColumn).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelColumn
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelColumn' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelColumn).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelColumn).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Remove-ExcelColumn).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Column parameter' {
        $p = (Get-Command Remove-ExcelColumn).Parameters['Column']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Count parameter with ValidateRange' {
        $p = (Get-Command Remove-ExcelColumn).Parameters['Count']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelColumn).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
