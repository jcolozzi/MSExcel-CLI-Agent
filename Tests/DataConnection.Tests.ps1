# Tests/DataConnection.Tests.ps1
# Pester 5+ tests — DataConnection parameter validation

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
# Update-ExcelDataConnection
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Update-ExcelDataConnection' {
    It 'Has CmdletBinding' {
        (Get-Command Update-ExcelDataConnection).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Update-ExcelDataConnection).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Name parameter' {
        (Get-Command Update-ExcelDataConnection).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has All switch' {
        (Get-Command Update-ExcelDataConnection).Parameters['All'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Update-ExcelDataConnection).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelDataConnection
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelDataConnection' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelDataConnection).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelDataConnection).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Remove-ExcelDataConnection).Parameters['Name']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelDataConnection).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# New-ExcelDataConnection
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'New-ExcelDataConnection' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelDataConnection).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command New-ExcelDataConnection).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command New-ExcelDataConnection).Parameters['Name']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory ConnectionString parameter' {
        $p = (Get-Command New-ExcelDataConnection).Parameters['ConnectionString']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory CommandText parameter' {
        $p = (Get-Command New-ExcelDataConnection).Parameters['CommandText']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has CommandType parameter' {
        (Get-Command New-ExcelDataConnection).Parameters['CommandType'] | Should -Not -BeNullOrEmpty
    }
    It 'Has optional Description parameter' {
        (Get-Command New-ExcelDataConnection).Parameters['Description'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command New-ExcelDataConnection).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
