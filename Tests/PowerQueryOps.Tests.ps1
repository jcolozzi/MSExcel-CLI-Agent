# Tests/PowerQueryOps.Tests.ps1
# Pester 5+ tests — PowerQueryOps parameter validation

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
# Get-ExcelPowerQuery
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Get-ExcelPowerQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelPowerQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Get-ExcelPowerQuery).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Name parameter' {
        (Get-Command Get-ExcelPowerQuery).Parameters['Name'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Get-ExcelPowerQuery).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# New-ExcelPowerQuery
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'New-ExcelPowerQuery' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelPowerQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command New-ExcelPowerQuery).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command New-ExcelPowerQuery).Parameters['Name']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Formula parameter' {
        $p = (Get-Command New-ExcelPowerQuery).Parameters['Formula']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Description parameter' {
        (Get-Command New-ExcelPowerQuery).Parameters['Description'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command New-ExcelPowerQuery).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Set-ExcelPowerQuery
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Set-ExcelPowerQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelPowerQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Set-ExcelPowerQuery).Parameters['Name']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Formula parameter' {
        (Get-Command Set-ExcelPowerQuery).Parameters['Formula'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelPowerQuery).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelPowerQuery
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelPowerQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelPowerQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Remove-ExcelPowerQuery).Parameters['Name']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelPowerQuery).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Update-ExcelPowerQuery
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Update-ExcelPowerQuery' {
    It 'Has CmdletBinding' {
        (Get-Command Update-ExcelPowerQuery).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Update-ExcelPowerQuery).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has All switch' {
        (Get-Command Update-ExcelPowerQuery).Parameters['All'].SwitchParameter | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Update-ExcelPowerQuery).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Import-ExcelPowerQueryToTable
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Import-ExcelPowerQueryToTable' {
    It 'Has CmdletBinding' {
        (Get-Command Import-ExcelPowerQueryToTable).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Import-ExcelPowerQueryToTable).Parameters['SheetName']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory QueryName parameter' {
        $p = (Get-Command Import-ExcelPowerQueryToTable).Parameters['QueryName']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Destination parameter' {
        (Get-Command Import-ExcelPowerQueryToTable).Parameters['Destination'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Import-ExcelPowerQueryToTable).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
