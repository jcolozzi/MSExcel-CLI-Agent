# Tests/CalculationOps.Tests.ps1
# Pester 5+ tests — CalculationOps parameter validation

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
# Set-ExcelPerformanceMode
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Set-ExcelPerformanceMode' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelPerformanceMode).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Set-ExcelPerformanceMode).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Calculation parameter with ValidateSet' {
        $p = (Get-Command Set-ExcelPerformanceMode).Parameters['Calculation']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Automatic'
        $vs.ValidValues | Should -Contain 'Manual'
        $vs.ValidValues | Should -Contain 'SemiAutomatic'
    }
    It 'Does not require SheetName' {
        (Get-Command Set-ExcelPerformanceMode).Parameters.Keys | Should -Not -Contain 'SheetName'
    }
    It 'Has ScreenUpdating parameter' {
        (Get-Command Set-ExcelPerformanceMode).Parameters['ScreenUpdating'] | Should -Not -BeNullOrEmpty
    }
    It 'Has EnableEvents parameter' {
        (Get-Command Set-ExcelPerformanceMode).Parameters['EnableEvents'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelPerformanceMode).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Invoke-ExcelCalculate
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Invoke-ExcelCalculate' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelCalculate).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Invoke-ExcelCalculate).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Full switch' {
        (Get-Command Invoke-ExcelCalculate).Parameters['Full'].SwitchParameter | Should -BeTrue
    }
    It 'SheetName is optional' {
        $p = (Get-Command Invoke-ExcelCalculate).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $mandatory = $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory })
        $mandatory | Should -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Invoke-ExcelCalculate).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Invoke-ExcelFunction
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Invoke-ExcelFunction' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelFunction).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory FunctionName parameter' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['FunctionName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Arguments parameter (object array)' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['Arguments']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'SheetName is optional' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Invoke-ExcelFunction).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Invoke-ExcelEvaluate
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Invoke-ExcelEvaluate' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelEvaluate).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Invoke-ExcelEvaluate).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Expression parameter' {
        $p = (Get-Command Invoke-ExcelEvaluate).Parameters['Expression']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'SheetName is optional' {
        $p = (Get-Command Invoke-ExcelEvaluate).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Invoke-ExcelEvaluate).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
