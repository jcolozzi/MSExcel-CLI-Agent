# Tests/SlicerOps.Tests.ps1
# Pester 5+ tests — SlicerOps parameter validation

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
# New-ExcelSlicer
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'New-ExcelSlicer' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelSlicer).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command New-ExcelSlicer).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command New-ExcelSlicer).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SourceName parameter' {
        $p = (Get-Command New-ExcelSlicer).Parameters['SourceName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SourceField parameter' {
        $p = (Get-Command New-ExcelSlicer).Parameters['SourceField']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command New-ExcelSlicer).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Get-ExcelSlicer
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Get-ExcelSlicer' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelSlicer).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Get-ExcelSlicer).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Get-ExcelSlicer).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Set-ExcelSlicer
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Set-ExcelSlicer' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelSlicer).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Set-ExcelSlicer).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Set-ExcelSlicer).Parameters['Name']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Caption parameter' {
        (Get-Command Set-ExcelSlicer).Parameters['Caption'] | Should -Not -BeNullOrEmpty
    }
    It 'Has optional NumberOfColumns parameter' {
        (Get-Command Set-ExcelSlicer).Parameters['NumberOfColumns'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelSlicer).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelSlicer
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelSlicer' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelSlicer).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelSlicer).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Remove-ExcelSlicer).Parameters['Name']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelSlicer).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# New-ExcelTimeline
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'New-ExcelTimeline' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelTimeline).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command New-ExcelTimeline).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command New-ExcelTimeline).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SourceName parameter' {
        $p = (Get-Command New-ExcelTimeline).Parameters['SourceName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory DateField parameter' {
        $p = (Get-Command New-ExcelTimeline).Parameters['DateField']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command New-ExcelTimeline).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Set-ExcelTimelineRange
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Set-ExcelTimelineRange' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelTimelineRange).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Set-ExcelTimelineRange).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Name parameter' {
        $p = (Get-Command Set-ExcelTimelineRange).Parameters['Name']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory StartDate parameter' {
        $p = (Get-Command Set-ExcelTimelineRange).Parameters['StartDate']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory EndDate parameter' {
        $p = (Get-Command Set-ExcelTimelineRange).Parameters['EndDate']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelTimelineRange).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
