# Tests/DataModelOps.Tests.ps1
# Pester 5+ tests — DataModelOps parameter validation

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
# Get-ExcelDataModel
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Get-ExcelDataModel' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelDataModel).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Get-ExcelDataModel).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Get-ExcelDataModel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Add-ExcelModelMeasure
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Add-ExcelModelMeasure' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelModelMeasure).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Add-ExcelModelMeasure).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory MeasureName parameter' {
        $p = (Get-Command Add-ExcelModelMeasure).Parameters['MeasureName']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory TableName parameter' {
        $p = (Get-Command Add-ExcelModelMeasure).Parameters['TableName']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory Formula parameter' {
        $p = (Get-Command Add-ExcelModelMeasure).Parameters['Formula']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has FormatType parameter' {
        (Get-Command Add-ExcelModelMeasure).Parameters['FormatType'] | Should -Not -BeNullOrEmpty
    }
    It 'Has optional Description parameter' {
        (Get-Command Add-ExcelModelMeasure).Parameters['Description'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelModelMeasure).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Remove-ExcelModelMeasure
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Remove-ExcelModelMeasure' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelModelMeasure).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Remove-ExcelModelMeasure).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory MeasureName parameter' {
        $p = (Get-Command Remove-ExcelModelMeasure).Parameters['MeasureName']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelModelMeasure).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Add-ExcelModelRelationship
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Add-ExcelModelRelationship' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelModelRelationship).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Add-ExcelModelRelationship).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory ForeignKeyTable parameter' {
        $p = (Get-Command Add-ExcelModelRelationship).Parameters['ForeignKeyTable']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory ForeignKeyColumn parameter' {
        $p = (Get-Command Add-ExcelModelRelationship).Parameters['ForeignKeyColumn']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory PrimaryKeyTable parameter' {
        $p = (Get-Command Add-ExcelModelRelationship).Parameters['PrimaryKeyTable']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory PrimaryKeyColumn parameter' {
        $p = (Get-Command Add-ExcelModelRelationship).Parameters['PrimaryKeyColumn']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelModelRelationship).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Update-ExcelDataModel
# ═══════════════════════════════════════════════════════════════════════════════
Describe 'Update-ExcelDataModel' {
    It 'Has CmdletBinding' {
        (Get-Command Update-ExcelDataModel).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Update-ExcelDataModel).Parameters['WorkbookPath']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Update-ExcelDataModel).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
