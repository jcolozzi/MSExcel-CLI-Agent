# Tests/ChartOps.Tests.ps1
# Parameter-validation tests for ChartOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'New-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command New-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command New-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command New-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command New-ExcelChart).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartType parameter with validation' {
        $param = (Get-Command New-ExcelChart).Parameters['ChartType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command New-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Title parameter' {
        (Get-Command New-ExcelChart).Parameters['Title'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Left parameter' {
        (Get-Command New-ExcelChart).Parameters['Left'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Top parameter' {
        (Get-Command New-ExcelChart).Parameters['Top'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command New-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command New-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command New-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Get-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Get-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Set-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Set-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartType parameter with validation' {
        $param = (Get-Command Set-ExcelChart).Parameters['ChartType']
        $param | Should -Not -BeNullOrEmpty
        $validateSet = $param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $validateSet | Should -Not -BeNullOrEmpty
    }
    It 'Has Title parameter' {
        (Get-Command Set-ExcelChart).Parameters['Title'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Left parameter' {
        (Get-Command Set-ExcelChart).Parameters['Left'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Top parameter' {
        (Get-Command Set-ExcelChart).Parameters['Top'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Set-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Set-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SourceRange parameter' {
        (Get-Command Set-ExcelChart).Parameters['SourceRange'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Set-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Export-ExcelChart' {
    It 'Has CmdletBinding' {
        (Get-Command Export-ExcelChart).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Export-ExcelChart).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter' {
        (Get-Command Export-ExcelChart).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has ChartName parameter' {
        (Get-Command Export-ExcelChart).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has OutputPath parameter' {
        (Get-Command Export-ExcelChart).Parameters['OutputPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Width parameter' {
        (Get-Command Export-ExcelChart).Parameters['Width'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Height parameter' {
        (Get-Command Export-ExcelChart).Parameters['Height'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson parameter' {
        (Get-Command Export-ExcelChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelChartSeries' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChartSeries).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory ChartName parameter' {
        $p = (Get-Command Set-ExcelChartSeries).Parameters['ChartName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SeriesIndex parameter' {
        $p = (Get-Command Set-ExcelChartSeries).Parameters['SeriesIndex']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional Name/Values/XValues/FillColor/LineColor/LineWeight params' {
        $cmd = Get-Command Set-ExcelChartSeries
        $cmd.Parameters['Name']       | Should -Not -BeNullOrEmpty
        $cmd.Parameters['Values']     | Should -Not -BeNullOrEmpty
        $cmd.Parameters['XValues']    | Should -Not -BeNullOrEmpty
        $cmd.Parameters['FillColor']  | Should -Not -BeNullOrEmpty
        $cmd.Parameters['LineColor']  | Should -Not -BeNullOrEmpty
        $cmd.Parameters['LineWeight'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelChartSeries).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-ExcelChartAxis' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChartAxis).CmdletBinding | Should -BeTrue
    }
    It 'Has AxisType with ValidateSet' {
        $p = (Get-Command Set-ExcelChartAxis).Parameters['AxisType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Category'
        $vs.ValidValues | Should -Contain 'Value'
        $vs.ValidValues | Should -Contain 'SeriesAxis'
    }
    It 'Has AxisGroup with ValidateSet defaulting to Primary' {
        $p = (Get-Command Set-ExcelChartAxis).Parameters['AxisGroup']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.ValidValues | Should -Contain 'Primary'
        $vs.ValidValues | Should -Contain 'Secondary'
    }
    It 'Has optional Title/MinimumScale/MaximumScale/NumberFormat params' {
        $cmd = Get-Command Set-ExcelChartAxis
        $cmd.Parameters['Title']        | Should -Not -BeNullOrEmpty
        $cmd.Parameters['MinimumScale'] | Should -Not -BeNullOrEmpty
        $cmd.Parameters['MaximumScale'] | Should -Not -BeNullOrEmpty
        $cmd.Parameters['NumberFormat'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelChartAxis).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-ExcelChartLegend' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChartLegend).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Show parameter' {
        $p = (Get-Command Set-ExcelChartLegend).Parameters['Show']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has Position with ValidateSet' {
        $p = (Get-Command Set-ExcelChartLegend).Parameters['Position']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Bottom'
        $vs.ValidValues | Should -Contain 'Top'
        $vs.ValidValues | Should -Contain 'Left'
        $vs.ValidValues | Should -Contain 'Right'
        $vs.ValidValues | Should -Contain 'Corner'
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelChartLegend).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Set-ExcelChartDataLabels' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelChartDataLabels).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory SeriesIndex parameter' {
        $p = (Get-Command Set-ExcelChartDataLabels).Parameters['SeriesIndex']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has optional ShowValue/ShowCategory/ShowPercentage/NumberFormat params' {
        $cmd = Get-Command Set-ExcelChartDataLabels
        $cmd.Parameters['ShowValue']      | Should -Not -BeNullOrEmpty
        $cmd.Parameters['ShowCategory']   | Should -Not -BeNullOrEmpty
        $cmd.Parameters['ShowPercentage'] | Should -Not -BeNullOrEmpty
        $cmd.Parameters['NumberFormat']   | Should -Not -BeNullOrEmpty
    }
    It 'Has Position with ValidateSet' {
        $p = (Get-Command Set-ExcelChartDataLabels).Parameters['Position']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Center'
        $vs.ValidValues | Should -Contain 'BestFit'
    }
    It 'Has AsJson switch' {
        (Get-Command Set-ExcelChartDataLabels).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
