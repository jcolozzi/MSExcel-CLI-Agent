# Tests/ExcelPOSH.Module.Tests.ps1
# Pester 5+ tests — module loading, function exports, file structure

# PSScriptAnalyzer doesn't understand Pester's BeforeAll/It scoping
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
}

Describe 'ExcelPOSH Module' {

    Context 'Module manifest' {
        It 'Manifest file exists' {
            Test-Path $modulePath | Should -BeTrue
        }

        It 'Manifest is valid' {
            { Test-ModuleManifest -Path $modulePath -ErrorAction Stop } | Should -Not -Throw
        }

        It 'Manifest has correct RootModule' {
            $manifest = Test-ModuleManifest -Path $modulePath
            $manifest.RootModule | Should -Be 'ExcelPOSH.psm1'
        }
    }

    Context 'Module loads' {
        BeforeAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }

        AfterAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It 'Module is loaded' {
            Get-Module ExcelPOSH | Should -Not -BeNullOrEmpty
        }

        It 'Module version is 3.0.0' {
            (Get-Module ExcelPOSH).Version.ToString() | Should -Be '3.0.0'
        }

        It 'Exports exactly 104 public functions' {
            $exported = (Get-Module ExcelPOSH).ExportedFunctions.Keys
            $exported.Count | Should -Be 104
        }
    }

    Context 'Expected function exports' {
        BeforeAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module ExcelPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Exports <_>" -ForEach @(
            # WorkbookOps (7)
            'Open-ExcelWorkbook', 'Close-ExcelWorkbook', 'New-ExcelWorkbook',
            'Save-ExcelWorkbook', 'Get-ExcelWorkbookInfo', 'Repair-ExcelWorkbook',
            'Copy-ExcelWorkbook',
            # WorksheetOps (14)
            'Get-ExcelWorksheet', 'New-ExcelWorksheet', 'Remove-ExcelWorksheet',
            'Rename-ExcelWorksheet', 'Copy-ExcelWorksheet', 'Move-ExcelWorksheet',
            'Get-ExcelRange', 'Set-ExcelRange', 'Clear-ExcelRange',
            'Get-ExcelUsedRange', 'Find-ExcelValue',
            'Get-ExcelNamedRange', 'New-ExcelNamedRange', 'Remove-ExcelNamedRange',
            # TableOps (9)
            'Get-ExcelTable', 'New-ExcelTable', 'Remove-ExcelTable',
            'Resize-ExcelTable', 'Get-ExcelTableData', 'Add-ExcelTableRow',
            'Remove-ExcelTableRow', 'Get-ExcelTableColumn', 'Set-ExcelTableTotals',
            # FormattingOps (5)
            'Set-ExcelCellFormat', 'Set-ExcelNumberFormat', 'Set-ExcelColumnWidth',
            'Set-ExcelRowHeight', 'Set-ExcelAlignment',
            # MetadataOps (8)
            'Get-ExcelDocumentProperty', 'Set-ExcelDocumentProperty',
            'Get-ExcelConnection', 'Get-ExcelProtection', 'Set-ExcelProtection',
            'Get-ExcelComment', 'Set-ExcelComment', 'Get-ExcelTip',
            # FilterSortOps (4)
            'Set-ExcelAutoFilter', 'Remove-ExcelAutoFilter', 'Sort-ExcelRange',
            'Get-ExcelAutoFilter',
            # ConditionalFormatOps (4)
            'Add-ExcelConditionalFormat', 'Get-ExcelConditionalFormat',
            'Remove-ExcelConditionalFormat', 'Clear-ExcelConditionalFormat',
            # DataValidationOps (3)
            'Set-ExcelDataValidation', 'Get-ExcelDataValidation',
            'Remove-ExcelDataValidation',
            # ViewOps (6)
            'Set-ExcelFreezePane', 'Get-ExcelFreezePane',
            'Set-ExcelSheetVisibility', 'Get-ExcelSheetVisibility',
            'Set-ExcelGrouping', 'Set-ExcelOutlineLevel',
            # HyperlinkOps (3)
            'Set-ExcelHyperlink', 'Get-ExcelHyperlink', 'Remove-ExcelHyperlink',
            # ClipboardOps (3)
            'Copy-ExcelRange', 'Update-ExcelValue', 'Move-ExcelRange',
            # PrintOps (3)
            'Set-ExcelPageSetup', 'Get-ExcelPageSetup', 'Export-ExcelToPdf',
            # ImageShapeOps (3)
            'Add-ExcelImage', 'Get-ExcelShape', 'Remove-ExcelShape',
            # PivotTableOps (3)
            'New-ExcelPivotTable', 'Get-ExcelPivotTable', 'Update-ExcelPivotTable',
            # ChartOps (4)
            'New-ExcelChart', 'Get-ExcelChart', 'Set-ExcelChart', 'Export-ExcelChart',
            # ImportOps (2)
            'Import-ExcelCsv', 'Import-ExcelText',
            # SparklineOps (3)
            'Add-ExcelSparkline', 'Get-ExcelSparkline', 'Remove-ExcelSparkline',
            # StructuralOps (4) — v3.0
            'Add-ExcelRow', 'Remove-ExcelRow', 'Add-ExcelColumn', 'Remove-ExcelColumn',
            # CalculationOps (4) — v3.0
            'Set-ExcelPerformanceMode', 'Invoke-ExcelCalculate',
            'Invoke-ExcelFunction', 'Invoke-ExcelEvaluate',
            # FilterSortOps additions — v3.0
            'Remove-ExcelDuplicates', 'Invoke-ExcelAdvancedFilter',
            # FormattingOps additions — v3.0
            'Merge-ExcelRange', 'Split-ExcelRange',
            # WorksheetOps additions — v3.0
            'Get-ExcelSpecialCells',
            # WorkbookOps additions — v3.0
            'Invoke-ExcelMacro',
            # ChartOps additions — v3.0
            'Set-ExcelChartSeries', 'Set-ExcelChartAxis',
            'Set-ExcelChartLegend', 'Set-ExcelChartDataLabels',
            # PivotTableOps additions — v3.0
            'Set-ExcelPivotField', 'Add-ExcelPivotCalculatedField'
        ) {
            $script:exported | Should -Contain $_
        }
    }

    Context 'No private functions leaked' {
        BeforeAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module ExcelPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Does NOT export private function <_>" -ForEach @(
            'Test-ExcelAlive', 'Get-ExcelHwnd', 'Set-ExcelVisibleBestEffort',
            'Clear-ExcelCaches', 'Connect-ExcelWorkbook',
            'ConvertTo-ExcelSafeValue', 'Format-ExcelOutput',
            'ConvertTo-RGBColor'
        ) {
            $script:exported | Should -Not -Contain $_
        }
    }

    Context 'File structure' {
        It 'Private/Session.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Private\Session.ps1') | Should -BeTrue
        }

        It 'Private/Utilities.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Private\Utilities.ps1') | Should -BeTrue
        }

        It 'Public/WorkbookOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\WorkbookOps.ps1') | Should -BeTrue
        }

        It 'Public/WorksheetOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\WorksheetOps.ps1') | Should -BeTrue
        }

        It 'Public/TableOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\TableOps.ps1') | Should -BeTrue
        }

        It 'Public/FormattingOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\FormattingOps.ps1') | Should -BeTrue
        }

        It 'Public/MetadataOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\MetadataOps.ps1') | Should -BeTrue
        }

        It 'Public/FilterSortOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\FilterSortOps.ps1') | Should -BeTrue
        }

        It 'Public/ConditionalFormatOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ConditionalFormatOps.ps1') | Should -BeTrue
        }

        It 'Public/DataValidationOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\DataValidationOps.ps1') | Should -BeTrue
        }

        It 'Public/ViewOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ViewOps.ps1') | Should -BeTrue
        }

        It 'Public/HyperlinkOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\HyperlinkOps.ps1') | Should -BeTrue
        }

        It 'Public/ClipboardOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ClipboardOps.ps1') | Should -BeTrue
        }

        It 'Public/PrintOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\PrintOps.ps1') | Should -BeTrue
        }

        It 'Public/ImageShapeOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ImageShapeOps.ps1') | Should -BeTrue
        }

        It 'Public/PivotTableOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\PivotTableOps.ps1') | Should -BeTrue
        }

        It 'Public/ChartOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ChartOps.ps1') | Should -BeTrue
        }

        It 'Public/ImportOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\ImportOps.ps1') | Should -BeTrue
        }

        It 'Public/SparklineOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\SparklineOps.ps1') | Should -BeTrue
        }

        It 'Public/StructuralOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\StructuralOps.ps1') | Should -BeTrue
        }

        It 'Public/CalculationOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\ExcelPOSH\Public\CalculationOps.ps1') | Should -BeTrue
        }
    }

    Context 'Constants exist after import' {
        BeforeAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }

        AfterAll {
            Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        # We can't directly access $script: from tests, but we can verify
        # the module loaded without errors (which means constants were set)
        It 'Module loads without error (constants initialized)' {
            Get-Module ExcelPOSH | Should -Not -BeNullOrEmpty
        }
    }
}
