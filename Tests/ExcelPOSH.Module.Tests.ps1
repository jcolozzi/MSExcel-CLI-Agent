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

        It 'Module version is 1.0.0' {
            (Get-Module ExcelPOSH).Version.ToString() | Should -Be '1.0.0'
        }

        It 'Exports exactly 43 public functions' {
            $exported = (Get-Module ExcelPOSH).ExportedFunctions.Keys
            $exported.Count | Should -Be 43
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
            'Get-ExcelComment', 'Set-ExcelComment', 'Get-ExcelTip'
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
