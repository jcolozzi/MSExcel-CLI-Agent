@{
    RootModule        = 'ExcelPOSH.psm1'
    ModuleVersion     = '2.0.0'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'Excel-POSH'
    Description       = 'PowerShell Excel Workbook Automation via COM — 84 functions for workbook, worksheet, table, formatting, metadata, filter/sort, conditional format, data validation, view, hyperlink, clipboard, print, image/shape, pivot table, chart, import, and sparkline operations'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Workbook (7)
        'Open-ExcelWorkbook'
        'Close-ExcelWorkbook'
        'New-ExcelWorkbook'
        'Save-ExcelWorkbook'
        'Get-ExcelWorkbookInfo'
        'Repair-ExcelWorkbook'
        'Copy-ExcelWorkbook'

        # Worksheet (14)
        'Get-ExcelWorksheet'
        'New-ExcelWorksheet'
        'Remove-ExcelWorksheet'
        'Rename-ExcelWorksheet'
        'Copy-ExcelWorksheet'
        'Move-ExcelWorksheet'
        'Get-ExcelRange'
        'Set-ExcelRange'
        'Clear-ExcelRange'
        'Get-ExcelUsedRange'
        'Find-ExcelValue'
        'Get-ExcelNamedRange'
        'New-ExcelNamedRange'
        'Remove-ExcelNamedRange'

        # Tables (9)
        'Get-ExcelTable'
        'New-ExcelTable'
        'Remove-ExcelTable'
        'Resize-ExcelTable'
        'Get-ExcelTableData'
        'Add-ExcelTableRow'
        'Remove-ExcelTableRow'
        'Get-ExcelTableColumn'
        'Set-ExcelTableTotals'

        # Formatting (5)
        'Set-ExcelCellFormat'
        'Set-ExcelNumberFormat'
        'Set-ExcelColumnWidth'
        'Set-ExcelRowHeight'
        'Set-ExcelAlignment'

        # Metadata (8)
        'Get-ExcelDocumentProperty'
        'Set-ExcelDocumentProperty'
        'Get-ExcelConnection'
        'Get-ExcelProtection'
        'Set-ExcelProtection'
        'Get-ExcelComment'
        'Set-ExcelComment'
        'Get-ExcelTip'

        # Filter & Sort (4)
        'Set-ExcelAutoFilter'
        'Remove-ExcelAutoFilter'
        'Sort-ExcelRange'
        'Get-ExcelAutoFilter'

        # Conditional Formatting (4)
        'Add-ExcelConditionalFormat'
        'Get-ExcelConditionalFormat'
        'Remove-ExcelConditionalFormat'
        'Clear-ExcelConditionalFormat'

        # Data Validation (3)
        'Set-ExcelDataValidation'
        'Get-ExcelDataValidation'
        'Remove-ExcelDataValidation'

        # View — Freeze Panes, Sheet Visibility, Grouping (6)
        'Set-ExcelFreezePane'
        'Get-ExcelFreezePane'
        'Set-ExcelSheetVisibility'
        'Get-ExcelSheetVisibility'
        'Set-ExcelGrouping'
        'Set-ExcelOutlineLevel'

        # Hyperlinks (3)
        'Set-ExcelHyperlink'
        'Get-ExcelHyperlink'
        'Remove-ExcelHyperlink'

        # Clipboard — Copy, Replace, Move (3)
        'Copy-ExcelRange'
        'Replace-ExcelValue'
        'Move-ExcelRange'

        # Print & Page Setup (3)
        'Set-ExcelPageSetup'
        'Get-ExcelPageSetup'
        'Export-ExcelToPdf'

        # Images & Shapes (3)
        'Add-ExcelImage'
        'Get-ExcelShape'
        'Remove-ExcelShape'

        # Pivot Tables (3)
        'New-ExcelPivotTable'
        'Get-ExcelPivotTable'
        'Update-ExcelPivotTable'

        # Charts (4)
        'New-ExcelChart'
        'Get-ExcelChart'
        'Set-ExcelChart'
        'Export-ExcelChart'

        # Import (2)
        'Import-ExcelCsv'
        'Import-ExcelText'

        # Sparklines (3)
        'Add-ExcelSparkline'
        'Get-ExcelSparkline'
        'Remove-ExcelSparkline'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
