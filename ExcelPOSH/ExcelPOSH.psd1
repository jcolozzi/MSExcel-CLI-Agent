@{
    RootModule        = 'ExcelPOSH.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b2c3d4e5-f6a7-8901-bcde-f23456789012'
    Author            = 'Excel-POSH'
    Description       = 'PowerShell Excel Workbook Automation via COM — 43 functions for workbook, worksheet, table, formatting, and metadata operations'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Workbook
        'Open-ExcelWorkbook'
        'Close-ExcelWorkbook'
        'New-ExcelWorkbook'
        'Save-ExcelWorkbook'
        'Get-ExcelWorkbookInfo'
        'Repair-ExcelWorkbook'
        'Copy-ExcelWorkbook'

        # Worksheet
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

        # Tables
        'Get-ExcelTable'
        'New-ExcelTable'
        'Remove-ExcelTable'
        'Resize-ExcelTable'
        'Get-ExcelTableData'
        'Add-ExcelTableRow'
        'Remove-ExcelTableRow'
        'Get-ExcelTableColumn'
        'Set-ExcelTableTotals'

        # Formatting
        'Set-ExcelCellFormat'
        'Set-ExcelNumberFormat'
        'Set-ExcelColumnWidth'
        'Set-ExcelRowHeight'
        'Set-ExcelAlignment'

        # Metadata
        'Get-ExcelDocumentProperty'
        'Set-ExcelDocumentProperty'
        'Get-ExcelConnection'
        'Get-ExcelProtection'
        'Set-ExcelProtection'
        'Get-ExcelComment'
        'Set-ExcelComment'
        'Get-ExcelTip'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
