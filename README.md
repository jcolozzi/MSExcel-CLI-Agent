# MSExcel-agent

**ExcelPOSH** is a PowerShell module providing 84 functions for Excel workbook automation via COM.

No MCP server needed — A custom excel-dev agent calls functions directly via terminal.

The agent will use the **ExcelPOSH** module to interact with Excel workbooks via COM automation.

Setup:
1. Clone or download the repo
2. Put the two .md files in C:\Users\\%USERNAME%\AppData\Roaming\Code\User\prompts folder (user level access) or create a .github\agents folder in your project folder and save the two .md files in the agent folder (_VS Code detects any .md files in the .github/agents folder of your workspace as custom agents_)
3. Replace the path in the .md files to the location of the **ExcelPOSH.psd1** module on your computer
4. Select excel-dev from the agent picker before prompting

## Module Structure

```
MSExcel-agent/
  ExcelPOSH/
    ExcelPOSH.psd1              # Module manifest — 84 public functions
    ExcelPOSH.psm1              # Root: constants, session state, dot-sourcing, exit handler
    Private/
      Session.ps1               # COM session management (5 functions)
      Utilities.ps1             # Value conversion & output formatting (2 functions)
    Public/
      WorkbookOps.ps1           # Workbook lifecycle (7 functions)
      WorksheetOps.ps1          # Worksheet & range operations (14 functions)
      TableOps.ps1              # ListObject/Table operations (9 functions)
      FormattingOps.ps1         # Cell formatting (5 functions + ConvertTo-RGBColor helper)
      MetadataOps.ps1           # Properties, connections, protection, comments (8 functions)
      FilterSortOps.ps1         # AutoFilter & sorting (4 functions)
      ConditionalFormatOps.ps1  # Conditional formatting rules (4 functions)
      DataValidationOps.ps1     # Data validation (3 functions)
      ViewOps.ps1               # Freeze panes, visibility, grouping (6 functions)
      HyperlinkOps.ps1          # Hyperlinks (3 functions)
      ClipboardOps.ps1          # Copy, replace, move ranges (3 functions)
      PrintOps.ps1              # Page setup & PDF export (3 functions)
      ImageShapeOps.ps1         # Images & shapes (3 functions)
      PivotTableOps.ps1         # Pivot tables (3 functions)
      ChartOps.ps1              # Charts (4 functions)
      ImportOps.ps1             # CSV & text import (2 functions)
      SparklineOps.ps1          # Sparklines (3 functions)
  Tests/
    ExcelPOSH.Module.Tests.ps1  # Module loading & export validation
    WorkbookOps.Tests.ps1       # Workbook function tests
    WorksheetOps.Tests.ps1      # Worksheet & range function tests
    TableOps.Tests.ps1          # Table function tests
    FormattingOps.Tests.ps1     # Formatting function tests
    MetadataOps.Tests.ps1       # Metadata function tests
    FilterSortOps.Tests.ps1     # Filter & sort tests
    ConditionalFormatOps.Tests.ps1 # Conditional format tests
    DataValidationOps.Tests.ps1 # Data validation tests
    ViewOps.Tests.ps1           # View ops tests
    HyperlinkOps.Tests.ps1      # Hyperlink tests
    ClipboardOps.Tests.ps1      # Clipboard tests
    PrintOps.Tests.ps1          # Print & page setup tests
    ImageShapeOps.Tests.ps1     # Image & shape tests
    PivotTableOps.Tests.ps1     # Pivot table tests
    ChartOps.Tests.ps1          # Chart tests
    ImportOps.Tests.ps1         # Import tests
    SparklineOps.Tests.ps1      # Sparkline tests
```

## Available Functions (84 public)

| Category | Functions |
|----------|-----------|
| **Workbook** | `Open-ExcelWorkbook`, `Close-ExcelWorkbook`, `New-ExcelWorkbook`, `Save-ExcelWorkbook`, `Get-ExcelWorkbookInfo`, `Repair-ExcelWorkbook`, `Copy-ExcelWorkbook` |
| **Worksheet** | `Get-ExcelWorksheet`, `New-ExcelWorksheet`, `Remove-ExcelWorksheet`, `Rename-ExcelWorksheet`, `Copy-ExcelWorksheet`, `Move-ExcelWorksheet` |
| **Range** | `Get-ExcelRange`, `Set-ExcelRange`, `Clear-ExcelRange`, `Get-ExcelUsedRange`, `Find-ExcelValue` |
| **Named Ranges** | `Get-ExcelNamedRange`, `New-ExcelNamedRange`, `Remove-ExcelNamedRange` |
| **Tables** | `Get-ExcelTable`, `New-ExcelTable`, `Remove-ExcelTable`, `Resize-ExcelTable`, `Get-ExcelTableData`, `Add-ExcelTableRow`, `Remove-ExcelTableRow`, `Get-ExcelTableColumn`, `Set-ExcelTableTotals` |
| **Formatting** | `Set-ExcelCellFormat`, `Set-ExcelNumberFormat`, `Set-ExcelColumnWidth`, `Set-ExcelRowHeight`, `Set-ExcelAlignment` |
| **Filter & Sort** | `Set-ExcelAutoFilter`, `Remove-ExcelAutoFilter`, `Sort-ExcelRange`, `Get-ExcelAutoFilter` |
| **Conditional Format** | `Add-ExcelConditionalFormat`, `Get-ExcelConditionalFormat`, `Remove-ExcelConditionalFormat`, `Clear-ExcelConditionalFormat` |
| **Data Validation** | `Set-ExcelDataValidation`, `Get-ExcelDataValidation`, `Remove-ExcelDataValidation` |
| **View** | `Set-ExcelFreezePane`, `Get-ExcelFreezePane`, `Set-ExcelSheetVisibility`, `Get-ExcelSheetVisibility`, `Set-ExcelGrouping`, `Set-ExcelOutlineLevel` |
| **Hyperlinks** | `Set-ExcelHyperlink`, `Get-ExcelHyperlink`, `Remove-ExcelHyperlink` |
| **Clipboard** | `Copy-ExcelRange`, `Update-ExcelValue`, `Move-ExcelRange` |
| **Print & Page Setup** | `Set-ExcelPageSetup`, `Get-ExcelPageSetup`, `Export-ExcelToPdf` |
| **Images & Shapes** | `Add-ExcelImage`, `Get-ExcelShape`, `Remove-ExcelShape` |
| **Pivot Tables** | `New-ExcelPivotTable`, `Get-ExcelPivotTable`, `Update-ExcelPivotTable` |
| **Charts** | `New-ExcelChart`, `Get-ExcelChart`, `Set-ExcelChart`, `Export-ExcelChart` |
| **Import** | `Import-ExcelCsv`, `Import-ExcelText` |
| **Sparklines** | `Add-ExcelSparkline`, `Get-ExcelSparkline`, `Remove-ExcelSparkline` |
| **Properties** | `Get-ExcelDocumentProperty`, `Set-ExcelDocumentProperty` |
| **Connections** | `Get-ExcelConnection` |
| **Protection** | `Get-ExcelProtection`, `Set-ExcelProtection` |
| **Comments** | `Get-ExcelComment`, `Set-ExcelComment` |
| **Tips** | `Get-ExcelTip` |
