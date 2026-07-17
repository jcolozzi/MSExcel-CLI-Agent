# Excel Workbook Data Analysis Skill

Use this skill when you need to inspect, clean, analyze, and summarize data inside Microsoft Excel workbooks with the help of the **ExcelPOSH** PowerShell module.

## When to Use This Skill

Use this skill when you need to:
- **Profile a workbook** and understand its sheets, tables, formulas, and named ranges
- **Audit data quality** for blanks, duplicates, inconsistent values, and obvious errors
- **Analyze data** with formulas, lookup functions, statistical functions, and time-based calculations
- **Summarize trends** with PivotTables, charts, and conditional formatting
- **Prepare a workbook for reporting** without turning the task into a VBA development project

## Core Analysis Workflow

This skill follows a practical Excel-analysis flow based on common workbook analytics practices:

### 1. Profile the workbook
Start by understanding the file structure before analyzing anything.

- Inspect worksheets, tables, named ranges, and used range
- Identify relevant data areas and headers
- Note any formulas, comments, or protected ranges that may affect interpretation

Useful ExcelPOSH commands:
```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1" -Force
$wb = "C:\path\to\workbook.xlsx"

Open-ExcelWorkbook -WorkbookPath $wb -AsJson
Get-ExcelWorkbookInfo -WorkbookPath $wb -AsJson
Get-ExcelWorksheet -WorkbookPath $wb -AsJson
Get-ExcelTable -WorkbookPath $wb -AsJson
Get-ExcelRange -WorkbookPath $wb -SheetName "Data" -Range "A1:E100" -AsJson
```

### 2. Clean and shape the data
Before analysis, make sure the data is usable.

- Check for blank cells, duplicates, and mixed formats
- Sort and filter the relevant range
- Import external CSV/text data if needed
- Keep source data intact and place derived results on a separate summary sheet when practical

Useful ExcelPOSH commands:
```powershell
Find-ExcelValue -WorkbookPath $wb -SheetName "Data" -SearchText "Error" -AsJson
Remove-ExcelDuplicates -WorkbookPath $wb -SheetName "Data" -Range "A1:D100" -AsJson
Sort-ExcelRange -WorkbookPath $wb -SheetName "Data" -Range "A1:D100" -SortKey1 "B1" -Order1 "ascending" -AsJson
Import-ExcelCsv -WorkbookPath $wb -SheetName "Import" -CsvPath "C:\data.csv" -Delimiter "comma" -HasHeaders -AsJson
```

### 3. Analyze with formulas and workbook-native logic
Excel analysis usually starts with formulas and structured workbook logic.

Common techniques include:
- Lookup and matching: `XLOOKUP`, `VLOOKUP`, `INDEX/MATCH`
- Logic: `IF`, `IFS`, `CHOOSE`, `SWITCH`
- Text cleanup: `TRIM`, `TEXT`, `LEFT`, `MID`, `RIGHT`, `CONCAT`
- Dates and times: `DATE`, `EOMONTH`, `YEARFRAC`, `DATEDIF`, `NETWORKDAYS`
- Statistical analysis: `AVERAGE`, `MEDIAN`, `MODE`, `VAR`, `STDEV`, `CORREL`, `RANK`, `SMALL`
- Financial analysis: `NPV`, `PMT`, `PPMT`, `XIRR`, `XNPV`

Useful ExcelPOSH commands:
```powershell
Invoke-ExcelEvaluate -WorkbookPath $wb -SheetName "Data" -Expression "AVERAGE(B2:B100)" -AsJson
Invoke-ExcelFunction -WorkbookPath $wb -FunctionName "SUM" -Args @(1, 2, 3) -AsJson
Invoke-ExcelCalculate -WorkbookPath $wb -SheetName "Data" -AsJson
```

### 4. Summarize with PivotTables and charts
When the analysis needs a compact story, turn the results into a PivotTable or chart.

- Create summaries by category, region, month, or product
- Add dashboards and visual highlights
- Use conditional formatting to emphasize exceptions or trends

Useful ExcelPOSH commands:
```powershell
New-ExcelPivotTable -WorkbookPath $wb -SheetName "Data" -SourceRange "A1:D100" -PivotTableName "SalesPivot" -DestinationSheet "Pivot" -RowFields @("Region") -DataFields @(@{Name="Revenue";Function="Sum"}) -AsJson
New-ExcelChart -WorkbookPath $wb -SheetName "Data" -SourceRange "A1:B10" -ChartType "ColumnClustered" -Title "Sales" -AsJson
```

### 5. Report clearly
The final output should be understandable to a business user or stakeholder.

- Summarize key findings in plain language
- Include the formulas, ranges, or PivotTable structure used
- Call out assumptions, blanks, and data caveats
- Offer to place the results into a separate report sheet or export a chart

## Map dependencies & relationships (graph)

When a question is about **structure or relationships** rather than values — "what depends on this
table?", "which named ranges are unused?", "how are these sheets connected?", "what are the foreign
keys?" — build a dependency graph and query it instead of reading formulas by hand.

```powershell
# Build once: writes graph.json + an interactive index.html to <workbook-folder>\excel-graph-out\
Export-ExcelGraph -WorkbookPath $wb -FormulaMode Both

# Query repeatedly — these auto-locate the graph.json next to the workbook (no re-scan, no Excel):
Get-ExcelGraphQuery -Action summary   -WorkbookPath $wb -AsJson   # counts by node group + edge kind
Get-ExcelGraphQuery -Action neighbors -WorkbookPath $wb -Node "table:Orders" -Depth 2 -AsJson
Get-ExcelGraphQuery -Action impact    -WorkbookPath $wb -Node "table:Customers" -AsJson
Get-ExcelGraphQuery -Action orphans   -WorkbookPath $wb -AsJson   # nothing points to these
```

- **Check for an existing graph first:** `Get-ExcelGraphQuery -WorkbookPath $wb` auto-finds
  `<workbook-folder>\excel-graph-out\graph.json`. If it reports the graph is missing, run
  `Export-ExcelGraph` once, then query. Re-export only after the workbook's structure changes.
- **Two layers:** *structure* (sheets, tables, names, pivots, charts, connections, Power Query,
  Data Model, slicers, VBA) and *data relationships* (Data Model FKs, lookup-formula FKs,
  value-overlap inferred FKs, primary keys). The `index.html` viewer toggles between them and
  includes report panels (inferred-FK candidates, tables without relationships, high fan-in,
  unused named ranges).
- VBA code edges need *Trust access to the VBA project object model* (`enable_vba_trust.ps1`).

## Best Practices

- **Use structured tables where possible.** Tables make ranges easier to inspect and summarize.
- **Separate source data from analysis outputs.** Keep raw data intact and push results to a summary sheet.
- **Prefer workbook-native formulas over manual edits.** This makes the workbook auditable.
- **Keep the analysis transparent.** Explain which formulas or ExcelPOSH steps produced each result.
- **Be explicit about caveats.** Blank cells, text values, and date assumptions can strongly influence results.
- **Use performance settings when the workbook is large.** `Set-ExcelPerformanceMode` can help when recalculation or formatting is heavy.

## ExcelPOSH Guidance

- Use `-AsJson` whenever you need structured output for inspection.
- Prefer `Get-ExcelRange` and `Get-ExcelTableData` previews over full-workbook reads when the workbook is large.
- Use `Close-ExcelWorkbook` at the end of a session to release the COM connection.
- For deeper workbook automation or VBA features, hand off to the `@excel-dev` agent.

## Useful ExcelPOSH Functions

| Category | Functions |
|----------|-----------|
| **Workbook & sheets** | `Open-ExcelWorkbook`, `Get-ExcelWorkbookInfo`, `Get-ExcelWorksheet`, `Get-ExcelUsedRange`, `Get-ExcelNamedRange` |
| **Range & tables** | `Get-ExcelRange`, `Get-ExcelTable`, `Get-ExcelTableData`, `Find-ExcelValue` |
| **Cleaning & shaping** | `Sort-ExcelRange`, `Remove-ExcelDuplicates`, `Invoke-ExcelAdvancedFilter`, `Import-ExcelCsv`, `Import-ExcelText` |
| **Formulas & evaluation** | `Invoke-ExcelEvaluate`, `Invoke-ExcelFunction`, `Invoke-ExcelCalculate`, `Set-ExcelPerformanceMode` |
| **Reporting outputs** | `New-ExcelPivotTable`, `Update-ExcelPivotTable`, `New-ExcelChart`, `Set-ExcelChart`, `Export-ExcelChart` |
| **Workbook management** | `Save-ExcelWorkbook`, `Copy-ExcelWorkbook`, `Close-ExcelWorkbook` |
| **Dependency & data graph** | `Export-ExcelGraph`, `Get-ExcelGraphQuery`, `Import-ExcelGraph` |

## Example Outcome

A good analysis session usually ends with:
- a short summary of the key finding
- the relevant workbook ranges or formulas used
- a clear note on any assumptions or data caveats
- an offer to build a report sheet, chart, or PivotTable if the user wants a more polished deliverable
