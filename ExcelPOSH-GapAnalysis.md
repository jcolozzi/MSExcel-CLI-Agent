# ExcelPOSH Gap Analysis & Missing Features

**Module:** ExcelPOSH v2.0.0  
**Date:** 2026-04-14  
**Scope:** AI-agent-driven Excel workbook automation via COM  
**Current Functions:** 84 exported (80 confirmed in source; 4 unaccounted in manifest)

---

## 1. Executive Summary

ExcelPOSH provides solid coverage of the most common Excel COM automation tasks — workbook lifecycle, range I/O, tables, formatting, filtering, charting, and pivot tables. However, several **high-value gaps** exist that limit the AI agent's ability to perform structural data transformations, leverage Excel's calculation engine, customize chart output, and manage workbook performance at scale.

The most impactful missing capabilities fall into three clusters:

1. **Structural range editing** (insert/delete rows/columns, merge/unmerge, remove duplicates) — the agent cannot reshape data without these.
2. **Excel calculation engine access** (`WorksheetFunction`, `Evaluate`, `Calculate`) — the agent must read data into PowerShell to compute anything, even simple SUMIFs.
3. **Performance controls** (`ScreenUpdating`, `Calculation = xlManual`, `EnableEvents`) — bulk operations on large workbooks are unnecessarily slow.

Addressing the 12 High-priority gaps below would raise effective coverage from ~60% to ~85% of what an AI agent typically needs.

---

## 2. Current Module Coverage (Brief)

| Domain | Functions | Status |
|---|---|---|
| Workbook lifecycle | 7 | Complete — open, close, new, save, info, repair, copy |
| Worksheet CRUD | 6 | Complete |
| Range read/write | 8 | Good — get, set, clear, usedRange, find, namedRange CRUD |
| Tables (ListObject) | 9 | Good — full CRUD + data + totals |
| Formatting | 5 | Core only — font, number, column width, row height, alignment |
| Filter & Sort | 4 | Basic — single/dual criteria AutoFilter, simple sort |
| Conditional Format | 4 | Good — add, get, remove, clear |
| Data Validation | 3 | Good |
| View / Layout | 6 | Good — freeze panes, visibility, grouping, outline |
| Hyperlinks | 3 | Complete |
| Clipboard / Move | 3 | Good — copy, update (find/replace), move |
| Print / Page Setup | 3 | Good — page setup + PDF export |
| Images & Shapes | 3 | Basic — add image, get/remove shape |
| Pivot Tables | 3 | Basic — create, query, refresh (no field customization) |
| Charts | 4 | Basic — create, query, set, export (no series/axis control) |
| Import | 2 | CSV and delimited text |
| Sparklines | 3 | Good |
| Metadata / Comments | 8 | Good — doc properties, connections, protection, comments, tips |

**Private helpers:** Session management (5), Utilities (2), `ConvertTo-RGBColor` (1).

---

## 3. Gap Analysis

### 3.1 HIGH Priority — Core Gaps Blocking Common Agent Workflows

---

#### H1. Insert / Delete Rows and Columns

| Attribute | Detail |
|---|---|
| **What's missing** | No way to insert or delete entire rows, columns, or cell ranges with shift. The agent cannot add a header row, insert a summary column, or remove blank rows. |
| **Why it matters** | Structural editing is one of the most frequent data-prep tasks an AI agent performs. Without it, the agent must recreate sheets from scratch. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Range.Insert(Shift)`, `Range.Delete(Shift)`, `Range.EntireRow.Insert`, `Range.EntireColumn.Delete` |
| **Suggested functions** | `Add-ExcelRow`, `Remove-ExcelRow`, `Add-ExcelColumn`, `Remove-ExcelColumn` (or a unified `Edit-ExcelRangeStructure`) |
| **Destination file** | `WorksheetOps.ps1` |

---

#### H2. Remove Duplicates

| Attribute | Detail |
|---|---|
| **What's missing** | No function to remove duplicate rows from a range or table. |
| **Why it matters** | Data deduplication is a top-3 data-cleaning operation. The agent currently has no way to do this except reading all data into PowerShell, deduplicating, clearing the range, and writing it back — slow and error-prone. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Range.RemoveDuplicates(Columns, Header)` |
| **Suggested function** | `Remove-ExcelDuplicates` |
| **Destination file** | `FilterSortOps.ps1` or new `DataOps.ps1` |

---

#### H3. Application.WorksheetFunction Access

| Attribute | Detail |
|---|---|
| **What's missing** | No wrapper to call Excel's built-in worksheet functions (VLOOKUP, SUMIF, COUNTIF, INDEX, MATCH, AVERAGEIF, etc.) from PowerShell. |
| **Why it matters** | The agent must read entire datasets into PowerShell to compute aggregates or lookups. `WorksheetFunction` lets Excel do the heavy lifting in-process, orders of magnitude faster for large data. |
| **Impact** | **High** |
| **Complexity** | **M** — need dynamic parameter dispatch since each function has different arity |
| **COM objects/methods** | `Application.WorksheetFunction.VLookup(...)`, `.SumIf(...)`, `.CountIf(...)`, etc. |
| **Suggested function** | `Invoke-ExcelFunction -FunctionName "SUMIF" -Args @($range, $criteria, $sumRange)` |
| **Destination file** | New `CalculationOps.ps1` |

---

#### H4. Application Performance Controls (ScreenUpdating, Calculation Mode, EnableEvents)

| Attribute | Detail |
|---|---|
| **What's missing** | No way to toggle `ScreenUpdating`, `Calculation` mode (automatic/manual), or `EnableEvents`. Every Set/Format/Add operation triggers a screen repaint and recalculation. |
| **Why it matters** | Bulk operations (formatting 1,000 cells, importing 50K rows) are 5–20× slower without these controls. The agent's session hangs or times out on large workbooks. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Application.ScreenUpdating`, `Application.Calculation` (xlAutomatic=−4105, xlManual=−4135), `Application.EnableEvents`, `Application.Calculate`, `Application.CalculateFull` |
| **Suggested functions** | `Set-ExcelPerformanceMode -ScreenUpdating $false -Calculation Manual`, `Invoke-ExcelCalculate [-Full]` |
| **Destination file** | New `CalculationOps.ps1` or `WorkbookOps.ps1` |

---

#### H5. Merge / Unmerge Cells

| Attribute | Detail |
|---|---|
| **What's missing** | No function to merge or unmerge cell ranges. No way to detect merged areas. |
| **Why it matters** | Merged cells are extremely common in user-facing reports and templates. The agent can't build formatted reports or read data from legacy sheets with merged headers without understanding merge areas. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Range.Merge()`, `Range.UnMerge()`, `Range.MergeCells` (bool), `Range.MergeArea` (returns the merged range) |
| **Suggested functions** | `Merge-ExcelRange`, `Split-ExcelRange` (unmerge), or add `-Merge`/`-Unmerge` switches to `Set-ExcelCellFormat` |
| **Destination file** | `FormattingOps.ps1` |

---

#### H6. Application.Evaluate — Formula Evaluation

| Attribute | Detail |
|---|---|
| **What's missing** | No way to evaluate an arbitrary Excel formula string and return the result, without writing it to a cell. |
| **Why it matters** | The agent often needs to compute a value (e.g., `"=SUMPRODUCT((A2:A100=""East"")*(B2:B100))"`) without modifying the workbook. `Evaluate` is the COM equivalent of a read-only formula engine. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Application.Evaluate(expression)` or `Worksheet.Evaluate(expression)` |
| **Suggested function** | `Invoke-ExcelEvaluate -WorkbookPath ... -SheetName ... -Expression "=SUM(A1:A10)"` |
| **Destination file** | `CalculationOps.ps1` |

---

#### H7. SpecialCells — Select by Cell Type

| Attribute | Detail |
|---|---|
| **What's missing** | No way to select cells by type: blanks, formulas, constants, errors, visible cells only, last cell. |
| **Why it matters** | The agent cannot efficiently find all formula cells (for auditing), all blank cells (for cleanup), or all error cells (for repair). Currently must read every cell and filter in PowerShell. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Range.SpecialCells(xlCellType, [xlValue])` — constants: `xlCellTypeConstants`, formulas: `xlCellTypeFormulas`, blanks: `xlCellTypeBlanks`, errors: values with `xlErrors`, visible: `xlCellTypeVisible`, lastCell: `xlCellTypeLastCell` |
| **Suggested function** | `Get-ExcelSpecialCells -WorkbookPath ... -SheetName ... -Range "A:Z" -CellType Blanks` |
| **Destination file** | `WorksheetOps.ps1` |

---

#### H8. Multi-Key Sort and Sort with Custom Order

| Attribute | Detail |
|---|---|
| **What's missing** | `Sort-ExcelRange` appears to support only a single sort key. No support for 3+ key sorting or custom sort orders. |
| **Why it matters** | Business data routinely needs multi-level sorting (e.g., Region asc → Date desc → Amount desc). Single-key sort is insufficient for real reporting. |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Worksheet.Sort.SortFields.Add(Key, Order)` (repeated), then `Sort.Apply`. Supports up to 64 sort keys. |
| **Suggested change** | Extend `Sort-ExcelRange` to accept an array of `@{Key; Order}` hashtables. |
| **Destination file** | `FilterSortOps.ps1` |

---

#### H9. Advanced Filter (Unique Values & Criteria Filter)

| Attribute | Detail |
|---|---|
| **What's missing** | No function to extract unique values from a column or apply criteria-based advanced filtering. |
| **Why it matters** | "Get unique values in column X" is one of the most common AI agent questions. Currently requires reading the full column and using `Sort-Object -Unique` in PowerShell, which is slow for large datasets. |
| **Impact** | **High** |
| **Complexity** | **M** |
| **COM objects/methods** | `Range.AdvancedFilter(Action, CriteriaRange, CopyToRange, Unique)` — `xlFilterInPlace`=1, `xlFilterCopy`=2 |
| **Suggested function** | `Invoke-ExcelAdvancedFilter` |
| **Destination file** | `FilterSortOps.ps1` |

---

#### H10. Chart Series, Axes, Legend, and Data Label Customization

| Attribute | Detail |
|---|---|
| **What's missing** | `New-ExcelChart` creates charts but `Set-ExcelChart` has limited customization. No way to add/remove/format individual series, configure axis scales/labels, toggle/format legends, or add data labels. |
| **Why it matters** | AI-generated charts almost always need axis titles, formatted data labels, or series color changes to be presentation-ready. Without these, the agent produces raw charts that need manual polish. |
| **Impact** | **High** |
| **Complexity** | **M** |
| **COM objects/methods** | `Chart.SeriesCollection(n)`, `Series.Format.Fill`, `Chart.Axes(xlCategory/xlValue)`, `Axis.HasTitle`, `Axis.AxisTitle.Text`, `Axis.MinimumScale`, `Axis.MaximumScale`, `Chart.HasLegend`, `Chart.Legend.Position`, `Series.HasDataLabels`, `Series.DataLabels.ShowValue` |
| **Suggested functions** | `Set-ExcelChartSeries`, `Set-ExcelChartAxis`, `Set-ExcelChartLegend`, `Set-ExcelChartDataLabels` — or expand `Set-ExcelChart` with sub-parameter sets |
| **Destination file** | `ChartOps.ps1` |

---

#### H11. PivotTable Field Customization (Subtotals, Calculated Fields, Number Format)

| Attribute | Detail |
|---|---|
| **What's missing** | `New-ExcelPivotTable` creates basic pivots, but there's no way to modify field subtotals, add calculated fields, change field number formats, or adjust field layout (Compact/Tabular/Outline). |
| **Why it matters** | Real-world pivot tables need calculated metrics (e.g., "Revenue per Unit"), suppressed subtotals, and formatted numbers. The agent creates pivots but can't refine them. |
| **Impact** | **High** |
| **Complexity** | **M** |
| **COM objects/methods** | `PivotField.Subtotals`, `PivotTable.CalculatedFields.Add(Name, Formula)`, `PivotField.NumberFormat`, `PivotField.LayoutForm` (xlTabular=0, xlOutline=1) |
| **Suggested functions** | `Set-ExcelPivotField`, `Add-ExcelPivotCalculatedField` |
| **Destination file** | `PivotTableOps.ps1` |

---

#### H12. Application.Run — Execute VBA Macros

| Attribute | Detail |
|---|---|
| **What's missing** | No way to execute an existing VBA macro in an `.xlsm` workbook. |
| **Why it matters** | Many enterprise workbooks contain VBA macros for business logic. The agent must be able to trigger them (e.g., "run the month-end close macro"). |
| **Impact** | **High** |
| **Complexity** | **S** |
| **COM objects/methods** | `Application.Run("WorkbookName!MacroName", arg1, arg2, ...)` |
| **Suggested function** | `Invoke-ExcelMacro -WorkbookPath ... -MacroName "Module1.CloseMonth" -Arguments @(2026, 3)` |
| **Destination file** | New `MacroOps.ps1` or `WorkbookOps.ps1` |

---

### 3.2 MEDIUM Priority — Important but Less Frequently Needed

---

#### M1. Range.TextToColumns — Parse Delimited Text

| Attribute | Detail |
|---|---|
| **What's missing** | No way to split a column of delimited text into multiple columns. |
| **Why it matters** | Common data-import cleanup step (e.g., "Full Name" → "First" + "Last"). |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Range.TextToColumns(Destination, DataType, TextQualifier, FieldInfo)` |
| **Suggested function** | `Split-ExcelColumn` |
| **Destination file** | `ImportOps.ps1` or new `DataOps.ps1` |

---

#### M2. Range.AutoFill / FillDown / FillRight

| Attribute | Detail |
|---|---|
| **What's missing** | No way to fill a series (dates, numbers) or copy a formula pattern down/across a range. |
| **Why it matters** | The agent frequently needs to extend formulas or number sequences to all rows. Without AutoFill, it must write each cell individually. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Range.AutoFill(Destination, Type)`, `Range.FillDown()`, `Range.FillRight()` |
| **Suggested function** | `Invoke-ExcelAutoFill` |
| **Destination file** | `WorksheetOps.ps1` or new `DataOps.ps1` |

---

#### M3. Range.Subtotal — Automatic Subtotals

| Attribute | Detail |
|---|---|
| **What's missing** | No function to insert automatic subtotals with group breaks. |
| **Why it matters** | Classic reporting operation: subtotals by region, category, date. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Range.Subtotal(GroupBy, Function, TotalList, Replace, PageBreaks, SummaryBelowData)` |
| **Suggested function** | `Add-ExcelSubtotal` |
| **Destination file** | `FilterSortOps.ps1` or new `DataOps.ps1` |

---

#### M4. Worksheet Tab Color

| Attribute | Detail |
|---|---|
| **What's missing** | No way to set or get worksheet tab colors. |
| **Why it matters** | Visual organization of multi-sheet workbooks. Agents building multi-tab reports benefit from color-coding sheets. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Worksheet.Tab.Color` (RGB long), `Worksheet.Tab.ColorIndex`, `Worksheet.Tab.ThemeColor` |
| **Suggested function** | `Set-ExcelSheetTab -Color "#FF0000"` (extend `Rename-ExcelWorksheet` or new function) |
| **Destination file** | `WorksheetOps.ps1` |

---

#### M5. Formula Auditing (Precedents / Dependents)

| Attribute | Detail |
|---|---|
| **What's missing** | No way to trace which cells feed into a formula (precedents) or which formulas depend on a cell (dependents). |
| **Why it matters** | Critical for debugging complex spreadsheets. The agent can identify circular references, broken formulas, or impact analysis. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Range.Precedents`, `Range.Dependents`, `Range.HasFormula`, `Range.Formula` |
| **Suggested function** | `Get-ExcelFormulaDependencies` |
| **Destination file** | `WorksheetOps.ps1` or `MetadataOps.ps1` |

---

#### M6. Named Styles (Workbook.Styles)

| Attribute | Detail |
|---|---|
| **What's missing** | No support for named styles (e.g., "Heading 1", "Currency", custom styles). Formatting is applied per-cell/range only. |
| **Why it matters** | Consistent formatting across large reports. Named styles enable the agent to apply corporate templates efficiently. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Workbook.Styles`, `Style.Font`, `Style.Interior`, `Style.NumberFormat`, `Range.Style = "StyleName"` |
| **Suggested functions** | `Get-ExcelStyle`, `New-ExcelStyle`, `Set-ExcelRangeStyle` |
| **Destination file** | `FormattingOps.ps1` |

---

#### M7. Export to XPS Format

| Attribute | Detail |
|---|---|
| **What's missing** | `Export-ExcelToPdf` supports PDF only. The same COM method supports XPS. |
| **Why it matters** | Minor — XPS is less common but trivial to add since the infrastructure exists. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Workbook.ExportAsFixedFormat(xlTypeXPS)` — constant already defined in `$script:XL_FIXED_FORMAT` |
| **Suggested change** | Add `-Format` parameter to `Export-ExcelToPdf` (or rename to `Export-ExcelFixedFormat`) accepting `PDF` or `XPS`. |
| **Destination file** | `PrintOps.ps1` |

---

#### M8. Worksheet.PrintOut — Direct Printing

| Attribute | Detail |
|---|---|
| **What's missing** | No way to send a worksheet or range directly to a printer. |
| **Why it matters** | Enterprise workflows sometimes require physical printouts (invoices, packing slips). Agent can set page setup but not actually print. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Worksheet.PrintOut(From, To, Copies, Preview, ActivePrinter)` |
| **Suggested function** | `Send-ExcelPrint` or `Invoke-ExcelPrint` |
| **Destination file** | `PrintOps.ps1` |

---

#### M9. Application.StatusBar — Progress Feedback

| Attribute | Detail |
|---|---|
| **What's missing** | No way to set/read the Excel status bar text. |
| **Why it matters** | For long-running agent operations (bulk import, format 10K cells), the status bar provides the only visual progress indicator to anyone watching the Excel window. |
| **Impact** | Medium |
| **Complexity** | S |
| **COM objects/methods** | `Application.StatusBar = "Processing row 500 of 10000..."`, reset with `Application.StatusBar = $false` |
| **Suggested function** | `Set-ExcelStatusBar -Text "..."` / `Set-ExcelStatusBar -Reset` |
| **Destination file** | `WorkbookOps.ps1` or `MetadataOps.ps1` |

---

#### M10. CopyFromRecordset — ADO Recordset Dump

| Attribute | Detail |
|---|---|
| **What's missing** | No way to dump an ADO recordset directly into a range. |
| **Why it matters** | Fastest way to import SQL query results into Excel. Orders of magnitude faster than cell-by-cell writes for datasets > 1,000 rows. |
| **Impact** | Medium |
| **Complexity** | M |
| **COM objects/methods** | `Range.CopyFromRecordset(Recordset, MaxRows, MaxColumns)` — requires ADODB.Recordset or DAO.Recordset |
| **Suggested function** | `Import-ExcelRecordset -WorkbookPath ... -SheetName ... -ConnectionString "..." -Query "SELECT ..."` |
| **Destination file** | `ImportOps.ps1` |

---

### 3.3 LOW Priority — Nice to Have / Niche

---

#### L1. Threaded Comments (CommentsThreaded)

| Attribute | Detail |
|---|---|
| **What's missing** | `Get/Set-ExcelComment` uses legacy `Comment` objects. Excel 365 uses `CommentsThreaded` for reply-based comments. |
| **Why it matters** | Modern Excel files may not show legacy comments in the new comment pane. Agent feedback via threaded comments is more visible. |
| **Impact** | Low |
| **Complexity** | M |
| **COM objects/methods** | `Worksheet.CommentsThreaded`, `Range.AddCommentThreaded(Text)`, `CommentThreaded.Replies.Add(Text)` |

---

#### L2. Named Ranges with Complex Formulas

| Attribute | Detail |
|---|---|
| **What's missing** | `New-ExcelNamedRange` creates names referring to cell ranges. No support for formula-based names (e.g., `=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)` for dynamic ranges). |
| **Why it matters** | Dynamic named ranges are used in dashboards and chart sources. |
| **Impact** | Low |
| **Complexity** | S |
| **COM objects/methods** | `Workbook.Names.Add(Name, RefersTo)` where `RefersTo` is a formula string |

---

#### L3. Worksheet Scenarios (What-If)

| Attribute | Detail |
|---|---|
| **What's missing** | No support for Excel's Scenario Manager. |
| **Why it matters** | Rarely used in automation. Mostly interactive. |
| **Impact** | Low |
| **Complexity** | M |
| **COM objects/methods** | `Worksheet.Scenarios.Add(Name, ChangingCells, Values)`, `.Show()` |

---

#### L4. PivotTable Slicers

| Attribute | Detail |
|---|---|
| **What's missing** | No support for adding or manipulating slicers on pivot tables. |
| **Why it matters** | Slicers are a visual interactivity feature; less useful for headless agent automation. |
| **Impact** | Low |
| **Complexity** | M |
| **COM objects/methods** | `Workbook.SlicerCaches.Add2(PivotTable, SourceField)`, `SlicerCache.Slicers.Add()` |

---

#### L5. Shape Formatting and Manipulation

| Attribute | Detail |
|---|---|
| **What's missing** | `Add-ExcelImage` adds images and `Get/Remove-ExcelShape` handles basic shape CRUD, but no support for shape text, fill color, line style, positioning, or grouping. |
| **Why it matters** | Useful for building dashboards with text boxes, callouts, or branded elements. |
| **Impact** | Low |
| **Complexity** | M |
| **COM objects/methods** | `Shape.TextFrame2.TextRange.Text`, `Shape.Fill.ForeColor.RGB`, `Shape.Line`, `Worksheet.Shapes.AddTextbox()`, `Shape.Group()` |

---

#### L6. Borders Shorthand

| Attribute | Detail |
|---|---|
| **What's missing** | `Set-ExcelCellFormat` supports borders, but setting all-borders or box-borders requires specifying each edge. No convenience parameter for common border patterns. |
| **Why it matters** | "Put a border around this range" is a frequent agent request. |
| **Impact** | Low |
| **Complexity** | S |
| **COM objects/methods** | `Range.BorderAround(LineStyle, Weight, Color)`, `Range.Borders` collection |

---

## 4. Feature Comparison vs Full Excel COM Object Model

| Excel COM Capability | ExcelPOSH Coverage | Gap Priority |
|---|---|---|
| Workbook open/close/save/new | ✅ Complete | — |
| Worksheet CRUD + navigate | ✅ Complete | — |
| Range read/write/clear | ✅ Good | — |
| Range find/replace | ✅ Find + Update-ExcelValue | — |
| Named ranges (cell refs) | ✅ CRUD | — |
| Named ranges (formula-based) | ❌ Missing | Low |
| Tables (ListObject) CRUD | ✅ Good | — |
| Cell formatting (font, fill, border) | ✅ Core set | Low (border shorthand) |
| Number formatting | ✅ Complete | — |
| Alignment (H/V, wrap, indent) | ✅ Complete | — |
| Merge/unmerge cells | ❌ Missing | **High** |
| Named styles | ❌ Missing | Medium |
| Conditional formatting | ✅ Good | — |
| Data validation | ✅ Good | — |
| AutoFilter (basic) | ✅ Good | — |
| Advanced filter (unique/criteria) | ❌ Missing | **High** |
| Sort (single key) | ✅ Present | — |
| Sort (multi-key) | ❌ Missing | **High** |
| Insert/delete rows/columns | ❌ Missing | **High** |
| Remove duplicates | ❌ Missing | **High** |
| SpecialCells | ❌ Missing | **High** |
| TextToColumns | ❌ Missing | Medium |
| AutoFill / FillDown | ❌ Missing | Medium |
| Subtotals | ❌ Missing | Medium |
| WorksheetFunction | ❌ Missing | **High** |
| Application.Evaluate | ❌ Missing | **High** |
| Application.Run (macros) | ❌ Missing | **High** |
| ScreenUpdating / Calculation mode | ❌ Missing | **High** |
| StatusBar | ❌ Missing | Medium |
| Freeze panes | ✅ Complete | — |
| Sheet visibility | ✅ Complete | — |
| Grouping / outline | ✅ Good | — |
| Hyperlinks | ✅ Complete | — |
| Copy/move range | ✅ Good | — |
| Page setup | ✅ Good | — |
| Export to PDF | ✅ Present | — |
| Export to XPS | ❌ Missing (infra exists) | Medium |
| Direct printing | ❌ Missing | Medium |
| Images (add) | ✅ Basic | — |
| Shapes (text, format, group) | ❌ Missing | Low |
| Pivot tables (basic create/refresh) | ✅ Present | — |
| Pivot field customization | ❌ Missing | **High** |
| Pivot calculated fields | ❌ Missing | **High** |
| Charts (basic create/export) | ✅ Present | — |
| Chart series/axis/legend/labels | ❌ Missing | **High** |
| Import CSV/text | ✅ Good | — |
| Import ADO recordset | ❌ Missing | Medium |
| Sparklines | ✅ Good | — |
| Document properties | ✅ Good | — |
| Protection (sheet/workbook) | ✅ Good | — |
| Legacy comments | ✅ Good | — |
| Threaded comments | ❌ Missing | Low |
| Formula auditing (precedents) | ❌ Missing | Medium |
| Tab colors | ❌ Missing | Medium |
| Scenarios (what-if) | ❌ Missing | Low |
| Slicers | ❌ Missing | Low |

**Coverage summary:**  
- ✅ Covered: ~28 capability areas  
- ❌ Missing: ~25 capability areas (12 High, 10 Medium, 6 Low)

---

## 5. Test Infrastructure Assessment

### Current State

| Metric | Value |
|---|---|
| Test framework | Pester |
| Total tests | 561 passing |
| Test files | 18 (1 per public domain + 1 module-level) |
| Test type | **Parameter validation only** — mock-free, no COM calls |
| Behavioral tests | 3 (WorkbookOps file I/O only) |
| COM integration tests | **0** |
| Coverage tool | None configured |

### Gaps

1. **No COM functional tests.** Every function's happy path is untested. Regressions in COM interaction, constant mapping, or output formatting are invisible.

2. **No mock infrastructure.** Pester `Mock` is not used for `Connect-ExcelWorkbook` or the COM application object. Adding mocks would enable isolated behavioral testing without Excel installed.

3. **No error-path tests.** No tests verify behavior when: workbook doesn't exist, sheet name is wrong, range is invalid, COM connection is dead, or Excel is not installed.

4. **No output-schema tests.** Functions return `@{status='ok'; ...}` hashtables, but no test asserts the presence of required keys or value types. `-AsJson` output is untested.

5. **No performance/scale tests.** No benchmarks for large workbooks (100K+ rows), many sheets, or bulk operations.

6. **Missing: Pester tags and categories.** Tests are not tagged (`-Tag 'Unit'`, `'Integration'`, `'COM'`), so there's no way to run subsets.

### Recommendations

| Action | Priority | Effort |
|---|---|---|
| Add `Mock Connect-ExcelWorkbook` returning a fake COM tree for each domain | High | M |
| Add output-schema assertions to all existing test files | High | S |
| Add `-Tag` annotations (`Unit`, `Integration`) | Medium | S |
| Create a `COM.Integration.Tests.ps1` that runs against a real Excel instance with a test workbook | Medium | L |
| Add error-path tests (bad paths, missing sheets, dead COM) | Medium | M |
| Configure Pester code coverage reporting | Low | S |

---

## 6. Recommendations Summary

### Phase 1 — Immediate (12 functions, ~2–3 days)

These are all **Small complexity** and unlock the most common agent workflows:

| # | Function(s) | Addresses Gap |
|---|---|---|
| 1 | `Add-ExcelRow`, `Remove-ExcelRow`, `Add-ExcelColumn`, `Remove-ExcelColumn` | H1 — Insert/Delete |
| 2 | `Remove-ExcelDuplicates` | H2 — Deduplication |
| 3 | `Set-ExcelPerformanceMode` | H4 — ScreenUpdating/Calc |
| 4 | `Invoke-ExcelCalculate` | H4 — Force recalc |
| 5 | `Merge-ExcelRange`, `Split-ExcelRange` | H5 — Merge/Unmerge |
| 6 | `Get-ExcelSpecialCells` | H7 — Cell type selection |
| 7 | `Invoke-ExcelMacro` | H12 — Run VBA macros |

### Phase 2 — Core Enhancements (6 functions + 1 refactor, ~3–4 days)

| # | Function(s) | Addresses Gap |
|---|---|---|
| 1 | `Invoke-ExcelFunction` | H3 — WorksheetFunction |
| 2 | `Invoke-ExcelEvaluate` | H6 — Formula evaluation |
| 3 | Refactor `Sort-ExcelRange` for multi-key | H8 — Multi-key sort |
| 4 | `Invoke-ExcelAdvancedFilter` | H9 — Advanced filter |
| 5 | `Set-ExcelChartSeries`, `Set-ExcelChartAxis` | H10 — Chart detail |
| 6 | `Set-ExcelPivotField`, `Add-ExcelPivotCalculatedField` | H11 — Pivot detail |

### Phase 3 — Medium Priority Polish (~2 days)

All Medium-priority items: TextToColumns, AutoFill, Subtotals, Tab color, formula auditing, named styles, XPS export, PrintOut, StatusBar, CopyFromRecordset.

### Phase 4 — Test Infrastructure (ongoing)

Add mocked behavioral tests and output-schema assertions in parallel with feature development. Target: every new function ships with ≥3 behavioral tests (happy path, bad input, edge case).

### Manifest Discrepancy

The module manifest declares **84** functions, but only **80** appear in the `FunctionsToExport` list and source files. Audit for the 4 missing entries — they may be `ConvertTo-RGBColor` and the private helpers incorrectly counted, or planned functions that were never implemented. Update the manifest description to match actual count.
