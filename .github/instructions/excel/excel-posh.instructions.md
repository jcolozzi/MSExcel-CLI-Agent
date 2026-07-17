---
description: "Use when the user mentions Excel workbooks (.xlsx/.xlsm/.xlsb/.xls), Excel-POSH, ExcelPOSH, PowerShell COM automation for Excel, or asks to read/write/format an Excel workbook."
---
# ExcelPOSH Module

When working with Microsoft Excel workbooks, import the PowerShell module:

```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1" -Force
```

This module provides 147 PowerShell functions for full Excel workbook automation via COM. Use `-AsJson` on any function for structured output. The `@excel-dev` agent has the complete function reference. New in v4.1.0: **dependency & data-relationship graphing** (`Export-ExcelGraph`, `Get-ExcelGraphQuery`, `Import-ExcelGraph`). New in v4.0.0: Power Query (Get & Transform), data-connection refresh/add/remove, Data Model / Power Pivot (DAX measures & relationships), slicers & timelines, dynamic-array formulas (`Set-ExcelFormula2`), linked data types (Stocks/Geography), named styles, threaded comments, Goal Seek, what-if scenarios, subtotals, AutoFill, Text-to-Columns, ADO recordset import, status bar, sheet tab color, and direct print / XPS export.

## Dependency & data graph

Three functions map how everything in a workbook connects — use them for any question about
**structure, relationships, dependencies, impact, lineage, or orphaned/unused objects**:

- `Export-ExcelGraph -WorkbookPath $wb` — scans the workbook and writes `graph.json` **and** an
  interactive `index.html` viewer to `<workbook-folder>\excel-graph-out\` (override with `-OutDir`).
  Builds a **structure** layer (sheets, tables, named ranges, pivots, charts, connections, Power
  Query, Data Model, slicers, VBA) and a **data** layer (Data Model foreign keys, lookup-formula
  relationships, value-overlap inferred FKs, primary keys). `-FormulaMode Aggregate|Cell|Both|None`.
- `Get-ExcelGraphQuery -WorkbookPath $wb -Action summary|neighbors|impact|path|orphans` — queries
  the graph. It **auto-locates** `<workbook-folder>\excel-graph-out\graph.json`, so you never re-scan.
- `Import-ExcelGraph -WorkbookPath $wb` — loads the full graph object for custom inspection.

**How to know whether a graph exists:** just call `Get-ExcelGraphQuery -WorkbookPath $wb -Action summary`.
Because it auto-locates `excel-graph-out\graph.json` beside the workbook, an existing graph answers
instantly (no Excel/COM needed); if it reports the graph is missing, run `Export-ExcelGraph` once and
then query. Build the graph once, query it many times; only re-export after the workbook's structure
changes. VBA code edges require *Trust access to the VBA project object model* (`enable_vba_trust.ps1`).
