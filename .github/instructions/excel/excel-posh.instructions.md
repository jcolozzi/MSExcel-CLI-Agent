---
description: "Use when the user mentions Excel workbooks (.xlsx/.xlsm/.xlsb/.xls), Excel-POSH, ExcelPOSH, PowerShell COM automation for Excel, or asks to read/write/format an Excel workbook."
---
# ExcelPOSH Module

When working with Microsoft Excel workbooks, import the PowerShell module:

```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1" -Force
```

This module provides 144 PowerShell functions for full Excel workbook automation via COM. Use `-AsJson` on any function for structured output. The `@excel-dev` agent has the complete function reference. New in v4.0.0: Power Query (Get & Transform), data-connection refresh/add/remove, Data Model / Power Pivot (DAX measures & relationships), slicers & timelines, dynamic-array formulas (`Set-ExcelFormula2`), linked data types (Stocks/Geography), named styles, threaded comments, Goal Seek, what-if scenarios, subtotals, AutoFill, Text-to-Columns, ADO recordset import, status bar, sheet tab color, and direct print / XPS export.
