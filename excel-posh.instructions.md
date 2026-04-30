---
description: "Use when the user mentions Excel workbooks (.xlsx/.xlsm/.xlsb/.xls), Excel-POSH, ExcelPOSH, PowerShell COM automation for Excel, or asks to read/write/format an Excel workbook."
---
# ExcelPOSH Module

When working with Microsoft Excel workbooks, import the PowerShell module:

```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1" -Force
```

This module provides 104 PowerShell functions for full Excel workbook automation via COM. Use `-AsJson` on any function for structured output. The `@excel-dev` agent has the complete function reference. New in v3.0.0: Structural operations (add/remove rows/columns), Calculation controls (performance mode, calculate, evaluate), advanced filter/duplicates, range merge/split, special cells, macro invocation, chart/pivot enhancements, and more.
