# MSExcel-CLI-Agent

> **Automate Microsoft Excel from plain English inside VS Code — no MCP server, no Python, no extra processes.**

![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue?logo=windows)
![PowerShell: 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![VS Code](https://img.shields.io/badge/VS%20Code-GitHub%20Copilot%20Chat-blueviolet?logo=visual-studio-code)
![Functions: 144](https://img.shields.io/badge/functions-144-brightgreen)
![Module Version](https://img.shields.io/badge/version-1.0.0-orange)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

## What is this?

**MSExcel-CLI-Agent** is a VS Code agent (powered by GitHub Copilot Chat) that lets you talk to Microsoft Excel in plain language. You describe what you want and the agent translates it into PowerShell commands that manipulate your `.xlsx` / `.xlsm` workbooks live via COM — no manual VBA editing required.

```text
You:   "Create a table in Sheet1 with columns ID, Name, Email and add 100 rows of sample data"
Agent: → New-ExcelTable → confirms success
```

The **ExcelPOSH** module (included) is a comprehensive PowerShell interface to the Excel Object Model, providing **144 public functions** covering workbooks, worksheets, tables, formatting, filtering, pivot tables, charts, Power Query, the Data Model, slicers, and data manipulation.

## How it works

```text
VS Code Copilot Chat (agent mode)
        │
        ▼
  excel-dev agent (.md instructions)
        │  describes which PowerShell command to run
        ▼
  ExcelPOSH module  (imported in the VS Code terminal)
        │  COM calls via Excel Object Model
        ▼
  Microsoft Excel (.xlsx / .xlsm)
```

- **No separate server** — the module runs directly in the VS Code integrated terminal.
- **No Python / Node** — pure PowerShell 5.1+ on Windows.
- **Full COM access** — everything you can do from VBA, you can do from the agent.
- **-WhatIf / -Confirm** — all state-changing functions support PowerShell's standard risk-mitigation flags.
- **Pester tests** — 18 test files cover every public command group.

## Prerequisites

| Requirement | Details |
| --- | --- |
| **OS** | Windows 10 / 11 (COM automation is Windows-only) |
| **Microsoft Excel** | Excel 2016, 2019, 2021, or Microsoft 365 (desktop) |
| **PowerShell** | 5.1 (Windows PowerShell) **or** PowerShell 7+ |
| **VS Code** | Latest stable, with the **GitHub Copilot Chat** extension |
| **Copilot** | An active GitHub Copilot subscription |

## Setup

### 1 — Clone the repo

```powershell
git clone https://github.com/jcolozzi/MSExcel-CLI-Agent.git
```

### 2 — Import the module in your VS Code terminal

Open the integrated terminal in VS Code and import the module once per session (or add to your `$PROFILE`):

```powershell
Import-Module "C:\path\to\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1"
```

Verify it loaded:

```powershell
Get-Command -Module ExcelPOSH | Measure-Object  # should show 84
```

### 3 — Install the agent instructions

Choose **one** of the following:

### Option A — User-level (available in every workspace)

Copy both `.md` files from the repo root to:

```text
C:\Users\%USERNAME%\AppData\Roaming\Code\User\prompts\
```

### Option B — Workspace-level (scoped to this project)

Copy both `.md` files into a `.github\agents\` folder in your workspace root. VS Code automatically detects any `.md` files in that folder as custom agents.

> **Note:** VS Code detects any `.md` files in the `.github/agents/` folder of your workspace as custom agents.

### 4 — Update the module path inside the agent files

Open each `.md` agent file and replace the placeholder path with the actual path to `ExcelPOSH.psd1` on your machine:

```text
# Before
Import-Module "C:\path\to\ExcelPOSH\ExcelPOSH.psd1"

# After (example)
Import-Module "C:\Projects\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1"
```

### 5 — Select the agent and start prompting

In VS Code Copilot Chat, click the agent picker and choose **excel-dev**. Open (or have the agent open) an `.xlsx` or `.xlsm` file, then start describing what you want.

## Usage examples

| Prompt | Functions called |
| --- | --- |
| "Create a new Excel workbook and add three sheets" | `New-ExcelWorkbook`, `New-ExcelWorksheet` |
| "List all tables in the current workbook" | `Get-ExcelTable` |
| "Create a table called Customers with columns ID, Name, Email" | `New-ExcelTable` |
| "Add 100 rows of sample data to the Customers table" | `Add-ExcelTableRow` |
| "Apply bold and blue font to the header row" | `Set-ExcelCellFormat` |
| "Freeze the first row and first column" | `Set-ExcelFreezePane` |
| "Filter the Sales column to show only values > 1000" | `Set-ExcelAutoFilter` |
| "Sort the data by Date descending" | `Sort-ExcelRange` |
| "Create a pivot table from the data" | `New-ExcelPivotTable` |
| "Export the current sheet to PDF" | `Export-ExcelToPdf` |
| "What would happen if I created a new worksheet? (dry run)" | `New-ExcelWorksheet -WhatIf` |

## Project structure

```text
MSExcel-agent/
├── ExcelPOSH/              # PowerShell module (the engine)
│   ├── ExcelPOSH.psd1      # Module manifest (v1.0.0, PS 5.1+, Desktop + Core)
│   ├── ExcelPOSH.psm1      # Module loader
│   ├── Public/             # 18 files — one per command category
│   │   ├── WorkbookOps.ps1
│   │   ├── WorksheetOps.ps1
│   │   ├── TableOps.ps1
│   │   ├── FormattingOps.ps1
│   │   ├── MetadataOps.ps1
│   │   ├── FilterSortOps.ps1
│   │   ├── ConditionalFormatOps.ps1
│   │   ├── DataValidationOps.ps1
│   │   ├── ViewOps.ps1
│   │   ├── HyperlinkOps.ps1
│   │   ├── ClipboardOps.ps1
│   │   ├── PrintOps.ps1
│   │   ├── ImageShapeOps.ps1
│   │   ├── PivotTableOps.ps1
│   │   ├── ChartOps.ps1
│   │   ├── ImportOps.ps1
│   │   └── SparklineOps.ps1
│   └── Private/            # Internal helpers (COM session, utilities, etc.)
├── Tests/                  # Pester test suite — 18 test files
│   ├── ExcelPOSH.Module.Tests.ps1
│   ├── WorkbookOps.Tests.ps1
│   ├── WorksheetOps.Tests.ps1
│   ├── TableOps.Tests.ps1
│   ├── FormattingOps.Tests.ps1
│   └── ...
├── excel-dev.md            # Agent instructions (the Copilot Chat prompt)
└── README.md
```

## Running the tests

```powershell
# From the repo root
Invoke-Pester .\Tests\ -Output Detailed
```

> Requires [Pester](https://github.com/pester/Pester) 5.x: `Install-Module Pester -MinimumVersion 5.0 -Force`

## Function reference

<details>
  <summary><strong>View all 144 public functions</strong></summary>

| Category | Functions |
| --- | --- |
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
| **Structural** (v3.0) | `Add-ExcelRow`, `Remove-ExcelRow`, `Add-ExcelColumn`, `Remove-ExcelColumn` |
| **Calculation** (v3.0) | `Set-ExcelPerformanceMode`, `Invoke-ExcelCalculate`, `Invoke-ExcelFunction`, `Invoke-ExcelEvaluate` |
| **Other v3.0** | `Remove-ExcelDuplicates`, `Invoke-ExcelAdvancedFilter`, `Merge-ExcelRange`, `Split-ExcelRange`, `Get-ExcelSpecialCells`, `Invoke-ExcelMacro`, `Set-ExcelChartSeries`, `Set-ExcelChartAxis`, `Set-ExcelChartLegend`, `Set-ExcelChartDataLabels`, `Set-ExcelPivotField`, `Add-ExcelPivotCalculatedField` |
| **Power Query** (v4.0) | `Get-ExcelPowerQuery`, `New-ExcelPowerQuery`, `Set-ExcelPowerQuery`, `Remove-ExcelPowerQuery`, `Update-ExcelPowerQuery`, `Import-ExcelPowerQueryToTable` |
| **Data Connections** (v4.0) | `Update-ExcelDataConnection`, `New-ExcelDataConnection`, `Remove-ExcelDataConnection` |
| **Data Model / Power Pivot** (v4.0) | `Get-ExcelDataModel`, `Add-ExcelModelMeasure`, `Remove-ExcelModelMeasure`, `Add-ExcelModelRelationship`, `Update-ExcelDataModel` |
| **Slicers & Timelines** (v4.0) | `New-ExcelSlicer`, `Get-ExcelSlicer`, `Set-ExcelSlicer`, `Remove-ExcelSlicer`, `New-ExcelTimeline`, `Set-ExcelTimelineRange` |
| **Styles** (v4.0) | `New-ExcelStyle`, `Set-ExcelRangeStyle`, `Get-ExcelStyle` |
| **Threaded Comments** (v4.0) | `Add-ExcelThreadedComment`, `Get-ExcelThreadedComment`, `Add-ExcelThreadedCommentReply`, `Remove-ExcelThreadedComment` |
| **What-If** (v4.0) | `Invoke-ExcelGoalSeek`, `Add-ExcelScenario`, `Get-ExcelScenario` |
| **Worksheet additions** (v4.0) | `Set-ExcelSheetTab`, `Invoke-ExcelAutoFill`, `Set-ExcelFormula2`, `Get-ExcelFormulaDependencies`, `Convert-ExcelToLinkedDataType` |
| **Data prep** (v4.0) | `Split-ExcelColumn`, `Import-ExcelRecordset`, `Add-ExcelSubtotal` |
| **Workbook / Print additions** (v4.0) | `Set-ExcelStatusBar`, `Send-ExcelPrint` (plus `Export-ExcelToPdf -Format XPS`) |

</details>

All state-changing functions support `-WhatIf` and `-Confirm` via PowerShell's standard `ShouldProcess` mechanism.

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you would like to change. Please include or update Pester tests for any new or modified functions.

## Credits
