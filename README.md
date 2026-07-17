# MSExcel-CLI-Agent

> **Automate Microsoft Excel from plain English inside VS Code — no MCP server, no Python, no extra processes.**

![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue?logo=windows)
![PowerShell: 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![VS Code](https://img.shields.io/badge/VS%20Code-GitHub%20Copilot%20Chat-blueviolet?logo=visual-studio-code)
![Functions: 147](https://img.shields.io/badge/functions-147-brightgreen)
![Module Version](https://img.shields.io/badge/version-1.0.0-orange)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

## What is this?

**MSExcel-CLI-Agent** is a VS Code agent (powered by GitHub Copilot Chat) that lets you talk to Microsoft Excel in plain language. You describe what you want and the agent translates it into PowerShell commands that manipulate your `.xlsx` / `.xlsm` workbooks live via COM — no manual VBA editing required.

```text
You:   "Create a table in Sheet1 with columns ID, Name, Email and add 100 rows of sample data"
Agent: → New-ExcelTable → confirms success
```

The **ExcelPOSH** module (included) is a comprehensive PowerShell interface to the Excel Object Model, providing **147 public functions** covering workbooks, worksheets, tables, formatting, filtering, pivot tables, charts, Power Query, the Data Model, slicers, data manipulation, and dependency & data-relationship graphing.

## How it works

```text
VS Code Copilot Chat (agent mode)
        │
        ▼
  excel-dev / excel-analysis agent (.github/agents/)
        │  describes which PowerShell command to run
        ▼
  ExcelPOSH module  (imported in the VS Code terminal)
        │  COM calls via Excel Object Model
        ▼
  Microsoft Excel (.xlsx / .xlsm)
```

- **Two agents** — `excel-dev` builds & edits workbooks; `excel-analysis` does read-first data analysis and reporting.
- **No separate server** — the module runs directly in the VS Code integrated terminal.
- **No Python / Node** — pure PowerShell 5.1+ on Windows.
- **Full COM access** — everything you can do from VBA, you can do from the agent.
- **-WhatIf / -Confirm** — all state-changing functions support PowerShell's standard risk-mitigation flags.
- **Dependency & data graph** — `Export-ExcelGraph` maps how everything in a workbook connects (structure *and* data relationships) and renders it as an interactive HTML graph.
- **Pester tests** — 26 test files cover every public command group.

## Agents

Two custom agents ship in `.github/agents/` (auto-detected by VS Code Copilot Chat when this folder is in your workspace):

| Agent | Use it for |
| --- | --- |
| **excel-dev** — *Excel Workbook Development Expert* | Building & editing workbooks: sheets, tables, formatting, formulas, charts, pivots, Power Query, the Data Model, slicers, imports/exports. |
| **excel-analysis** — *Excel Data Analysis Expert* | Read-first data analysis & reporting: profiling, data-quality auditing, statistics, aggregation, PivotTables/charts, reproducible findings. |

Both drive the same **ExcelPOSH** module — `excel-dev` writes, `excel-analysis` reads and summarizes.

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
Get-Command -Module ExcelPOSH | Measure-Object  # should show 147
```

### 3 — Use the bundled agents, instructions & skills (already in `.github/`)

This repo already ships everything under `.github/` — no copying required:

```text
.github/
├── agents/               excel-dev.agent.md, excel-analysis.agent.md
├── instructions/excel/   excel-posh + vba-naming guardrails (auto-applied)
└── skills/               excel-workbook-analysis, excel-workbook-planning,
                         excel-vba-reserved-words, prd
```

**Workspace-level (recommended):** open the cloned repo (or any workspace whose root contains this `.github/` folder) in VS Code. Copilot Chat auto-detects the agents, instructions, and skills — nothing to install.

**User-level (optional):** to make the agents available in *every* workspace, copy the files from `.github/agents/` (and, if you want them, `.github/instructions/` and `.github/skills/`) into:

```text
C:\Users\%USERNAME%\AppData\Roaming\Code\User\prompts\
```

### 4 — Point the agents at your ExcelPOSH path

The agent and instruction files import ExcelPOSH by absolute path. Update it to where you cloned the repo in each of:

- `.github/agents/excel-dev.agent.md`
- `.github/agents/excel-analysis.agent.md`
- `.github/instructions/excel/excel-posh.instructions.md`

```text
# Replace this line's path in each file
Import-Module "C:\path\to\MSExcel-agent\ExcelPOSH\ExcelPOSH.psd1"
```

### 5 — Select an agent and start prompting

In VS Code Copilot Chat, open the agent picker and choose **excel-dev** (build/edit) or **excel-analysis** (analyze/report). Open (or have the agent open) an `.xlsx` / `.xlsm` file, then describe what you want.

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
| "Graph this workbook's dependencies and relationships" | `Export-ExcelGraph` |
| "Profile this workbook and flag data-quality issues" *(excel-analysis)* | `Get-ExcelWorkbookInfo`, `Get-ExcelSpecialCells`, `Remove-ExcelDuplicates` |
| "Summarize sales by region with a PivotTable" *(excel-analysis)* | `New-ExcelPivotTable`, `Invoke-ExcelFunction` |
| "What would happen if I created a new worksheet? (dry run)" | `New-ExcelWorksheet -WhatIf` |

## Dependency & data graph

`Export-ExcelGraph` scans a workbook and builds an interactive graph of everything inside it and how it connects — writing `graph.json` plus a self-contained `index.html` viewer (vis.js). It captures two layers, toggled in the viewer:

- **Structure** — worksheets, tables, named ranges, PivotTables, charts, connections, Power Query, Data Model tables/measures, slicers, and VBA modules, plus the references between them (formulas, chart/pivot sources, connection feeds, query loads, slicer filters, VBA code references).
- **Data relationships** — the actual entity relationships: Data Model foreign keys, lookup-formula relationships (VLOOKUP/XLOOKUP/INDEX-MATCH), value-overlap inferred foreign keys, and primary-key detection.

```powershell
# Build graph.json + index.html
Export-ExcelGraph -WorkbookPath C:\Sales.xlsm -OutDir .\excel-graph-out -FormulaMode Both

# Query the graph without re-scanning
Get-ExcelGraphQuery -Action summary   -GraphPath .\excel-graph-out\graph.json
Get-ExcelGraphQuery -Action neighbors -GraphPath .\excel-graph-out\graph.json -Node table:Orders
Get-ExcelGraphQuery -Action impact    -GraphPath .\excel-graph-out\graph.json -Node table:Customers
```

The viewer has collapsible side panels, group & edge-kind filters, a **Structure / Data** toggle, light/dark mode, and report panels (inferred-FK candidates, tables without relationships, high fan-in objects, unused named ranges, and more). A worked example lives in [`misc/`](misc/) — `excel-grapher-design.md`, the `ExcelGraph-Sample.xlsm` sample workbook, and the generated `excel-graph-out/` graph.

> VBA module analysis requires *Trust access to the VBA project object model* (Excel Trust Center). Run `enable_vba_trust.ps1` (HKCU, no admin) to enable it.

## Project structure

```text
MSExcel-agent/
├── .github/                     # Copilot Chat customizations (auto-detected)
│   ├── agents/                  # excel-dev.agent.md, excel-analysis.agent.md
│   ├── instructions/excel/      # excel-posh + vba-naming guardrails
│   ├── skills/                  # workbook-analysis, workbook-planning, reserved-words, prd
│   └── hooks/                   # optional session hooks (auto-commit, logging)
├── ExcelPOSH/                   # PowerShell module (the engine)
│   ├── ExcelPOSH.psd1           # Module manifest (PS 5.1+, Desktop + Core)
│   ├── ExcelPOSH.psm1           # Module loader + COM constants
│   ├── Public/                  # 25 files — one per command category
│   │   ├── WorkbookOps.ps1, WorksheetOps.ps1, TableOps.ps1, FormattingOps.ps1
│   │   ├── FilterSortOps.ps1, PivotTableOps.ps1, ChartOps.ps1, ImportOps.ps1, ...
│   │   ├── PowerQueryOps.ps1, DataConnection.ps1, DataModelOps.ps1, SlicerOps.ps1
│   │   └── GraphOps.ps1, GraphQueryOps.ps1   # dependency + data-relationship graph
│   ├── Private/                 # Internal helpers (COM session, utilities, graph helpers)
│   └── Resources/               # excel-graph-viewer.html (embedded vis.js viewer)
├── Tests/                       # Pester suite — 26 test files (900+ tests)
├── misc/                        # Worked graph example (design doc, sample workbook, output)
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
  <summary><strong>View all 147 public functions</strong></summary>

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
| **Graph** (v4.1) | `Export-ExcelGraph`, `Import-ExcelGraph`, `Get-ExcelGraphQuery` |

</details>

All state-changing functions support `-WhatIf` and `-Confirm` via PowerShell's standard `ShouldProcess` mechanism.

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you would like to change. Please include or update Pester tests for any new or modified functions.

## Credits
