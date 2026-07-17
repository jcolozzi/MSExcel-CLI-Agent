# Excel Workbook Grapher — Design Notes

Port of the AccessPOSH dependency grapher (`Export-AccessGraph` / `Get-AccessGraphQuery`)
to Excel workbooks, shipped inside the **ExcelPOSH** module. Produces `graph.json` plus a
self-contained interactive `index.html` (vis.js) viewer.

Two graph layers in **one viewer** with a **Structure ⇄ Data** toggle:

1. **Structure graph** — workbook objects and how they reference each other (the analogue of
   Access tables/queries/forms/reports/macros/modules).
2. **Data graph** — the *actual data* and its relationships (an ER / lineage view) built from
   the Data Model, lookup formulas, and value-overlap inference.

---

## 1. Decisions (from requirements interview)

| # | Question | Decision |
|---|----------|----------|
| 1 | Structure + Data: one viewer or two? | **One viewer** with a Structure/Data layer toggle. |
| 2 | Data-relationship discovery methods | **All four**: Data Model FKs, lookup-formula inference, value-overlap FK inference, primary-key detection. |
| 3 | Formula edge granularity | **Both** aggregate (sheet/table/name) and per-cell, via `-FormulaMode Aggregate\|Cell\|Both\|None` (default `Aggregate`). |
| 4 | Test workbook domain | **Northwind-style sales** (Customers, Categories, Products, Employees, Orders, OrderDetails). |
| 5 | Packaging | ExcelPOSH module functions **`Export-ExcelGraph`**, **`Get-ExcelGraphQuery`**, **`Import-ExcelGraph`** (mirrors AccessPOSH). |
| 6 | VBA analysis | **Included** behind `-DisableVbaHeuristics`; requires *Trust access to the VBA project object model*. Degrades to a warning if untrusted. |

**Scope (v1)**
- IN: single workbook, both graph layers, HTML viewer, sample workbook, query API, Pester tests.
- OUT: recursing into *external* linked workbooks (shown as `external` nodes only), live refresh,
  non-Windows hosts.
- Data Model reads degrade gracefully (`status='unsupported'`) on editions without Power Pivot.

---

## 2. Excel object model → graph mapping

Reuses the ExcelPOSH COM session (`Connect-ExcelWorkbook`) and output helpers
(`Format-ExcelOutput`, `ConvertTo-ExcelSafeValue`).

| Excel object (COM) | Node group | Notes |
|--------------------|-----------|-------|
| `Workbook` | `workbook` | Root node. |
| `Worksheet` (`Workbook.Worksheets`) | `sheet` | Container for tables/charts/pivots/names. |
| `ListObject` (`Worksheet.ListObjects`) | `table` | Structured table + columns. |
| `Name` (`Workbook.Names`, sheet-scoped names) | `name` | Defined name / named range; has `RefersTo`. |
| `PivotTable` (`Worksheet.PivotTables()`) | `pivot` | Source via `PivotCache()` / `SourceData`. |
| `ChartObject.Chart` (`Worksheet.ChartObjects()`) + chart sheets (`Workbook.Charts`) | `chart` | Series source ranges from `SeriesCollection().Formula`. |
| `WorkbookConnection` (`Workbook.Connections`) | `connection` | OLEDB/ODBC/Text/Web/Model/Query; feeds caches/tables. |
| `WorkbookQuery` (`Workbook.Queries`) | `query` | Power Query M (`.Formula`). |
| `ModelTable` (`Workbook.Model.ModelTables`) | `modeltable` | Data Model table + `ModelTableColumns`. |
| `ModelMeasure` (`Workbook.Model.ModelMeasures`) | `measure` | DAX measure, associated table. |
| `Slicer`/`Timeline` (`Workbook.SlicerCaches`) | `slicer` | Filters pivots/tables (`SlicerCacheType` 1=slicer, 2=timeline). |
| `VBComponent` (`Workbook.VBProject.VBComponents`) | `module` | Std/Class/Form/Document VBA modules. |
| table/model column | `column` | Data-graph field node; carries PK/FK flags. |
| external workbook link | `external` | Cross-workbook reference target (not recursed). |

**COM gotchas (verified against ExcelPOSH + repo notes)**
- `Worksheet.PivotTables()`, `Workbook.PivotCaches()`, `Worksheet.ChartObjects()`,
  `Chart.SeriesCollection()`, `Worksheet.Scenarios()` are **parameterized properties** — call
  **with parentheses** or PowerShell returns a `PSMethod` and iteration silently yields nothing.
- Resolve the target workbook by `FullName` (not `ActiveWorkbook`) — `Connect-ExcelWorkbook`
  does not re-activate a workbook that was already open.
- `Workbook.Model` and `Workbook.VBProject` can throw; wrap in `try/catch` → warning, never fail
  the whole export.

---

## 3. Node & edge schema (`graph.json`)

Mirrors the AccessPOSH contract, plus a `layer` field for the viewer toggle.

```jsonc
node = { id, label, group, title, meta{}, layer }   // layer: "structure" | "data" | "both"
edge = { id, from, to, label, kind, arrows, title, meta{}, layer }
```

- `id` = `"<group>:<name>"` (e.g. `table:Orders`, `sheet:Data`, `column:Orders.CustomerID`).
- Node `layer`: `table` and `modeltable` = **both** (anchor both views); `column` = **data**;
  everything else = **structure**.
- Edge `layer`: structure kinds = `structure`; data kinds = `data`.

**Structure edge kinds**

| kind | from → to | source |
|------|-----------|--------|
| `contains` | sheet → table/chart/pivot/name | object ownership |
| `formula-ref` | sheet/table/name (or cell) → sheet/table/name | formula precedents |
| `chart-source` | chart → table/sheet range | `SERIES()` formula |
| `pivot-source` | pivot → table/sheet/connection/modeltable | `PivotCache`/`SourceData` |
| `connection-feeds` | connection → pivot/table/modeltable | connection ranges/model |
| `query-loads` | query → table/connection/modeltable | Power Query load target |
| `slicer-filters` | slicer → pivot/table | slicer cache links |
| `name-refersto` | name → sheet range/table | `RefersTo` |
| `code-ref` | module → sheet/name/table (+ module→module calls) | VBA code heuristics |

**Data edge kinds**

| kind | from → to | method | meta |
|------|-----------|--------|------|
| `column-owner` | table/modeltable → column | schema | — |
| `datamodel-fk` | FK column → PK column | `Model.ModelRelationships` | `active` |
| `lookup-fk` | table → table | VLOOKUP/XLOOKUP/INDEX-MATCH parse | `formula`, `via` |
| `inferred-fk` | column → column | value-overlap | `confidence`, `overlap`, `nameMatch` |

---

## 4. Data-relationship discovery (Level 2 algorithms)

1. **Explicit Data Model FKs** — iterate `Workbook.Model.ModelRelationships`; each exposes
   `ForeignKeyTable`, `ForeignKeyColumn`, `PrimaryKeyTable`, `PrimaryKeyColumn`, `Active`.
   Emit `datamodel-fk` (highest confidence = explicit).

2. **Lookup-formula inference** — regex-scan table-column formulas and worksheet formulas for
   `VLOOKUP`, `XLOOKUP`, `HLOOKUP`, `LOOKUP`, `INDEX`/`MATCH`. Resolve the lookup array argument
   to a table or range → owning table `lookup-fk`→ target table. Confidence = high (author intent).

3. **Value-overlap FK inference** — for each table column, sample distinct non-blank values
   (capped by `-MaxOverlapRows`, default 5000). For a candidate child column vs. a parent
   **PK** column, compute `overlap = |distinct(child) ∩ distinct(parent)| / |distinct(child)|`.
   If `overlap ≥ -FkOverlapThreshold` (default 0.85) emit `inferred-fk` with `confidence`
   (boosted when column names match, e.g. `CustomerID` ↔ `CustomerID`).

4. **Primary-key detection** — a column whose sampled values are **all non-blank and unique**
   is flagged `isPrimaryKey` on its `column` node (drives report + parent-key candidates for #3).

---

## 5. Viewer (`excel-graph-viewer.html`)

Adapted from `access-graph-viewer.html` (vis.js). Changes:
- Reskin `groupStyles` for the Excel node groups (workbook, sheet, table, name, pivot, chart,
  connection, query, modeltable, measure, slicer, module, column, external).
- **Structure ⇄ Data** toggle filtering nodes/edges by `layer`.
- Excel report panels: broken references (`#REF!`), external-link inventory, volatile formulas,
  high fan-in objects, circular references, tables without relationships, inferred-FK candidates
  (with confidence), unused named ranges, duplicate Power Query M, orphan sheets.
- Retained from Access viewer: search, group filters, dark mode, pinned popups, embedded-data
  bootstrap (`<!-- EMBED_GRAPH_DATA -->` → `index.html`).

Browser-validation reminders (from prior sessions): defeat `file://` caching with
`?v=Date.now()`; `new vis.Network` wipes `#network` innerHTML (build popups dynamically);
`network.emit('click')` does not fire real handlers — compute `canvasToDOM` coords + `page.mouse.click`.

---

## 6. Query API

- `Import-ExcelGraph -Path graph.json` → in-memory graph object (adjacency maps).
- `Get-ExcelGraphQuery -Action neighbors|impact|path|orphans|summary` — ported from
  `Get-AccessGraphQuery` (BFS neighbors/impact, undirected shortest path, orphan detection,
  summary stats). Accepts `-WorkbookPath` (auto-locates `excel-graph-out/graph.json`) or `-GraphPath`.

---

## 7. Test workbook (`ExcelGraph-Sample.xlsm`)

Built repeatably by `misc/Build-ExcelGraphSample.ps1` using ExcelPOSH functions. Contents,
chosen to exercise every edge type:

- **Tables** (ListObjects, one per sheet): Customers, Categories, Products, Employees, Orders,
  OrderDetails (~6–40 rows each, typed).
- **Lookup formulas**: OrderDetails uses `XLOOKUP` into Products (name/price); Orders uses
  `VLOOKUP` into Customers and Employees. → `lookup-fk` edges.
- **Named ranges**: e.g. `TaxRate`, a category lookup range. → `name-refersto`.
- **Data Model**: all six tables added; relationships
  Orders[CustomerID]→Customers, Orders[EmployeeID]→Employees,
  OrderDetails[OrderID]→Orders, OrderDetails[ProductID]→Products,
  Products[CategoryID]→Categories. → `datamodel-fk` edges.
- **PivotTable**: sales by category off the Data Model. → `pivot-source`.
- **Chart**: from a summary range. → `chart-source`.
- **VBA macro** (hence `.xlsm`): a small `Sub` referencing a sheet/range. → `code-ref`.

Foreign keys are *also* discoverable by value overlap (Orders.CustomerID ⊆ Customers.CustomerID),
so the sample validates methods 1–4 independently.

---

## 8. Files

**New**
- `misc/excel-grapher-design.md` (this file)
- `ExcelPOSH/Private/ExcelGraphHelpers.ps1`
- `ExcelPOSH/Public/GraphOps.ps1` (`Export-ExcelGraph`)
- `ExcelPOSH/Public/GraphQueryOps.ps1` (`Import-ExcelGraph`, `Get-ExcelGraphQuery`)
- `ExcelPOSH/Resources/excel-graph-viewer.html`
- `misc/Build-ExcelGraphSample.ps1` → `misc/ExcelGraph-Sample.xlsm`
- `Tests/ExcelGraph.Tests.ps1`

**Modified**
- `ExcelPOSH/ExcelPOSH.psd1` (export new functions, version + description bump)
- `Tests/ExcelPOSH.Module.Tests.ps1` (function count + version)

---

## 9. Verification

1. `Import-Module ExcelPOSH -Force`; `Get-Command Export-ExcelGraph, Get-ExcelGraphQuery, Import-ExcelGraph`.
2. `Build-ExcelGraphSample.ps1` → confirm 6 tables + 5 model relationships + pivot + chart.
3. `Export-ExcelGraph` on the sample → assert `graph.json` has ≥6 `table` nodes, ≥5 `datamodel-fk`
   edges, `lookup-fk` edges present, `inferred-fk` edges present.
4. Open `index.html`; toggle Structure/Data; groups render; reports populate; zero `pageerror`.
5. `Get-ExcelGraphQuery -Action summary|neighbors|impact`; `Invoke-Pester Tests/ExcelGraph.Tests.ps1`.
