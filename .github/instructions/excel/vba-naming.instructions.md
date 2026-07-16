---
description: Guardrails for Excel/VBA naming
applyTo: "**/*.{bas,cls,frm,vba}"
---

# Excel/VBA Naming Guardrails

## Core Rules

- **Do not** use reserved words for identifiers (variables, procedures, modules, worksheet/workbook objects, named ranges, table names, or column names).
- Reserved-word matching is case-insensitive.
- Names must start with a letter and avoid spaces/special characters.
- Use `Option Explicit` in every module.
- Prefer CamelCase and Leszynski/Reddick-style prefixes.
- Rename collisions; do not rely on qualification as the primary fix.

## VBA Built-in Collision Hazards

Avoid using built-in keyword/function names as identifiers, including:

`Array`, `Date`, `Time`, `Now`, `Left`, `Right`, `Mid`, `Replace`, `Filter`, `Error`, `String`, `Sub`, `Function`, `If`, `For`, `Next`, `Do`, `Loop`, `Select`, `Case`, `Set`, `Dim`, `Type`, `Variant`, `Object`, `Null`, `Empty`, `True`, `False`

## Excel Object Model Collision Hazards

Never use common Excel object/property/method names as identifiers:

`Cells`, `Cell`, `Range`, `Rows`, `Columns`, `Row`, `Column`, `Sheets`, `Sheet`, `Workbook`, `Workbooks`, `Worksheet`, `Worksheets`, `Name`, `Names`, `Value`, `Formula`, `Address`, `Count`, `Index`, `Item`, `Chart`, `Charts`, `Filter`, `Table`, `PivotTable`

## Worksheet Function Name Hazards

Avoid worksheet function names as identifiers:

`Sum`, `Average`, `Min`, `Max`, `Match`, `Index`, `Offset`, `VLookup`, `HLookup`, `CountIf`, `Trim`, `Upper`, `Lower`, `Text`, `Today`, `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`

## Named Ranges and Table Naming

- Do not use names that look like references (`A1`, `B2`, `R1C1`).
- Do not reuse built-in function names.
- Prefix named ranges with `nm` and ListObjects with `tbl`.

## Sheet CodeName vs Tab Name

- Prefer Worksheet `CodeName` for VBA references because it is stable in code.
- Treat worksheet tab `Name` as user-facing and mutable.
- Apply `ws` prefix to CodeNames where possible: `wsData`, `wsSummary`, `wsConfig`.

## Prefix Guidance

- Excel objects: `wb` (Workbook), `ws` (Worksheet), `rng` (Range), `cht` (Chart), `tbl` (ListObject), `pt` (PivotTable), `pf` (PivotField), `shp` (Shape), `nm` (Named Range).
- Variables: `strName`, `lngRow`, `dblRate`, `blnFound`, `dtStart`, `varResult`, `objSheet`, `arrData`.
- Modules/classes/userforms: `modUtilities`, `clsLogger`, `frmOptions`.

## Procedure Naming

- Use PascalCase and verb-first names for procedures: `CalculateTotal`, `RefreshPivot`, `LoadWorksheetData`.
- Keep event handlers in `ObjectName_EventName` format: `btnRun_Click`, `wsData_Change`.

> References: VBA naming rules, avoiding naming conflicts, VBA keywords, Excel Worksheet.CodeName documentation, and Excel custom function naming constraints on Microsoft Learn.
