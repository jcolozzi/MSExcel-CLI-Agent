---
name: excel-vba-reserved-words
description: >-
  Detect and avoid Microsoft Excel / VBA reserved words in identifiers.
  Use when naming variables, procedures, modules, worksheet/workbook objects,
  named ranges, tables, or formula-facing names.
  Guide learners to safe alternatives and consistent conventions.
license: CC-BY-4.0
---

# Excel/VBA Reserved Words – Naming Safety Skill

## Purpose
Help developers avoid case-insensitive collisions with:
- **VBA keywords & operators** (`Dim`, `If`, `Select`, `Function`, `ByVal`, `And`, etc.)
- **Built-in VBA function names** (`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, etc.)
- **Excel object model names** (`Range`, `Cells`, `Rows`, `Columns`, `Worksheet`, `Workbook`, `Name`, `Value`, etc.)
- **Worksheet function names** (`SUM`, `IF`, `INDEX`, `MATCH`, `VLOOKUP`, etc.)
- **Defined name restrictions** (cell-reference-like names such as `A1`, `R1C1`, or names that collide with function names)

> Excel/VBA naming is **case-insensitive**. Reusing reserved or built-in names as identifiers often causes compile/runtime bugs and ambiguous references.

## Procedure
1. **Scan identifiers** in the current context (variables, procedure names, worksheet/workbook codenames, named ranges, table names/columns, and formula-facing names).
2. **Flag exact case-insensitive matches** to reserved words or built-ins and report each occurrence with location.
3. **Flag Excel object-model collisions** and worksheet-function collisions.
4. **Suggest safe replacements** using descriptive names and prefixes:
   - `Range` -> `rngData`
   - `Name` -> `strName` or `nmCustomer`
   - `Date` -> `dtStart` or `reportDate`
   - `Count` -> `lngCount` or `itemCount`
   - `Sheet` -> `wsData`
5. **Apply conventions**:
   - Use CamelCase, no spaces/special characters.
   - Prefer Leszynski/Reddick-style prefixes (`strName`, `lngCount`, `dtStart`, `wsData`, `rngInput`, `tblOrders`).
6. **For defined names** (named ranges/tables), reject names that look like references (`A1`, `B2`, `R1C1`) or collide with built-in function names.
7. **(Optional) Run a repository scan** when requested to list offenders and propose bulk renames.

## Common Offenders (teach by example)
- VBA/built-in identifiers: `date`, `time`, `now`, `left`, `right`, `mid`, `len`, `replace`, `filter`, `array`
- Excel object names: `range`, `cell`, `cells`, `row`, `rows`, `column`, `columns`, `sheet`, `sheets`, `worksheet`, `workbook`, `name`, `value`, `formula`, `address`, `count`, `item`, `index`
- Worksheet function names used as identifiers: `sum`, `average`, `min`, `max`, `if`, `match`, `index`, `lookup`, `vlookup`, `hlookup`, `trim`, `text`, `today`
- Invalid defined names: `A1`, `R1C1`, `SUM`, `IF`

## Naming Recommendations
- Prefer descriptive names: `salesDate`, `isEligible`, `totalAmount`, `wsSummary`, `rngCriteria`, `tblInvoiceLines`.
- Use prefixes to communicate type/object intent:
  - `ws` worksheet, `wb` workbook, `rng` range, `tbl` table/listobject, `pt` pivottable
  - `str` string, `lng` long, `dbl` double, `bln` boolean, `dt` date/time
- Prefer `Worksheet.CodeName` references in VBA and keep tab names user-facing.
- Avoid generic names like `Data`, `Info`, `Temp` when they collide with built-ins or reduce clarity.

## Example Prompts (to trigger this skill)
- "Scan this Excel VBA module for reserved-word collisions."
- "Suggest safer names for my worksheet variables and named ranges."
- "Check if my table and defined names conflict with functions or references."

## References
- VBA reserved words and language reference: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-reserved-words
- VBA naming rules: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules
- Avoiding naming conflicts: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/avoiding-naming-conflicts
- Excel object model reference: https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model
- Worksheet.CodeName: https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.codename
- Worksheet.Name: https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.name
- Excel defined names rules: https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-a1c4e74e-fd5d-4b67-9b1d-3b6b907f8cb8
