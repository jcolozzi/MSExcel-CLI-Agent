# ExcelPOSH v3.0.0 — Feature Expansion Implementation Plan

**Module:** ExcelPOSH  
**Current Version:** 2.0.0 (84 exported / 80 confirmed in source)  
**Target Version:** 3.0.0  
**Date:** 2026-04-14  
**Basis:** ExcelPOSH Gap Analysis (H1–H12, M1–M10, L1–L6)

---

## Table of Contents

1. [Architecture Conventions](#1-architecture-conventions)
2. [Phase 1 — Structural & Performance (12 functions)](#2-phase-1--structural--performance)
3. [Phase 2 — Calculation Engine & Charting (8 functions)](#3-phase-2--calculation-engine--charting)
4. [Phase 3 — Medium Priority (10 functions)](#4-phase-3--medium-priority)
5. [New Constants](#5-new-constants)
6. [Manifest Updates](#6-manifest-updates)
7. [File Summary](#7-file-summary)
8. [Updated Function Count](#8-updated-function-count)
9. [Implementation Order & Dependencies](#9-implementation-order--dependencies)

---

## 1. Architecture Conventions

Every new function **must** follow these patterns exactly. Reviewers should reject PRs that deviate.

| Rule | Detail |
|---|---|
| **Signature** | `[CmdletBinding()] param( [Parameter(Mandatory)][string]$WorkbookPath, [Parameter(Mandatory)][string]$SheetName, ... [switch]$AsJson )` |
| **Session** | First statement: `$app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath` |
| **Worksheet** | `$ws = $app.ActiveWorkbook.Worksheets.Item($SheetName)` (omit when function is workbook-scoped) |
| **Range** | `$rng = $ws.Range($Range)` |
| **Return** | Build `$result = @{ status='ok'; ... }`, then `Format-ExcelOutput -Data $result -AsJson:$AsJson` |
| **Constants** | Enum-like COM values → `$script:XL_*` hashtables in `ExcelPOSH.psm1`; look up with `[int]$script:XL_NAME[$key]` |
| **ValidateSet** | Every parameter that maps to an enum constant gets a `[ValidateSet()]` attribute |
| **File placement** | One `Public/<Domain>Ops.ps1` per domain; one `Tests/<Domain>Ops.Tests.ps1` per public file |
| **Tests** | Pester 5+; `BeforeAll { Import-Module ... -Force }`; `Describe` per function; `It` per parameter/behavior; parameter-validation-only (no COM) |
| **Error handling** | COM methods that throw on "no match" (e.g., `SpecialCells`) → wrap in `try/catch`, return empty/status |
| **Help** | Every function gets comment-based help: `.SYNOPSIS`, `.PARAMETER`, `.EXAMPLE` |

---

## 2. Phase 1 — Structural & Performance

### 2.1 StructuralOps.ps1 (**NEW** file — 4 functions)

Create `Public/StructuralOps.ps1` and `Tests/StructuralOps.Tests.ps1`.

---

#### 2.1.1 `Add-ExcelRow`

**Purpose:** Insert one or more blank rows at a specified position, shifting existing rows down.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory, string) · `SheetName` (Mandatory, string) · `Row` (Mandatory, int, 1-based) · `Count` (int, default 1) · `AsJson` (switch) |
| **Validation** | `[ValidateRange(1, [int]::MaxValue)]` on `Row`; `[ValidateRange(1, 65536)]` on `Count` |
| **COM call** | `$ws.Rows.Item($Row).Resize($Count).Insert([int]$script:XL_INSERT_SHIFT.down)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; rowsInserted=$Count; atRow=$Row }` |

**Implementation:**

```powershell
function Add-ExcelRow {
    <#
    .SYNOPSIS  Insert rows into a worksheet.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Row           1-based row number where insertion begins.
    .PARAMETER Count         Number of rows to insert (default 1).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Add-ExcelRow -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Row 5 -Count 3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$Row,
        [ValidateRange(1, 65536)][int]$Count = 1,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $ws.Rows.Item($Row).Resize($Count).Insert([int]$script:XL_INSERT_SHIFT.down)

    $result = @{
        status       = 'ok'
        sheet        = $SheetName
        rowsInserted = $Count
        atRow        = $Row
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (`StructuralOps.Tests.ps1`):**

```powershell
Describe 'Add-ExcelRow' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelRow).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory WorkbookPath parameter' {
        $p = (Get-Command Add-ExcelRow).Parameters['WorkbookPath']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It 'Has mandatory SheetName parameter' {
        $p = (Get-Command Add-ExcelRow).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Row parameter with ValidateRange' {
        $p = (Get-Command Add-ExcelRow).Parameters['Row']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Count defaults to 1' {
        $p = (Get-Command Add-ExcelRow).Parameters['Count']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelRow).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
```

---

#### 2.1.2 `Remove-ExcelRow`

**Purpose:** Delete one or more rows at a specified position, shifting remaining rows up.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Row` (Mandatory, int, 1-based) · `Count` (int, default 1) · `AsJson` (switch) |
| **Validation** | `[ValidateRange(1, [int]::MaxValue)]` on `Row`; `[ValidateRange(1, 65536)]` on `Count` |
| **COM call** | `$ws.Rows.Item($Row).Resize($Count).Delete([int]$script:XL_DELETE_SHIFT.up)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; rowsDeleted=$Count; atRow=$Row }` |

**Implementation:**

```powershell
function Remove-ExcelRow {
    <#
    .SYNOPSIS  Delete rows from a worksheet.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Row           1-based row number where deletion begins.
    .PARAMETER Count         Number of rows to delete (default 1).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Remove-ExcelRow -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Row 5 -Count 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$Row,
        [ValidateRange(1, 65536)][int]$Count = 1,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $ws.Rows.Item($Row).Resize($Count).Delete([int]$script:XL_DELETE_SHIFT.up)

    $result = @{
        status      = 'ok'
        sheet       = $SheetName
        rowsDeleted = $Count
        atRow       = $Row
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests:**

```powershell
Describe 'Remove-ExcelRow' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelRow).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Row parameter with ValidateRange' {
        $p = (Get-Command Remove-ExcelRow).Parameters['Row']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelRow).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
```

---

#### 2.1.3 `Add-ExcelColumn`

**Purpose:** Insert one or more blank columns at a specified position, shifting existing columns right.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Column` (Mandatory, string — letter e.g. `"C"` or int) · `Count` (int, default 1) · `AsJson` (switch) |
| **Validation** | `[ValidateRange(1, 16384)]` on `Count` |
| **COM call** | `$ws.Columns.Item($Column).Resize($null, $Count).Insert([int]$script:XL_INSERT_SHIFT.right)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; columnsInserted=$Count; atColumn=$Column }` |

**Implementation notes:**
- `$Column` accepts both letter (`"C"`) and integer (`3`) — Excel's `Columns.Item()` handles both natively.
- `Resize` for columns uses the second parameter: `.Resize($null, $Count)` or equivalently `.Resize([System.Reflection.Missing]::Value, $Count)`.

```powershell
function Add-ExcelColumn {
    <#
    .SYNOPSIS  Insert columns into a worksheet.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Column        Column letter (e.g. "C") or 1-based integer.
    .PARAMETER Count         Number of columns to insert (default 1).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Add-ExcelColumn -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Column "D" -Count 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)]$Column,
        [ValidateRange(1, 16384)][int]$Count = 1,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $ws.Columns.Item($Column).Resize([System.Reflection.Missing]::Value, $Count).Insert([int]$script:XL_INSERT_SHIFT.right)

    $result = @{
        status          = 'ok'
        sheet           = $SheetName
        columnsInserted = $Count
        atColumn        = $Column
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests:**

```powershell
Describe 'Add-ExcelColumn' {
    It 'Has CmdletBinding' {
        (Get-Command Add-ExcelColumn).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Column parameter' {
        (Get-Command Add-ExcelColumn).Parameters['Column'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Add-ExcelColumn).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
```

---

#### 2.1.4 `Remove-ExcelColumn`

**Purpose:** Delete one or more columns at a specified position, shifting remaining columns left.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Column` (Mandatory, string/int) · `Count` (int, default 1) · `AsJson` (switch) |
| **COM call** | `$ws.Columns.Item($Column).Resize([System.Reflection.Missing]::Value, $Count).Delete([int]$script:XL_DELETE_SHIFT.left)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; columnsDeleted=$Count; atColumn=$Column }` |

```powershell
function Remove-ExcelColumn {
    <#
    .SYNOPSIS  Delete columns from a worksheet.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Column        Column letter (e.g. "C") or 1-based integer.
    .PARAMETER Count         Number of columns to delete (default 1).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Remove-ExcelColumn -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Column "B" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)]$Column,
        [ValidateRange(1, 16384)][int]$Count = 1,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $ws.Columns.Item($Column).Resize([System.Reflection.Missing]::Value, $Count).Delete([int]$script:XL_DELETE_SHIFT.left)

    $result = @{
        status         = 'ok'
        sheet          = $SheetName
        columnsDeleted = $Count
        atColumn       = $Column
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests:**

```powershell
Describe 'Remove-ExcelColumn' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelColumn).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Column parameter' {
        (Get-Command Remove-ExcelColumn).Parameters['Column'] | Should -Not -BeNullOrEmpty
    }
    It 'Has AsJson switch' {
        (Get-Command Remove-ExcelColumn).Parameters['AsJson'].SwitchParameter | Should -BeTrue
    }
}
```

---

### 2.2 `Remove-ExcelDuplicates` (add to `FilterSortOps.ps1`)

**Purpose:** Remove duplicate rows from a range based on specified column indices.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory, string) · `Columns` (Mandatory, int array — 1-based column indices within range) · `HasHeader` (switch, default `$true`) · `AsJson` (switch) |
| **COM call** | `$rng.RemoveDuplicates($Columns, [int]$script:XL_YES_NO_GUESS.yes)` |
| **Header logic** | If `$HasHeader`, pass `$script:XL_YES_NO_GUESS.yes`; else pass `$script:XL_YES_NO_GUESS.no` |
| **Return** | `@{ status='ok'; range=$Range; hasHeader=$HasHeader.IsPresent }` |

```powershell
function Remove-ExcelDuplicates {
    <#
    .SYNOPSIS  Remove duplicate rows from a range.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Range         Range address (e.g. "A1:D100").
    .PARAMETER Columns       1-based column indices within the range to compare for duplicates.
    .PARAMETER HasHeader     Indicates the first row is a header (default $true).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Remove-ExcelDuplicates -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D100" -Columns 1,3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][int[]]$Columns,
        [switch]$HasHeader,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $headerVal = if ($HasHeader) { [int]$script:XL_YES_NO_GUESS.yes } else { [int]$script:XL_YES_NO_GUESS.no }
    $rng.RemoveDuplicates($Columns, $headerVal)

    $result = @{
        status    = 'ok'
        range     = $Range
        hasHeader = $HasHeader.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (append to `FilterSortOps.Tests.ps1`):**

```powershell
Describe 'Remove-ExcelDuplicates' {
    It 'Has CmdletBinding' {
        (Get-Command Remove-ExcelDuplicates).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Range parameter' {
        (Get-Command Remove-ExcelDuplicates).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has mandatory Columns parameter (int array)' {
        $p = (Get-Command Remove-ExcelDuplicates).Parameters['Columns']
        $p | Should -Not -BeNullOrEmpty
        $p.ParameterType.Name | Should -Be 'Int32[]'
    }
    It 'Has HasHeader switch' {
        (Get-Command Remove-ExcelDuplicates).Parameters['HasHeader'].SwitchParameter | Should -BeTrue
    }
}
```

---

### 2.3 CalculationOps.ps1 (**NEW** file — 4 functions)

Create `Public/CalculationOps.ps1` and `Tests/CalculationOps.Tests.ps1`.

**New constant required** — add to `ExcelPOSH.psm1`:

```powershell
$script:XL_CALCULATION = @{
    automatic     = -4105  # xlCalculationAutomatic
    manual        = -4135  # xlCalculationManual
    semiautomatic = 2      # xlCalculationSemiautomatic
}
```

---

#### 2.3.1 `Set-ExcelPerformanceMode`

**Purpose:** Toggle Excel performance settings for bulk operations — `ScreenUpdating`, `Calculation`, `EnableEvents`.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `ScreenUpdating` (bool, optional) · `Calculation` (ValidateSet: `'Automatic','Manual','SemiAutomatic'`, optional) · `EnableEvents` (bool, optional) · `AsJson` (switch) |
| **COM call** | Set each property only if the parameter was explicitly bound (use `$PSBoundParameters.ContainsKey()`) |
| **Return** | `@{ status='ok'; screenUpdating=$app.ScreenUpdating; calculation=...; enableEvents=$app.EnableEvents }` |

**Implementation notes:**
- No `SheetName` parameter — this is application-scoped.
- Read current values *after* setting to confirm.
- Map `Calculation` string → `$script:XL_CALCULATION[$Calculation.ToLower()]`.

```powershell
function Set-ExcelPerformanceMode {
    <#
    .SYNOPSIS  Control Excel performance settings (ScreenUpdating, Calculation, EnableEvents).
    .PARAMETER WorkbookPath     Path to the Excel workbook.
    .PARAMETER ScreenUpdating   Enable/disable screen redraw.
    .PARAMETER Calculation       Calculation mode: Automatic, Manual, SemiAutomatic.
    .PARAMETER EnableEvents     Enable/disable event processing.
    .PARAMETER AsJson           Return JSON string.
    .EXAMPLE   Set-ExcelPerformanceMode -WorkbookPath C:\data.xlsx -ScreenUpdating $false -Calculation Manual -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [bool]$ScreenUpdating,
        [ValidateSet('Automatic','Manual','SemiAutomatic')]
        [string]$Calculation,
        [bool]$EnableEvents,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    if ($PSBoundParameters.ContainsKey('ScreenUpdating')) {
        $app.ScreenUpdating = $ScreenUpdating
    }
    if ($PSBoundParameters.ContainsKey('Calculation')) {
        $app.Calculation = [int]$script:XL_CALCULATION[$Calculation.ToLower()]
    }
    if ($PSBoundParameters.ContainsKey('EnableEvents')) {
        $app.EnableEvents = $EnableEvents
    }

    # Map numeric calculation value back to name
    $calcReverse = @{ -4105 = 'Automatic'; -4135 = 'Manual'; 2 = 'SemiAutomatic' }
    $calcName    = $calcReverse[[int]$app.Calculation]
    if (-not $calcName) { $calcName = [string]$app.Calculation }

    $result = @{
        status         = 'ok'
        screenUpdating = [bool]$app.ScreenUpdating
        calculation    = $calcName
        enableEvents   = [bool]$app.EnableEvents
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests:**

```powershell
Describe 'Set-ExcelPerformanceMode' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelPerformanceMode).CmdletBinding | Should -BeTrue
    }
    It 'Has Calculation parameter with ValidateSet' {
        $p = (Get-Command Set-ExcelPerformanceMode).Parameters['Calculation']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Calculation ValidateSet contains Automatic, Manual, SemiAutomatic' {
        $vs = (Get-Command Set-ExcelPerformanceMode).Parameters['Calculation'].Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.ValidValues | Should -Contain 'Automatic'
        $vs.ValidValues | Should -Contain 'Manual'
        $vs.ValidValues | Should -Contain 'SemiAutomatic'
    }
    It 'Does not require SheetName' {
        (Get-Command Set-ExcelPerformanceMode).Parameters.Keys | Should -Not -Contain 'SheetName'
    }
}
```

---

#### 2.3.2 `Invoke-ExcelCalculate`

**Purpose:** Trigger Excel's calculation engine — full workbook, active sheet, or specific sheet.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `Full` (switch — forces `CalculateFull`) · `SheetName` (optional — calculate specific sheet) · `AsJson` (switch) |
| **COM call** | If `$Full`: `$app.CalculateFull()`; elseif `$SheetName`: `$ws.Calculate()`; else `$app.Calculate()` |
| **Return** | `@{ status='ok'; scope=('full' or 'sheet:$SheetName' or 'workbook') }` |

```powershell
function Invoke-ExcelCalculate {
    <#
    .SYNOPSIS  Trigger Excel recalculation.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER Full          Force full recalculation of all open workbooks.
    .PARAMETER SheetName     Calculate a specific worksheet only.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Invoke-ExcelCalculate -WorkbookPath C:\data.xlsx -Full -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$Full,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    if ($Full) {
        $app.CalculateFull()
        $scope = 'full'
    } elseif (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $ws = $app.ActiveWorkbook.Worksheets.Item($SheetName)
        $ws.Calculate()
        $scope = "sheet:$SheetName"
    } else {
        $app.Calculate()
        $scope = 'workbook'
    }

    $result = @{
        status = 'ok'
        scope  = $scope
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

---

#### 2.3.3 `Invoke-ExcelFunction`

**Purpose:** Call an Excel `WorksheetFunction` method from PowerShell — e.g., `SUMIF`, `VLOOKUP`, `COUNTIF`.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `FunctionName` (Mandatory, string) · `Arguments` (object array) · `SheetName` (optional, for resolving range arguments) · `AsJson` (switch) |
| **COM call** | `$app.WorksheetFunction.$FunctionName.Invoke($resolvedArgs)` via `InvokeMember` |
| **Range resolution** | If an argument matches pattern `^[A-Z]+\d+` or contains `!`, resolve to a COM Range object |
| **Return** | `@{ status='ok'; function=$FunctionName; result=$result }` |

**Implementation notes:**
- Excel's `WorksheetFunction` has ~400+ methods. Use late-bound `InvokeMember` to call any method by name.
- Range arguments: scan each element of `$Arguments`; if it looks like a cell reference (regex `^\'?[\w\s]+\'?\![A-Z]+\d+|^[A-Z]+\d+`), resolve it to `$ws.Range($arg)` or `$app.ActiveWorkbook.Worksheets.Item($sheetRef).Range($cellRef)`.
- Non-range scalar arguments pass through as-is.
- Wrap the invocation in `try/catch` — WorksheetFunction throws on invalid args (e.g., `#N/A` from `VLOOKUP`).

```powershell
function Invoke-ExcelFunction {
    <#
    .SYNOPSIS  Call an Excel WorksheetFunction method.
    .PARAMETER WorkbookPath   Path to the Excel workbook.
    .PARAMETER FunctionName   Name of the WorksheetFunction method (e.g. "SumIf", "VLookup").
    .PARAMETER Arguments      Array of arguments — scalars or range address strings.
    .PARAMETER SheetName      Default worksheet for resolving bare range addresses (e.g. "A1:A10").
    .PARAMETER AsJson         Return JSON string.
    .EXAMPLE   Invoke-ExcelFunction -WorkbookPath C:\data.xlsx -FunctionName "Sum" -Arguments @("Sheet1!A1:A10") -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$FunctionName,
        [object[]]$Arguments,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    # Default worksheet for bare range refs
    $defaultWs = if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $wb.Worksheets.Item($SheetName)
    } else {
        $wb.ActiveSheet
    }

    # Resolve range references in arguments
    $resolvedArgs = @()
    foreach ($arg in $Arguments) {
        if ($arg -is [string] -and $arg -match "^'?(.+?)'?\!([A-Z]+\d+.*)$") {
            # Sheet-qualified reference: "Sheet1!A1:A10"
            $sheetRef = $matches[1]
            $cellRef  = $matches[2]
            $resolvedArgs += $wb.Worksheets.Item($sheetRef).Range($cellRef)
        } elseif ($arg -is [string] -and $arg -match '^[A-Z]+\d+(:[A-Z]+\d+)?$') {
            # Bare range reference: "A1:A10"
            $resolvedArgs += $defaultWs.Range($arg)
        } else {
            $resolvedArgs += $arg
        }
    }

    try {
        $wf     = $app.WorksheetFunction
        $result = $wf.GetType().InvokeMember(
            $FunctionName,
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $wf,
            $resolvedArgs
        )
    } catch {
        throw "WorksheetFunction.$FunctionName failed: $_"
    }

    $output = @{
        status   = 'ok'
        function = $FunctionName
        result   = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}
```

---

#### 2.3.4 `Invoke-ExcelEvaluate`

**Purpose:** Evaluate an Excel formula expression string and return the result — equivalent to typing `=expression` in a cell.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `Expression` (Mandatory, string) · `SheetName` (optional) · `AsJson` (switch) |
| **COM call** | If `$SheetName`: `$ws.Evaluate($Expression)`; else `$app.Evaluate($Expression)` |
| **Return** | `@{ status='ok'; expression=$Expression; result=$result }` |

**Implementation notes:**
- Expression must follow Excel formula syntax (e.g., `"SUM(A1:A10)"`, `"A1*B1+C1"`).
- If result is an error value (e.g., `2015` = `#VALUE!`), convert to descriptive string.

```powershell
function Invoke-ExcelEvaluate {
    <#
    .SYNOPSIS  Evaluate an Excel formula expression.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER Expression    Excel formula expression (without leading =).
    .PARAMETER SheetName     Worksheet context for cell references.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Invoke-ExcelEvaluate -WorkbookPath C:\data.xlsx -Expression "SUM(A1:A10)" -SheetName Sheet1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Expression,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
        $result = $ws.Evaluate($Expression)
    } else {
        $result = $app.Evaluate($Expression)
    }

    # Handle COM error values
    if ($result -is [int] -and $result -gt 2000) {
        $errorMap = @{ 2007 = '#DIV/0!'; 2015 = '#VALUE!'; 2023 = '#REF!'; 2029 = '#NAME?'; 2036 = '#NUM!'; 2042 = '#N/A'; 2000 = '#NULL!' }
        if ($errorMap.ContainsKey($result)) { $result = $errorMap[$result] }
    }

    $output = @{
        status     = 'ok'
        expression = $Expression
        result     = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}
```

**Tests (`CalculationOps.Tests.ps1`):**

```powershell
Describe 'Set-ExcelPerformanceMode' {
    # ... (see §2.3.1 above)
}

Describe 'Invoke-ExcelCalculate' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelCalculate).CmdletBinding | Should -BeTrue
    }
    It 'Has Full switch' {
        (Get-Command Invoke-ExcelCalculate).Parameters['Full'].SwitchParameter | Should -BeTrue
    }
    It 'SheetName is optional' {
        $p = (Get-Command Invoke-ExcelCalculate).Parameters['SheetName']
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory }) | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-ExcelFunction' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelFunction).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory FunctionName parameter' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['FunctionName']
        $p | Should -Not -BeNullOrEmpty
    }
    It 'Has Arguments parameter (object array)' {
        $p = (Get-Command Invoke-ExcelFunction).Parameters['Arguments']
        $p | Should -Not -BeNullOrEmpty
    }
}

Describe 'Invoke-ExcelEvaluate' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelEvaluate).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Expression parameter' {
        $p = (Get-Command Invoke-ExcelEvaluate).Parameters['Expression']
        $p | Should -Not -BeNullOrEmpty
    }
}
```

---

### 2.4 `Merge-ExcelRange` / `Split-ExcelRange` (add to `FormattingOps.ps1`)

#### 2.4.1 `Merge-ExcelRange`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory) · `Across` (switch — merge each row separately) · `AsJson` (switch) |
| **COM call** | If `$Across`: `$rng.Merge($true)`; else `$rng.Merge()` |
| **Return** | `@{ status='ok'; range=$Range; merged=$true; across=$Across.IsPresent }` |

```powershell
function Merge-ExcelRange {
    <#
    .SYNOPSIS  Merge cells in a range.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Range         Range address to merge (e.g. "A1:D1").
    .PARAMETER Across        Merge each row separately instead of the entire range.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Merge-ExcelRange -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$Across,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $app.DisplayAlerts = $false
    if ($Across) { $rng.Merge($true) } else { $rng.Merge() }
    $app.DisplayAlerts = $true

    $result = @{
        status = 'ok'
        range  = $Range
        merged = $true
        across = $Across.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Implementation note:** `$app.DisplayAlerts = $false` is required because merging a range with data in multiple cells triggers a confirmation dialog. Restore afterward.

#### 2.4.2 `Split-ExcelRange`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory) · `AsJson` (switch) |
| **COM call** | `$rng.UnMerge()` |
| **Return** | `@{ status='ok'; range=$Range; merged=$false }` |

```powershell
function Split-ExcelRange {
    <#
    .SYNOPSIS  Unmerge (split) previously merged cells.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Range         Range address to unmerge.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Split-ExcelRange -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $rng.UnMerge()

    $result = @{
        status = 'ok'
        range  = $Range
        merged = $false
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (append to `FormattingOps.Tests.ps1`):**

```powershell
Describe 'Merge-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Merge-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Range parameter' {
        (Get-Command Merge-ExcelRange).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Across switch' {
        (Get-Command Merge-ExcelRange).Parameters['Across'].SwitchParameter | Should -BeTrue
    }
}

Describe 'Split-ExcelRange' {
    It 'Has CmdletBinding' {
        (Get-Command Split-ExcelRange).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory Range parameter' {
        (Get-Command Split-ExcelRange).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
}
```

---

### 2.5 `Get-ExcelSpecialCells` (add to `WorksheetOps.ps1`)

**Purpose:** Return addresses of cells matching a type — blanks, constants, formulas, errors, visible, last cell, or comments.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (optional, defaults to UsedRange) · `CellType` (Mandatory, ValidateSet: `'Blanks','Constants','Formulas','Errors','Visible','LastCell','Comments'`) · `AsJson` (switch) |
| **New constant** | `$script:XL_CELL_TYPE` (see §5) |
| **COM call** | `$rng.SpecialCells([int]$script:XL_CELL_TYPE[$CellType.ToLower()])` |
| **Edge case** | `SpecialCells` throws if no matching cells exist — wrap in `try/catch`, return `count=0, addresses=@()` |
| **Return** | `@{ status='ok'; cellType=$CellType; addresses=@($cells.Address($false,$false) -split ','); count=$cells.Count }` |

```powershell
function Get-ExcelSpecialCells {
    <#
    .SYNOPSIS  Find cells of a specific type (blanks, formulas, constants, etc.).
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Range         Range to search (defaults to UsedRange).
    .PARAMETER CellType      Type of special cells to find.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Get-ExcelSpecialCells -WorkbookPath C:\data.xlsx -SheetName Sheet1 -CellType Formulas -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Range,
        [Parameter(Mandatory)]
        [ValidateSet('Blanks','Constants','Formulas','Errors','Visible','LastCell','Comments')]
        [string]$CellType,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $rng = if ([string]::IsNullOrWhiteSpace($Range)) { $ws.UsedRange } else { $ws.Range($Range) }

    try {
        $cells     = $rng.SpecialCells([int]$script:XL_CELL_TYPE[$CellType.ToLower()])
        $addresses = @($cells.Address($false, $false) -split ',')
        $cellCount = $cells.Count
    } catch {
        # No cells of this type exist — not an error
        $addresses = @()
        $cellCount = 0
    }

    $result = @{
        status    = 'ok'
        cellType  = $CellType
        addresses = $addresses
        count     = $cellCount
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (append to `WorksheetOps.Tests.ps1`):**

```powershell
Describe 'Get-ExcelSpecialCells' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelSpecialCells).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory CellType parameter with ValidateSet' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['CellType']
        $p | Should -Not -BeNullOrEmpty
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs | Should -Not -BeNullOrEmpty
        $vs.ValidValues | Should -Contain 'Blanks'
        $vs.ValidValues | Should -Contain 'Formulas'
        $vs.ValidValues | Should -Contain 'LastCell'
    }
    It 'Range is optional' {
        $p = (Get-Command Get-ExcelSpecialCells).Parameters['Range']
        $p | Should -Not -BeNullOrEmpty
        $mandatory = $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory })
        $mandatory | Should -BeNullOrEmpty
    }
}
```

---

### 2.6 `Invoke-ExcelMacro` (add to `WorkbookOps.ps1`)

**Purpose:** Execute a VBA macro in the workbook, passing up to 30 arguments.

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `MacroName` (Mandatory, string) · `Arguments` (object array, optional, max 30) · `AsJson` (switch) |
| **COM call** | Build `$app.Run($MacroName, $a1, $a2, ...)` via splatting — COM `Run` accepts up to 30 positional args |
| **Return** | `@{ status='ok'; macroName=$MacroName; result=$result }` |

**Implementation notes:**
- `$app.Run()` does not accept an array — each argument must be a separate positional parameter.
- Use `Invoke` with reflection to splat the array: `$app.GetType().InvokeMember('Run', ...)` with the macro name prepended to the args array.
- The workbook must be `.xlsm` or `.xlsb` — `.xlsx` has no macros. No enforcement needed (Excel will throw).

```powershell
function Invoke-ExcelMacro {
    <#
    .SYNOPSIS  Run a VBA macro in the workbook.
    .PARAMETER WorkbookPath  Path to the Excel workbook (.xlsm or .xlsb).
    .PARAMETER MacroName     Full macro name (e.g. "Sheet1.MyMacro" or just "MyMacro").
    .PARAMETER Arguments     Up to 30 arguments to pass to the macro.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Invoke-ExcelMacro -WorkbookPath C:\data.xlsm -MacroName "CleanData" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$MacroName,
        [ValidateCount(0, 30)][object[]]$Arguments,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    # Build argument list: macro name + up to 30 args
    $runArgs = @($MacroName) + @($Arguments)

    $result = $app.GetType().InvokeMember(
        'Run',
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null,
        $app,
        $runArgs
    )

    $output = @{
        status    = 'ok'
        macroName = $MacroName
        result    = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}
```

**Tests (append to `WorkbookOps.Tests.ps1`):**

```powershell
Describe 'Invoke-ExcelMacro' {
    It 'Has CmdletBinding' {
        (Get-Command Invoke-ExcelMacro).CmdletBinding | Should -BeTrue
    }
    It 'Has mandatory MacroName parameter' {
        (Get-Command Invoke-ExcelMacro).Parameters['MacroName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Arguments parameter with ValidateCount(0,30)' {
        $p = (Get-Command Invoke-ExcelMacro).Parameters['Arguments']
        $p | Should -Not -BeNullOrEmpty
        $vc = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateCountAttribute] })
        $vc | Should -Not -BeNullOrEmpty
    }
}
```

---

## 3. Phase 2 — Calculation Engine & Charting

### 3.1 `Sort-ExcelRange` Refactor (in `FilterSortOps.ps1`)

**Goal:** Add modern multi-key sorting via a `SortKeys` parameter while keeping backward compatibility with existing `SortKey1`/`Order1`/`SortKey2`/`Order2` parameters.

| Change | Detail |
|---|---|
| **New parameter** | `SortKeys` — array of hashtables `@(@{Key='A1'; Order='Ascending'}, ...)` |
| **Parameter sets** | `LegacySort` (existing params) + `ModernSort` (SortKeys) — use `[Parameter(ParameterSetName=...)]` |
| **Modern COM API** | Uses the `Sort` object on the worksheet |

**Modern Sort COM sequence:**

```powershell
if ($PSCmdlet.ParameterSetName -eq 'ModernSort') {
    $sort = $ws.Sort
    $sort.SortFields.Clear()
    foreach ($sk in $SortKeys) {
        $order = [int]$script:XL_SORT_ORDER[$sk.Order.ToLower()]
        $sort.SortFields.Add2($ws.Range($sk.Key), $null, $order)
    }
    $sort.SetRange($ws.Range($Range))
    $sort.Header = [int]$headerVal
    $sort.Apply()
}
```

**Tests (update `FilterSortOps.Tests.ps1`):**

```powershell
Describe 'Sort-ExcelRange (SortKeys)' {
    It 'Has SortKeys parameter' {
        (Get-Command Sort-ExcelRange).Parameters['SortKeys'] | Should -Not -BeNullOrEmpty
    }
    It 'SortKeys and SortKey1 are in different parameter sets' {
        $sk  = (Get-Command Sort-ExcelRange).Parameters['SortKeys']
        $sk1 = (Get-Command Sort-ExcelRange).Parameters['SortKey1']
        $skSet  = $sk.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).ParameterSetName
        $sk1Set = $sk1.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).ParameterSetName
        $skSet | Should -Not -Be $sk1Set
    }
}
```

---

### 3.2 `Invoke-ExcelAdvancedFilter` (add to `FilterSortOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory) · `Action` (Mandatory, ValidateSet: `'FilterInPlace','CopyToRange'`) · `CriteriaRange` (optional) · `CopyToRange` (optional, required when Action=CopyToRange) · `UniqueOnly` (switch) · `AsJson` (switch) |
| **New constant** | `$script:XL_FILTER_ACTION` (see §5) |
| **COM call** | `$rng.AdvancedFilter([int]$action, $criteria, $copyTo, $UniqueOnly)` |
| **Validation** | If `Action = 'CopyToRange'` and `CopyToRange` is empty, throw a descriptive error |
| **Return** | `@{ status='ok'; action=$Action; unique=$UniqueOnly.IsPresent }` |

```powershell
function Invoke-ExcelAdvancedFilter {
    <#
    .SYNOPSIS  Apply Excel Advanced Filter to a range.
    .PARAMETER WorkbookPath   Path to the Excel workbook.
    .PARAMETER SheetName      Target worksheet name.
    .PARAMETER Range          Source data range.
    .PARAMETER Action         FilterInPlace or CopyToRange.
    .PARAMETER CriteriaRange  Range containing filter criteria.
    .PARAMETER CopyToRange    Destination range (required for CopyToRange action).
    .PARAMETER UniqueOnly     Return only unique records.
    .PARAMETER AsJson         Return JSON string.
    .EXAMPLE   Invoke-ExcelAdvancedFilter -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D100" -Action FilterInPlace -UniqueOnly -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)]
        [ValidateSet('FilterInPlace','CopyToRange')]
        [string]$Action,
        [string]$CriteriaRange,
        [string]$CopyToRange,
        [switch]$UniqueOnly,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $actionVal = [int]$script:XL_FILTER_ACTION[$Action.ToLower()]

    $criteria = if (-not [string]::IsNullOrWhiteSpace($CriteriaRange)) { $ws.Range($CriteriaRange) } else { [System.Reflection.Missing]::Value }
    $copyTo   = if (-not [string]::IsNullOrWhiteSpace($CopyToRange))   { $ws.Range($CopyToRange)   } else { [System.Reflection.Missing]::Value }

    if ($Action -eq 'CopyToRange' -and [string]::IsNullOrWhiteSpace($CopyToRange)) {
        throw 'CopyToRange parameter is required when Action is CopyToRange.'
    }

    $rng.AdvancedFilter($actionVal, $criteria, $copyTo, $UniqueOnly.IsPresent)

    $result = @{
        status = 'ok'
        action = $Action
        unique = $UniqueOnly.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

---

### 3.3 Chart Functions (add to `ChartOps.ps1` — 4 functions)

All chart functions follow this common preamble:

```powershell
$app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
$ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
$co    = $ws.ChartObjects($ChartName)
$chart = $co.Chart
```

#### 3.3.1 `Set-ExcelChartSeries`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `ChartName` (all Mandatory) · `SeriesIndex` (Mandatory, int) · `Name` (string) · `Values` (string — range ref) · `XValues` (string — range ref) · `FillColor` (string — hex) · `LineColor` (string — hex) · `LineWeight` (double) · `AsJson` |
| **COM** | `$series = $chart.SeriesCollection($SeriesIndex)` → set `.Name`, `.Values = $ws.Range(...)`, `.XValues`, fill/line via `.Format` |
| **Return** | `@{ status='ok'; chart=$ChartName; seriesIndex=$SeriesIndex }` |

```powershell
function Set-ExcelChartSeries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][int]$SeriesIndex,
        [string]$Name,
        [string]$Values,
        [string]$XValues,
        [string]$FillColor,
        [string]$LineColor,
        [double]$LineWeight,
        [switch]$AsJson
    )

    $app    = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart  = $ws.ChartObjects($ChartName).Chart
    $series = $chart.SeriesCollection($SeriesIndex)

    if ($PSBoundParameters.ContainsKey('Name'))      { $series.Name    = $Name }
    if ($PSBoundParameters.ContainsKey('Values'))     { $series.Values  = $ws.Range($Values) }
    if ($PSBoundParameters.ContainsKey('XValues'))    { $series.XValues = $ws.Range($XValues) }
    if ($PSBoundParameters.ContainsKey('FillColor'))  { $series.Format.Fill.ForeColor.RGB = ConvertTo-RGBColor $FillColor }
    if ($PSBoundParameters.ContainsKey('LineColor'))  { $series.Format.Line.ForeColor.RGB = ConvertTo-RGBColor $LineColor }
    if ($PSBoundParameters.ContainsKey('LineWeight')) { $series.Format.Line.Weight = $LineWeight }

    $result = @{ status = 'ok'; chart = $ChartName; seriesIndex = $SeriesIndex }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

#### 3.3.2 `Set-ExcelChartAxis`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `ChartName` (all Mandatory) · `AxisType` (Mandatory, ValidateSet: `'Category','Value','SeriesAxis'`) · `AxisGroup` (ValidateSet: `'Primary','Secondary'`, default Primary) · `Title` (string) · `MinimumScale` (double) · `MaximumScale` (double) · `NumberFormat` (string) · `AsJson` |
| **New constants** | `$script:XL_AXIS_TYPE`, `$script:XL_AXIS_GROUP` (see §5) |
| **COM** | `$axis = $chart.Axes([int]$xlAxisType, [int]$xlAxisGroup)` |
| **Return** | `@{ status='ok'; chart=$ChartName; axisType=$AxisType; axisGroup=$AxisGroup }` |

```powershell
function Set-ExcelChartAxis {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)]
        [ValidateSet('Category','Value','SeriesAxis')]
        [string]$AxisType,
        [ValidateSet('Primary','Secondary')]
        [string]$AxisGroup = 'Primary',
        [string]$Title,
        [double]$MinimumScale,
        [double]$MaximumScale,
        [string]$NumberFormat,
        [switch]$AsJson
    )

    $app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart = $ws.ChartObjects($ChartName).Chart

    $axType  = [int]$script:XL_AXIS_TYPE[$AxisType.ToLower()]
    $axGroup = [int]$script:XL_AXIS_GROUP[$AxisGroup.ToLower()]
    $axis    = $chart.Axes($axType, $axGroup)

    if ($PSBoundParameters.ContainsKey('Title')) {
        $axis.HasTitle       = $true
        $axis.AxisTitle.Text = $Title
    }
    if ($PSBoundParameters.ContainsKey('MinimumScale')) { $axis.MinimumScale = $MinimumScale }
    if ($PSBoundParameters.ContainsKey('MaximumScale')) { $axis.MaximumScale = $MaximumScale }
    if ($PSBoundParameters.ContainsKey('NumberFormat')) { $axis.TickLabels.NumberFormat = $NumberFormat }

    $result = @{ status = 'ok'; chart = $ChartName; axisType = $AxisType; axisGroup = $AxisGroup }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

#### 3.3.3 `Set-ExcelChartLegend`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `ChartName` (all Mandatory) · `Show` (Mandatory, bool) · `Position` (ValidateSet: `'Bottom','Corner','Left','Right','Top'`) · `AsJson` |
| **New constant** | `$script:XL_LEGEND_POSITION` (see §5) |
| **COM** | `$chart.HasLegend = $Show`; if `$Show -and $Position`: `$chart.Legend.Position = [int]$xlPos` |
| **Return** | `@{ status='ok'; chart=$ChartName; hasLegend=$Show }` |

```powershell
function Set-ExcelChartLegend {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][bool]$Show,
        [ValidateSet('Bottom','Corner','Left','Right','Top')]
        [string]$Position,
        [switch]$AsJson
    )

    $app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart = $ws.ChartObjects($ChartName).Chart

    $chart.HasLegend = $Show
    if ($Show -and $PSBoundParameters.ContainsKey('Position')) {
        $chart.Legend.Position = [int]$script:XL_LEGEND_POSITION[$Position.ToLower()]
    }

    $result = @{ status = 'ok'; chart = $ChartName; hasLegend = $Show }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

#### 3.3.4 `Set-ExcelChartDataLabels`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `ChartName` (all Mandatory) · `SeriesIndex` (Mandatory, int) · `ShowValue` (bool) · `ShowCategory` (bool) · `ShowPercentage` (bool) · `NumberFormat` (string) · `Position` (ValidateSet) · `AsJson` |
| **COM** | `$series.HasDataLabels = $true`; then set `.DataLabels.ShowValue`, etc. |
| **Return** | `@{ status='ok'; chart=$ChartName; seriesIndex=$SeriesIndex }` |

```powershell
function Set-ExcelChartDataLabels {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][int]$SeriesIndex,
        [bool]$ShowValue,
        [bool]$ShowCategory,
        [bool]$ShowPercentage,
        [string]$NumberFormat,
        [ValidateSet('Center','InsideBase','InsideEnd','OutsideEnd','BestFit')]
        [string]$Position,
        [switch]$AsJson
    )

    $app    = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart  = $ws.ChartObjects($ChartName).Chart
    $series = $chart.SeriesCollection($SeriesIndex)

    $series.HasDataLabels = $true
    $labels = $series.DataLabels

    if ($PSBoundParameters.ContainsKey('ShowValue'))      { $labels.ShowValue            = $ShowValue }
    if ($PSBoundParameters.ContainsKey('ShowCategory'))   { $labels.ShowCategoryName     = $ShowCategory }
    if ($PSBoundParameters.ContainsKey('ShowPercentage')) { $labels.ShowPercentage        = $ShowPercentage }
    if ($PSBoundParameters.ContainsKey('NumberFormat'))   { $labels.NumberFormat          = $NumberFormat }
    if ($PSBoundParameters.ContainsKey('Position')) {
        $posMap = @{ Center=0; InsideBase=3; InsideEnd=1; OutsideEnd=2; BestFit=5 }
        $labels.Position = $posMap[$Position]
    }

    $result = @{ status = 'ok'; chart = $ChartName; seriesIndex = $SeriesIndex }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (append to `ChartOps.Tests.ps1`):**

```powershell
Describe 'Set-ExcelChartSeries' {
    It 'Has mandatory ChartName and SeriesIndex' {
        (Get-Command Set-ExcelChartSeries).Parameters['ChartName'] | Should -Not -BeNullOrEmpty
        (Get-Command Set-ExcelChartSeries).Parameters['SeriesIndex'] | Should -Not -BeNullOrEmpty
    }
}
Describe 'Set-ExcelChartAxis' {
    It 'Has AxisType with ValidateSet' {
        $p = (Get-Command Set-ExcelChartAxis).Parameters['AxisType']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] }) | Should -Not -BeNullOrEmpty
    }
}
Describe 'Set-ExcelChartLegend' {
    It 'Has mandatory Show parameter' {
        (Get-Command Set-ExcelChartLegend).Parameters['Show'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Position with ValidateSet' {
        $p = (Get-Command Set-ExcelChartLegend).Parameters['Position']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] }) | Should -Not -BeNullOrEmpty
    }
}
Describe 'Set-ExcelChartDataLabels' {
    It 'Has mandatory SeriesIndex parameter' {
        (Get-Command Set-ExcelChartDataLabels).Parameters['SeriesIndex'] | Should -Not -BeNullOrEmpty
    }
}
```

---

### 3.4 Pivot Table Functions (add to `PivotTableOps.ps1` — 2 functions)

#### 3.4.1 `Set-ExcelPivotField`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `PivotTableName` (all Mandatory) · `FieldName` (Mandatory) · `Subtotals` (ValidateSet array) · `NumberFormat` (string) · `LayoutForm` (ValidateSet: `'Compact','Tabular','Outline'`) · `AsJson` |
| **COM** | `$pvt = $ws.PivotTables($PivotTableName)` → `$field = $pvt.PivotFields($FieldName)` |
| **Return** | `@{ status='ok'; pivotTable=$PivotTableName; field=$FieldName }` |

```powershell
function Set-ExcelPivotField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$PivotTableName,
        [Parameter(Mandatory)][string]$FieldName,
        [ValidateSet('Automatic','Sum','Count','Average','Max','Min','Product','CountNums','StdDev','Var','None')]
        [string[]]$Subtotals,
        [string]$NumberFormat,
        [ValidateSet('Compact','Tabular','Outline')]
        [string]$LayoutForm,
        [switch]$AsJson
    )

    $app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $pvt   = $ws.PivotTables($PivotTableName)
    $field = $pvt.PivotFields($FieldName)

    if ($PSBoundParameters.ContainsKey('NumberFormat')) { $field.NumberFormat = $NumberFormat }
    if ($PSBoundParameters.ContainsKey('LayoutForm')) {
        $layoutMap = @{ Compact = 0; Tabular = 1; Outline = 2 }
        $field.LayoutForm = $layoutMap[$LayoutForm]
    }
    if ($PSBoundParameters.ContainsKey('Subtotals')) {
        # Subtotals is a 1-based boolean array: (1=Automatic,2=Sum,3=Count,...,12=Var)
        $subArray = @($false) * 12
        $subMap = @{ Automatic=0; Sum=1; Count=2; Average=3; Max=4; Min=5; Product=6; CountNums=7; StdDev=8; Var=9; None=-1 }
        foreach ($s in $Subtotals) {
            $idx = $subMap[$s]
            if ($idx -ge 0) { $subArray[$idx] = $true }
        }
        if ($Subtotals -contains 'None') { $subArray = @($false) * 12 }
        $field.Subtotals = $subArray
    }

    $result = @{ status = 'ok'; pivotTable = $PivotTableName; field = $FieldName }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

#### 3.4.2 `Add-ExcelPivotCalculatedField`

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath`, `SheetName`, `PivotTableName` (all Mandatory) · `Name` (Mandatory) · `Formula` (Mandatory) · `AsJson` |
| **COM** | `$pvt.CalculatedFields.Add($Name, $Formula)` |
| **Return** | `@{ status='ok'; pivotTable=$PivotTableName; calculatedField=$Name }` |

```powershell
function Add-ExcelPivotCalculatedField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$PivotTableName,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Formula,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $pvt = $ws.PivotTables($PivotTableName)

    $pvt.CalculatedFields.Add($Name, $Formula) | Out-Null

    $result = @{
        status          = 'ok'
        pivotTable      = $PivotTableName
        calculatedField = $Name
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
```

**Tests (append to `PivotTableOps.Tests.ps1`):**

```powershell
Describe 'Set-ExcelPivotField' {
    It 'Has mandatory PivotTableName and FieldName' {
        (Get-Command Set-ExcelPivotField).Parameters['PivotTableName'] | Should -Not -BeNullOrEmpty
        (Get-Command Set-ExcelPivotField).Parameters['FieldName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has LayoutForm with ValidateSet' {
        $p = (Get-Command Set-ExcelPivotField).Parameters['LayoutForm']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] }) | Should -Not -BeNullOrEmpty
    }
}
Describe 'Add-ExcelPivotCalculatedField' {
    It 'Has mandatory Name and Formula' {
        (Get-Command Add-ExcelPivotCalculatedField).Parameters['Name'] | Should -Not -BeNullOrEmpty
        (Get-Command Add-ExcelPivotCalculatedField).Parameters['Formula'] | Should -Not -BeNullOrEmpty
    }
}
```

---

## 4. Phase 3 — Medium Priority

### 4.1 `Split-ExcelColumn` (add to `ImportOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory — single column, e.g. `"A1:A100"`) · `Delimiter` (Mandatory — string, e.g. `","`, `"|"`, `"Tab"`) · `FieldInfo` (optional — array of `@(colIndex, dataFormat)` pairs for TextToColumns) · `AsJson` |
| **COM** | `$rng.TextToColumns($destination, [int]1, ...)` — `1` = `xlDelimited`; set appropriate delimiter flag |
| **Delimiter mapping** | `Tab` → `$true` for Tab param; `","` → Comma; `";"` → Semicolon; `" "` → Space; else Other + `OtherChar` |
| **Return** | `@{ status='ok'; range=$Range; delimiter=$Delimiter }` |

---

### 4.2 `Invoke-ExcelAutoFill` (add to `WorksheetOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `SourceRange` (Mandatory) · `DestinationRange` (Mandatory) · `FillType` (ValidateSet: `'Default','Copy','Series','FillDays','FillWeekdays','FillMonths','FillYears','LinearTrend','GrowthTrend'`) · `AsJson` |
| **New constant** | `$script:XL_AUTO_FILL_TYPE` (see §5) |
| **COM** | `$srcRng.AutoFill($ws.Range($DestinationRange), [int]$script:XL_AUTO_FILL_TYPE[$FillType.ToLower()])` |
| **Return** | `@{ status='ok'; source=$SourceRange; destination=$DestinationRange; fillType=$FillType }` |

---

### 4.3 `Add-ExcelSubtotal` (add to `FilterSortOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory) · `GroupByColumn` (Mandatory, int — 1-based) · `Function` (Mandatory, ValidateSet: `'Sum','Count','Average','Max','Min','Product','CountNums','StdDev','Var'`) · `TotalColumns` (Mandatory, int array — 1-based) · `Replace` (switch — replace existing subtotals) · `AsJson` |
| **New constant** | `$script:XL_SUBTOTAL_FUNC` (see §5) |
| **COM** | `$rng.Subtotal($GroupByColumn, [int]$xlFunc, $TotalColumns, $Replace, $false, $true)` |
| **Return** | `@{ status='ok'; range=$Range; groupByColumn=$GroupByColumn; function=$Function }` |

---

### 4.4 `Set-ExcelSheetTab` (add to `WorksheetOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Color` (Mandatory, string — hex e.g. `"#FF0000"`) · `AsJson` |
| **COM** | `$ws.Tab.Color = ConvertTo-RGBColor $Color` (reuse existing private helper) |
| **Return** | `@{ status='ok'; sheet=$SheetName; tabColor=$Color }` |

---

### 4.5 `Get-ExcelFormulaDependencies` (add to `WorksheetOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `Range` (Mandatory) · `Direction` (Mandatory, ValidateSet: `'Precedents','Dependents'`) · `AsJson` |
| **COM** | If `Precedents`: `$rng.Precedents.Address($false,$false)`; if `Dependents`: `$rng.Dependents.Address($false,$false)` |
| **Edge case** | Throws if cell has no precedents/dependents — wrap in `try/catch`, return empty |
| **Return** | `@{ status='ok'; range=$Range; direction=$Direction; addresses=@(...); count=... }` |

---

### 4.6 Style Functions (new `StyleOps.ps1` — 3 functions)

| Function | Purpose | Parameters |
|---|---|---|
| `Get-ExcelStyle` | List workbook styles | `WorkbookPath`, `Name` (optional filter), `AsJson` |
| `New-ExcelStyle` | Create a named style | `WorkbookPath`, `Name`, `BasedOn` (optional), font/fill/border/number format params, `AsJson` |
| `Set-ExcelRangeStyle` | Apply a named style to a range | `WorkbookPath`, `SheetName`, `Range`, `StyleName`, `AsJson` |

**COM:**
- `$wb.Styles` collection — iterate for `Get-`, `.Add()` for `New-`, `$rng.Style = $name` for `Set-`.

---

### 4.7 `Export-ExcelToPdf` Refactor

Add `Format` parameter (`[ValidateSet('PDF','XPS')]`, default `'PDF'`) to existing `Export-ExcelToPdf`. Use `$script:XL_FIXED_FORMAT[$Format.ToLower()]` (already defined).

**Backward compatible:** New parameter has a default; existing callers are unaffected.

---

### 4.8 `Invoke-ExcelPrint` (add to `PrintOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `From` (int — first page) · `To` (int — last page) · `Copies` (int, default 1) · `Printer` (string — printer name) · `AsJson` |
| **COM** | `$ws.PrintOut($From, $To, $Copies, $false, $Printer)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; copies=$Copies }` |

---

### 4.9 `Set-ExcelStatusBar` (add to `MetadataOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `Text` (string) · `Reset` (switch — restores default) · `AsJson` |
| **COM** | If `$Reset`: `$app.StatusBar = $false`; else `$app.StatusBar = $Text` |
| **Return** | `@{ status='ok'; statusBar=($Text or 'reset') }` |

---

### 4.10 `Import-ExcelRecordset` (add to `ImportOps.ps1`)

| Attribute | Detail |
|---|---|
| **Parameters** | `WorkbookPath` (Mandatory) · `SheetName` (Mandatory) · `ConnectionString` (Mandatory) · `Query` (Mandatory — SQL string) · `StartCell` (string, default `"A1"`) · `AsJson` |
| **COM** | Create `ADODB.Recordset` → `.Open($Query, $ConnectionString)` → `$ws.Range($StartCell).CopyFromRecordset($rs)` |
| **Return** | `@{ status='ok'; sheet=$SheetName; startCell=$StartCell; rowsCopied=$rs.RecordCount }` |

**Security note:** `ConnectionString` is caller-provided. Do not construct connection strings from untrusted input. The function passes the string through as-is — the responsibility is on the caller.

---

## 5. New Constants

Add the following to `ExcelPOSH.psm1` in the CONSTANTS section, after the existing `$script:XL_YES_NO_GUESS` block:

```powershell
# xlCalculation — calculation mode
$script:XL_CALCULATION = @{
    automatic     = -4105  # xlCalculationAutomatic
    manual        = -4135  # xlCalculationManual
    semiautomatic = 2      # xlCalculationSemiautomatic
}

# xlCellType — SpecialCells type
$script:XL_CELL_TYPE = @{
    blanks    = 4      # xlCellTypeBlanks
    constants = 2      # xlCellTypeConstants
    formulas  = -4123  # xlCellTypeFormulas
    lastcell  = 11     # xlCellTypeLastCell
    visible   = 12     # xlCellTypeVisible
    comments  = -4144  # xlCellTypeComments
}

# xlFilterAction — AdvancedFilter action
$script:XL_FILTER_ACTION = @{
    filterinplace = 1  # xlFilterInPlace
    copytorange   = 2  # xlFilterCopy
}

# xlAutoFillType — AutoFill type
$script:XL_AUTO_FILL_TYPE = @{
    default      = 0   # xlFillDefault
    copy         = 1   # xlFillCopy
    series       = 2   # xlFillSeries
    filldays     = 5   # xlFillDays
    fillweekdays = 6   # xlFillWeekdays
    fillmonths   = 7   # xlFillMonths
    fillyears    = 8   # xlFillYears
    lineartrend  = 3   # xlLinearTrend
    growthtrend  = 4   # xlGrowthTrend
}

# xlAxisType — chart axis type
$script:XL_AXIS_TYPE = @{
    category   = 1  # xlCategory
    value      = 2  # xlValue
    seriesaxis = 3  # xlSeriesAxis
}

# xlAxisGroup — chart axis group
$script:XL_AXIS_GROUP = @{
    primary   = 1  # xlPrimary
    secondary = 2  # xlSecondary
}

# xlLegendPosition — chart legend position
$script:XL_LEGEND_POSITION = @{
    bottom = -4107  # xlLegendPositionBottom
    corner = 2      # xlLegendPositionCorner
    left   = -4131  # xlLegendPositionLeft
    right  = -4152  # xlLegendPositionRight
    top    = -4160  # xlLegendPositionTop
}

# xlConsolidationFunction — Subtotal function
$script:XL_SUBTOTAL_FUNC = @{
    sum       = 9    # xlSum
    count     = 2    # xlCount
    average   = 1    # xlAverage
    max       = 4    # xlMax
    min       = 5    # xlMin
    product   = 6    # xlProduct
    countnums = 3    # xlCountNums
    stddev    = 7    # xlStDev
    var       = 8    # xlVar
}
```

---

## 6. Manifest Updates

Update `ExcelPOSH.psd1`:

1. **ModuleVersion** → `'3.0.0'`
2. **Description** → update function count
3. **FunctionsToExport** — append new functions grouped by domain:

```powershell
        # Structural (4) — NEW
        'Add-ExcelRow'
        'Remove-ExcelRow'
        'Add-ExcelColumn'
        'Remove-ExcelColumn'

        # Calculation (4) — NEW
        'Set-ExcelPerformanceMode'
        'Invoke-ExcelCalculate'
        'Invoke-ExcelFunction'
        'Invoke-ExcelEvaluate'

        # Filter & Sort — additions
        'Remove-ExcelDuplicates'
        'Invoke-ExcelAdvancedFilter'
        'Add-ExcelSubtotal'

        # Formatting — additions
        'Merge-ExcelRange'
        'Split-ExcelRange'

        # Worksheet — additions
        'Get-ExcelSpecialCells'
        'Invoke-ExcelAutoFill'
        'Set-ExcelSheetTab'
        'Get-ExcelFormulaDependencies'

        # Workbook — additions
        'Invoke-ExcelMacro'

        # Chart — additions
        'Set-ExcelChartSeries'
        'Set-ExcelChartAxis'
        'Set-ExcelChartLegend'
        'Set-ExcelChartDataLabels'

        # Pivot Table — additions
        'Set-ExcelPivotField'
        'Add-ExcelPivotCalculatedField'

        # Import — additions
        'Split-ExcelColumn'
        'Import-ExcelRecordset'

        # Print — additions
        'Invoke-ExcelPrint'

        # Metadata — additions
        'Set-ExcelStatusBar'

        # Style (3) — NEW
        'Get-ExcelStyle'
        'New-ExcelStyle'
        'Set-ExcelRangeStyle'
```

---

## 7. File Summary

| File | Status | Functions Added |
|---|---|---|
| `Public/StructuralOps.ps1` | **NEW** | `Add-ExcelRow`, `Remove-ExcelRow`, `Add-ExcelColumn`, `Remove-ExcelColumn` |
| `Public/CalculationOps.ps1` | **NEW** | `Set-ExcelPerformanceMode`, `Invoke-ExcelCalculate`, `Invoke-ExcelFunction`, `Invoke-ExcelEvaluate` |
| `Public/StyleOps.ps1` | **NEW** | `Get-ExcelStyle`, `New-ExcelStyle`, `Set-ExcelRangeStyle` |
| `Public/FormattingOps.ps1` | Modified | + `Merge-ExcelRange`, `Split-ExcelRange` |
| `Public/WorksheetOps.ps1` | Modified | + `Get-ExcelSpecialCells`, `Invoke-ExcelAutoFill`, `Set-ExcelSheetTab`, `Get-ExcelFormulaDependencies` |
| `Public/FilterSortOps.ps1` | Modified | + `Remove-ExcelDuplicates`, `Invoke-ExcelAdvancedFilter`, `Add-ExcelSubtotal`; refactor `Sort-ExcelRange` |
| `Public/ChartOps.ps1` | Modified | + `Set-ExcelChartSeries`, `Set-ExcelChartAxis`, `Set-ExcelChartLegend`, `Set-ExcelChartDataLabels` |
| `Public/PivotTableOps.ps1` | Modified | + `Set-ExcelPivotField`, `Add-ExcelPivotCalculatedField` |
| `Public/WorkbookOps.ps1` | Modified | + `Invoke-ExcelMacro` |
| `Public/ImportOps.ps1` | Modified | + `Split-ExcelColumn`, `Import-ExcelRecordset` |
| `Public/PrintOps.ps1` | Modified | + `Invoke-ExcelPrint`; refactor `Export-ExcelToPdf` |
| `Public/MetadataOps.ps1` | Modified | + `Set-ExcelStatusBar` |
| `ExcelPOSH.psm1` | Modified | + 9 new `$script:XL_*` constant hashtables |
| `ExcelPOSH.psd1` | Modified | + 30 new `FunctionsToExport` entries; version → 3.0.0 |
| `Tests/StructuralOps.Tests.ps1` | **NEW** | Tests for 4 structural functions |
| `Tests/CalculationOps.Tests.ps1` | **NEW** | Tests for 4 calculation functions |
| `Tests/StyleOps.Tests.ps1` | **NEW** | Tests for 3 style functions |
| `Tests/FilterSortOps.Tests.ps1` | Modified | + tests for `Remove-ExcelDuplicates`, `Invoke-ExcelAdvancedFilter`, `Add-ExcelSubtotal`, `Sort-ExcelRange` refactor |
| `Tests/FormattingOps.Tests.ps1` | Modified | + tests for `Merge-ExcelRange`, `Split-ExcelRange` |
| `Tests/WorksheetOps.Tests.ps1` | Modified | + tests for `Get-ExcelSpecialCells`, `Invoke-ExcelAutoFill`, `Set-ExcelSheetTab`, `Get-ExcelFormulaDependencies` |
| `Tests/WorkbookOps.Tests.ps1` | Modified | + tests for `Invoke-ExcelMacro` |
| `Tests/ChartOps.Tests.ps1` | Modified | + tests for 4 chart functions |
| `Tests/PivotTableOps.Tests.ps1` | Modified | + tests for 2 pivot table functions |
| `Tests/ImportOps.Tests.ps1` | Modified | + tests for `Split-ExcelColumn`, `Import-ExcelRecordset` |
| `Tests/PrintOps.Tests.ps1` | Modified | + tests for `Invoke-ExcelPrint` |
| `Tests/MetadataOps.Tests.ps1` | Modified | + tests for `Set-ExcelStatusBar` |

---

## 8. Updated Function Count

| Milestone | Functions | Total |
|---|---|---|
| Current (v2.0.0) | 84 manifest / 80 confirmed source | 84 |
| Phase 1 | +12 (Structural 4, CalculationOps 4, FilterSort 1, Formatting 2, Worksheet 1) | 96 |
| Phase 2 | +8 (FilterSort refactor+1, Chart 4, Pivot 2, Workbook 1) | 104 |
| Phase 3 | +10 (Import 2, Worksheet 3, FilterSort 1, Style 3, Print 1, Metadata 1) | 114 |
| **v3.0.0 Target** | | **114** |

---

## 9. Implementation Order & Dependencies

Recommended implementation order within each phase, based on dependencies:

### Phase 1 (do first — unblocks agent workflows)

```
1. Add constants to ExcelPOSH.psm1         ← prerequisite for everything
2. StructuralOps.ps1 + tests               ← no internal deps
3. CalculationOps.ps1 + tests              ← needs XL_CALCULATION constant
4. Remove-ExcelDuplicates + tests          ← uses existing XL_YES_NO_GUESS
5. Merge-ExcelRange / Split-ExcelRange     ← no deps
6. Get-ExcelSpecialCells                   ← needs XL_CELL_TYPE constant
7. Invoke-ExcelMacro                       ← no deps
8. Update manifest (FunctionsToExport)
9. Run full Pester suite
```

### Phase 2 (charting / pivot refinement)

```
1. Sort-ExcelRange refactor                ← uses existing XL_SORT_ORDER
2. Invoke-ExcelAdvancedFilter              ← needs XL_FILTER_ACTION constant
3. Chart functions (4)                     ← needs XL_AXIS_TYPE, XL_AXIS_GROUP, XL_LEGEND_POSITION
4. Pivot table functions (2)               ← no new constants
5. Update manifest + run Pester
```

### Phase 3 (medium priority fill-ins)

```
1. Invoke-ExcelAutoFill                    ← needs XL_AUTO_FILL_TYPE
2. Add-ExcelSubtotal                       ← needs XL_SUBTOTAL_FUNC
3. Split-ExcelColumn                       ← complex TextToColumns mapping
4. Set-ExcelSheetTab                       ← uses existing ConvertTo-RGBColor
5. Get-ExcelFormulaDependencies            ← no deps
6. Style functions (3)                     ← new file
7. Export-ExcelToPdf refactor              ← uses existing XL_FIXED_FORMAT
8. Invoke-ExcelPrint                       ← no deps
9. Set-ExcelStatusBar                      ← no deps
10. Import-ExcelRecordset                  ← ADODB, test last
11. Update manifest → v3.0.0 + full Pester
```

### Validation checklist per function

- [ ] Comment-based help (`.SYNOPSIS`, `.PARAMETER`, `.EXAMPLE`)
- [ ] `[CmdletBinding()]` with `[Parameter(Mandatory)]` on required params
- [ ] `ValidateSet` / `ValidateRange` / `ValidateCount` where applicable
- [ ] Calls `Connect-ExcelWorkbook` first
- [ ] Returns `@{ status='ok'; ... }` through `Format-ExcelOutput`
- [ ] Constant lookups use `[int]$script:XL_*[$key]`
- [ ] Pester test file exists with `Describe` per function
- [ ] Function listed in `FunctionsToExport` in `.psd1`
- [ ] No reserved-word identifiers (per VBA naming guardrails)
