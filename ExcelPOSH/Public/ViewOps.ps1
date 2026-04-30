# Public/ViewOps.ps1 — Freeze panes, sheet visibility, outline/grouping

function Set-ExcelFreezePane {
    <#
    .SYNOPSIS
        Freeze or unfreeze panes at a specific cell position.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to freeze panes on.
    .PARAMETER Row
        Row to freeze above (e.g. 2 freezes row 1).
    .PARAMETER Column
        Column to freeze left of (e.g. 1 means no column freeze).
    .PARAMETER Unfreeze
        Remove freeze panes.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelFreezePane -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Row 2 -Column 1 -AsJson
    .EXAMPLE
        Set-ExcelFreezePane -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Unfreeze -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [int]$Row = 2,
        [int]$Column = 1,
        [switch]$Unfreeze,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $ws.Activate()

    if ($Unfreeze) {
        $app.ActiveWindow.FreezePanes = $false
        $app.ActiveWindow.SplitRow    = 0
        $app.ActiveWindow.SplitColumn = 0
    } else {
        $app.ActiveWindow.FreezePanes = $false
        $app.ActiveWindow.SplitRow    = $Row - 1
        $app.ActiveWindow.SplitColumn = $Column - 1
        $app.ActiveWindow.FreezePanes = $true
    }

    $result = @{
        status      = 'ok'
        sheet       = $SheetName
        frozen      = [bool]$app.ActiveWindow.FreezePanes
        splitRow    = [int]$app.ActiveWindow.SplitRow
        splitColumn = [int]$app.ActiveWindow.SplitColumn
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelFreezePane {
    <#
    .SYNOPSIS
        Get the current freeze pane state of a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to query.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelFreezePane -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $ws.Activate()

    $result = @{
        status      = 'ok'
        sheet       = $SheetName
        frozen      = [bool]$app.ActiveWindow.FreezePanes
        splitRow    = [int]$app.ActiveWindow.SplitRow
        splitColumn = [int]$app.ActiveWindow.SplitColumn
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelSheetVisibility {
    <#
    .SYNOPSIS
        Show or hide a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to change visibility for.
    .PARAMETER Visibility
        Visibility state: visible, hidden, or veryhidden.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelSheetVisibility -WorkbookPath "C:\data.xlsx" -SheetName "Sheet2" -Visibility hidden -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][ValidateSet('visible','hidden','veryhidden')][string]$Visibility,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $visMap = @{
        'visible'    = -1   # xlSheetVisible
        'hidden'     =  0   # xlSheetHidden
        'veryhidden' =  2   # xlSheetVeryHidden
    }
    $visConst = $visMap[$Visibility]
    $ws.Visible = $visConst

    $result = @{
        status     = 'set'
        sheet      = $SheetName
        visibility = $Visibility
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelSheetVisibility {
    <#
    .SYNOPSIS
        Get visibility status of all sheets in the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelSheetVisibility -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $visLabels = @{
        -1 = 'visible'
         0 = 'hidden'
         2 = 'veryhidden'
    }

    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
        $visVal = [int]$ws.Visible
        $sheets += @{
            name       = $ws.Name
            visibility = if ($visLabels.ContainsKey($visVal)) { $visLabels[$visVal] } else { "unknown($visVal)" }
        }
    }

    $result = @{
        status = 'ok'
        sheets = $sheets
        count  = $sheets.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelGrouping {
    <#
    .SYNOPSIS
        Group or ungroup rows or columns.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to group/ungroup on.
    .PARAMETER Range
        Range to group, e.g. "2:5" for rows, "B:D" for columns.
    .PARAMETER Ungroup
        Ungroup instead of group.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelGrouping -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "2:5" -AsJson
    .EXAMPLE
        Set-ExcelGrouping -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B:D" -Ungroup -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$Ungroup,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $rng    = $ws.Range($Range)
    $action = if ($Ungroup) { 'ungrouped' } else { 'grouped' }

    if ($Range -match '^\d+:\d+$') {
        # Row range
        if ($Ungroup) { $rng.Rows.Ungroup() } else { $rng.Rows.Group() }
    } elseif ($Range -match '^[A-Za-z]+:[A-Za-z]+$') {
        # Column range
        if ($Ungroup) { $rng.Columns.Ungroup() } else { $rng.Columns.Group() }
    } else {
        # General range — group the range itself
        if ($Ungroup) { $rng.Ungroup() } else { $rng.Group() }
    }

    $result = @{
        status = 'ok'
        range  = $Range
        action = $action
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelOutlineLevel {
    <#
    .SYNOPSIS
        Set the outline summary direction and show/collapse outline levels.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to configure.
    .PARAMETER SummaryBelow
        If specified, sets whether the summary row is below detail rows.
    .PARAMETER SummaryRight
        If specified, sets whether the summary column is to the right of detail columns.
    .PARAMETER RowLevel
        Show row outline to this level (1-8).
    .PARAMETER ColumnLevel
        Show column outline to this level (1-8).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelOutlineLevel -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SummaryBelow $true -RowLevel 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [System.Nullable[bool]]$SummaryBelow,
        [System.Nullable[bool]]$SummaryRight,
        [int]$RowLevel,
        [int]$ColumnLevel,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if ($null -ne $SummaryBelow) {
        # xlBelow = 1, xlAbove = 0
        $ws.Outline.SummaryRow = if ($SummaryBelow) { 1 } else { 0 }
    }

    if ($null -ne $SummaryRight) {
        # xlRight = 1, xlLeft = 0
        $ws.Outline.SummaryColumn = if ($SummaryRight) { 1 } else { 0 }
    }

    if ($RowLevel -gt 0 -or $ColumnLevel -gt 0) {
        $missing = [System.Reflection.Missing]::Value
        $rl = if ($RowLevel -gt 0) { $RowLevel } else { $missing }
        $cl = if ($ColumnLevel -gt 0) { $ColumnLevel } else { $missing }
        $ws.Outline.ShowLevels($rl, $cl)
    }

    $result = @{
        status = 'ok'
        sheet  = $SheetName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
