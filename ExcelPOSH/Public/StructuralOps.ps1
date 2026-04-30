# Public/StructuralOps.ps1 — Insert and delete rows/columns

function Add-ExcelRow {
    <#
    .SYNOPSIS
        Insert rows into a worksheet.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Row
        1-based row number where insertion begins.
    .PARAMETER Count
        Number of rows to insert (default 1).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelRow -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Row 5 -Count 3 -AsJson
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

function Remove-ExcelRow {
    <#
    .SYNOPSIS
        Delete rows from a worksheet.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Row
        1-based row number where deletion begins.
    .PARAMETER Count
        Number of rows to delete (default 1).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelRow -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Row 5 -Count 2 -AsJson
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

function Add-ExcelColumn {
    <#
    .SYNOPSIS
        Insert columns into a worksheet.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Column
        Column letter (e.g. "C") or 1-based integer.
    .PARAMETER Count
        Number of columns to insert (default 1).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelColumn -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Column "D" -Count 2 -AsJson
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

function Remove-ExcelColumn {
    <#
    .SYNOPSIS
        Delete columns from a worksheet.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Column
        Column letter (e.g. "C") or 1-based integer.
    .PARAMETER Count
        Number of columns to delete (default 1).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelColumn -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Column "B" -AsJson
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
