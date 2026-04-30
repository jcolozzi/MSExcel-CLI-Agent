# Public/PivotTableOps.ps1 — PivotTable create, query, refresh

function New-ExcelPivotTable {
    <#
    .SYNOPSIS
        Create a pivot table from a source data range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the source data.
    .PARAMETER SourceRange
        Range address of the source data (e.g. "A1:D100").
    .PARAMETER PivotTableName
        Name for the new pivot table.
    .PARAMETER DestinationSheet
        Name of the target sheet. Will be created if it does not exist.
    .PARAMETER DestinationCell
        Cell address on the destination sheet (default "A1").
    .PARAMETER RowFields
        Field names to place in rows.
    .PARAMETER ColumnFields
        Field names to place in columns.
    .PARAMETER DataFields
        Array of hashtables, each with Name (field name) and Function (Sum,Count,Average,Max,Min,Product,CountNumbers,StdDev,Var).
    .PARAMETER FilterFields
        Field names to place in the page/filter area.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelPivotTable -WorkbookPath "C:\data.xlsx" -SheetName "Sales" -SourceRange "A1:D100" -PivotTableName "SalesPivot" -DestinationSheet "PivotSheet" -RowFields @("Region") -DataFields @(@{Name="Revenue";Function="Sum"}) -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceRange,
        [Parameter(Mandatory)][string]$PivotTableName,
        [Parameter(Mandatory)][string]$DestinationSheet,
        [string]$DestinationCell = 'A1',
        [string[]]$RowFields,
        [string[]]$ColumnFields,
        [hashtable[]]$DataFields,
        [string[]]$FilterFields,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $sourceRangeObj = $ws.Range($SourceRange)

    # xlDatabase = 1
    $cache = $wb.PivotCaches().Create(1, $sourceRangeObj)

    # Get or create destination sheet
    $destWs = try {
        $wb.Worksheets.Item($DestinationSheet)
    } catch {
        $wb.Worksheets.Add()
        $wb.ActiveSheet.Name = $DestinationSheet
        $wb.ActiveSheet
    }

    $destCell = $destWs.Range($DestinationCell)
    $pt = $cache.CreatePivotTable($destCell, $PivotTableName)

    # Add row fields (xlRowField = 1)
    foreach ($f in $RowFields) {
        $pf = $pt.PivotFields($f)
        $pf.Orientation = 1
    }

    # Add column fields (xlColumnField = 2)
    foreach ($f in $ColumnFields) {
        $pf = $pt.PivotFields($f)
        $pf.Orientation = 2
    }

    # Add filter/page fields (xlPageField = 3)
    foreach ($f in $FilterFields) {
        $pf = $pt.PivotFields($f)
        $pf.Orientation = 3
    }

    # Add data fields
    $validFunctions = @('Sum','Count','Average','Max','Min','Product','CountNumbers','StdDev','Var')
    $funcMap = @{
        Sum          = -4157
        Count        = -4112
        Average      = -4106
        Max          = -4136
        Min          = -4139
        Product      = -4149
        CountNumbers = -4113
        StdDev       = -4155
        Var          = -4164
    }
    foreach ($df in $DataFields) {
        if ($df.Function -and $df.Function -notin $validFunctions) {
            throw "Invalid data field function '$($df.Function)'. Valid values: $($validFunctions -join ', ')"
        }
        $pf = $pt.AddDataField($pt.PivotFields($df.Name))
        if ($df.Function -and $funcMap.ContainsKey($df.Function)) {
            $pf.Function = $funcMap[$df.Function]
        }
    }

    $result = @{
        status            = 'created'
        name              = $PivotTableName
        destination_sheet = $DestinationSheet
        row_fields        = if ($RowFields) { $RowFields } else { @() }
        column_fields     = if ($ColumnFields) { $ColumnFields } else { @() }
        data_fields_count = if ($DataFields) { $DataFields.Count } else { 0 }
        filter_fields     = if ($FilterFields) { $FilterFields } else { @() }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelPivotTable {
    <#
    .SYNOPSIS
        Get info about pivot tables on a sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to inspect for pivot tables.
    .PARAMETER PivotTableName
        Optional: name of a specific pivot table. If omitted, lists all.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelPivotTable -WorkbookPath "C:\data.xlsx" -SheetName "PivotSheet" -AsJson
    .EXAMPLE
        Get-ExcelPivotTable -WorkbookPath "C:\data.xlsx" -SheetName "PivotSheet" -PivotTableName "SalesPivot" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$PivotTableName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $pivotTables = @()

    if ($PivotTableName) {
        $pt = $ws.PivotTables($PivotTableName)

        $rowFieldNames = @()
        foreach ($rf in $pt.RowFields) { $rowFieldNames += $rf.Name }

        $colFieldNames = @()
        foreach ($cf in $pt.ColumnFields) { $colFieldNames += $cf.Name }

        $dataFieldInfos = @()
        foreach ($df in $pt.DataFields) {
            $dataFieldInfos += @{
                name     = $df.Name
                function = $df.Function
            }
        }

        $pageFieldNames = @()
        foreach ($pf in $pt.PageFields) { $pageFieldNames += $pf.Name }

        $pivotTables += @{
            name          = $pt.Name
            source_data   = $pt.SourceData
            row_fields    = $rowFieldNames
            column_fields = $colFieldNames
            data_fields   = $dataFieldInfos
            page_fields   = $pageFieldNames
            table_range   = $pt.TableRange2.Address($false, $false)
        }
    } else {
        foreach ($pt in $ws.PivotTables()) {
            $rowFieldNames = @()
            foreach ($rf in $pt.RowFields) { $rowFieldNames += $rf.Name }

            $colFieldNames = @()
            foreach ($cf in $pt.ColumnFields) { $colFieldNames += $cf.Name }

            $dataFieldInfos = @()
            foreach ($df in $pt.DataFields) {
                $dataFieldInfos += @{
                    name     = $df.Name
                    function = $df.Function
                }
            }

            $pageFieldNames = @()
            foreach ($pf in $pt.PageFields) { $pageFieldNames += $pf.Name }

            $pivotTables += @{
                name          = $pt.Name
                source_data   = $pt.SourceData
                row_fields    = $rowFieldNames
                column_fields = $colFieldNames
                data_fields   = $dataFieldInfos
                page_fields   = $pageFieldNames
                table_range   = $pt.TableRange2.Address($false, $false)
            }
        }
    }

    $result = @{
        status       = 'ok'
        count        = $pivotTables.Count
        pivot_tables = $pivotTables
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Update-ExcelPivotTable {
    <#
    .SYNOPSIS
        Refresh a pivot table or all pivot tables on a sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the pivot table(s).
    .PARAMETER PivotTableName
        Optional: name of a specific pivot table to refresh. If omitted, refreshes all.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-ExcelPivotTable -WorkbookPath "C:\data.xlsx" -SheetName "PivotSheet" -AsJson
    .EXAMPLE
        Update-ExcelPivotTable -WorkbookPath "C:\data.xlsx" -SheetName "PivotSheet" -PivotTableName "SalesPivot" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$PivotTableName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $refreshed = 0

    if ($PivotTableName) {
        $ws.PivotTables($PivotTableName).RefreshTable()
        $refreshed = 1
    } else {
        foreach ($pt in $ws.PivotTables()) {
            $pt.RefreshTable()
            $refreshed++
        }
    }

    $result = @{
        status = 'refreshed'
        count  = $refreshed
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelPivotField {
    <#
    .SYNOPSIS  Configure pivot field properties (subtotals, number format, layout).
    .PARAMETER WorkbookPath    Path to the Excel workbook.
    .PARAMETER SheetName       Worksheet containing the pivot table.
    .PARAMETER PivotTableName  Name of the pivot table.
    .PARAMETER FieldName       Name of the pivot field to configure.
    .PARAMETER Subtotals       Subtotal functions to apply.
    .PARAMETER NumberFormat    Number format string.
    .PARAMETER LayoutForm      Layout form: Compact, Tabular, or Outline.
    .PARAMETER AsJson          Return JSON string.
    .EXAMPLE   Set-ExcelPivotField -WorkbookPath C:\data.xlsx -SheetName PivotSheet -PivotTableName "SalesPivot" -FieldName "Region" -LayoutForm Tabular -AsJson
    #>
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

function Add-ExcelPivotCalculatedField {
    <#
    .SYNOPSIS  Add a calculated field to a pivot table.
    .PARAMETER WorkbookPath    Path to the Excel workbook.
    .PARAMETER SheetName       Worksheet containing the pivot table.
    .PARAMETER PivotTableName  Name of the pivot table.
    .PARAMETER Name            Name for the calculated field.
    .PARAMETER Formula         Formula expression for the calculated field.
    .PARAMETER AsJson          Return JSON string.
    .EXAMPLE   Add-ExcelPivotCalculatedField -WorkbookPath C:\data.xlsx -SheetName PivotSheet -PivotTableName "SalesPivot" -Name "Profit" -Formula "=Revenue-Cost" -AsJson
    #>
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
