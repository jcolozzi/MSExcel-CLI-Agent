# Public/TableOps.ps1 — ListObject (Table) operations

function Get-ExcelTable {
    <#
    .SYNOPSIS
        List all tables (ListObjects) in a worksheet or workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Optional: limit to tables on this sheet. If omitted, lists all tables.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelTable -WorkbookPath "C:\data.xlsx" -AsJson
    .EXAMPLE
        Get-ExcelTable -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $tables = @()
    $sheetsToSearch = if ([string]::IsNullOrWhiteSpace($SheetName)) {
        $wb.Worksheets
    } else {
        @($wb.Worksheets.Item($SheetName))
    }

    foreach ($ws in $sheetsToSearch) {
        foreach ($lo in $ws.ListObjects) {
            $cols = @()
            foreach ($col in $lo.ListColumns) {
                $cols += $col.Name
            }
            $tables += @{
                name        = $lo.Name
                sheet       = $ws.Name
                range       = $lo.Range.Address($false, $false)
                dataRange   = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Address($false, $false) } else { $null }
                headerRange = $lo.HeaderRowRange.Address($false, $false)
                rowCount    = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
                columnCount = $lo.ListColumns.Count
                columns     = $cols
                showTotals  = $lo.ShowTotals
            }
        }
    }

    $result = @{
        status = 'ok'
        count  = $tables.Count
        tables = $tables
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelTable {
    <#
    .SYNOPSIS
        Create a table (ListObject) from a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the data.
    .PARAMETER Range
        Range address for the table (e.g. "A1:D10"). First row becomes headers.
    .PARAMETER TableName
        Name for the new table.
    .PARAMETER HasHeaders
        Whether the first row contains headers. Default: $true.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelTable -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D10" -TableName "Sales" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$TableName,
        [bool]$HasHeaders = $true,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $headerFlag = if ($HasHeaders) { $script:XL_YES_NO_GUESS['yes'] } else { $script:XL_YES_NO_GUESS['no'] }
    $lo = $ws.ListObjects.Add($script:XL_SOURCE_TYPE['range'], $rng, $null, $headerFlag)
    $lo.Name = $TableName

    $result = @{
        status      = 'created'
        tableName   = $lo.Name
        sheet       = $SheetName
        range       = $lo.Range.Address($false, $false)
        columnCount = $lo.ListColumns.Count
        rowCount    = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelTable {
    <#
    .SYNOPSIS
        Delete a table (ListObject). Optionally keep the data as a plain range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table to delete.
    .PARAMETER KeepData
        If $true, converts table to range (keeps data). If $false, deletes everything.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelTable -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -KeepData -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [switch]$KeepData,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    if ($KeepData) {
        $lo.Unlist()
        $action = 'converted_to_range'
    } else {
        $lo.Delete()
        $action = 'deleted'
    }

    $result = @{
        status    = $action
        tableName = $TableName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Resize-ExcelTable {
    <#
    .SYNOPSIS
        Resize a table to a new range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER NewRange
        New range address for the table (e.g. "A1:F20").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Resize-ExcelTable -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -NewRange "A1:F20" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$NewRange,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    $rng = $ws.Range($NewRange)
    $lo.Resize($rng)

    $result = @{
        status    = 'resized'
        tableName = $TableName
        newRange  = $lo.Range.Address($false, $false)
        rowCount  = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
        colCount  = $lo.ListColumns.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelTableData {
    <#
    .SYNOPSIS
        Read data from a table (ListObject).
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER Limit
        Maximum number of rows to return. Default: 100.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelTableData -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -Limit 50 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [int]$Limit = 100,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    # Get column names
    $colNames = @()
    foreach ($col in $lo.ListColumns) {
        $colNames += $col.Name
    }

    # Get data rows
    $rows = @()
    if ($null -ne $lo.DataBodyRange) {
        $totalRows = $lo.DataBodyRange.Rows.Count
        $totalCols = $lo.DataBodyRange.Columns.Count
        $rowsToRead = [Math]::Min($totalRows, $Limit)

        if ($rowsToRead -eq 1 -and $totalCols -eq 1) {
            $val = $lo.DataBodyRange.Value2
            $rowData = @{}
            $rowData[$colNames[0]] = (ConvertTo-ExcelSafeValue $val)
            $rows += $rowData
        } else {
            $readRange = $lo.DataBodyRange.Resize($rowsToRead, $totalCols)
            $raw = $readRange.Value2
            for ($r = 1; $r -le $rowsToRead; $r++) {
                $rowData = @{}
                for ($c = 1; $c -le $totalCols; $c++) {
                    $rowData[$colNames[$c - 1]] = (ConvertTo-ExcelSafeValue $raw[$r, $c])
                }
                $rows += $rowData
            }
        }
    }

    $result = @{
        status     = 'ok'
        tableName  = $TableName
        columns    = $colNames
        rowCount   = $rows.Count
        totalRows  = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
        data       = $rows
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelTableRow {
    <#
    .SYNOPSIS
        Add one or more rows to a table.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER Rows
        Array of hashtables, each mapping column name to value.
        Example: @(@{Name="Alice"; Age=30}, @{Name="Bob"; Age=25})
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelTableRow -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Employees" -Rows @(@{Name="Alice";Age=30}) -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][hashtable[]]$Rows,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    # Build column name → index map
    $colMap = @{}
    foreach ($col in $lo.ListColumns) {
        $colMap[$col.Name] = $col.Index
    }

    $addedCount = 0
    foreach ($row in $Rows) {
        $lr = $lo.ListRows.Add()
        foreach ($key in $row.Keys) {
            if ($colMap.ContainsKey($key)) {
                $colIdx = $colMap[$key]
                $lr.Range.Cells.Item(1, $colIdx).Value2 = $row[$key]
            }
        }
        $addedCount++
    }

    $result = @{
        status    = 'added'
        tableName = $TableName
        rowsAdded = $addedCount
        totalRows = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelTableRow {
    <#
    .SYNOPSIS
        Delete rows from a table by row index (1-based).
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER RowIndex
        1-based row index or array of indexes to delete. Deleted from bottom up.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelTableRow -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -RowIndex @(3,5) -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][int[]]$RowIndex,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    # Delete from bottom to top to preserve indexes
    $sorted = $RowIndex | Sort-Object -Descending
    $deletedCount = 0
    foreach ($idx in $sorted) {
        if ($idx -ge 1 -and $idx -le $lo.ListRows.Count) {
            $lo.ListRows.Item($idx).Delete()
            $deletedCount++
        }
    }

    $result = @{
        status      = 'deleted'
        tableName   = $TableName
        rowsDeleted = $deletedCount
        totalRows   = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelTableColumn {
    <#
    .SYNOPSIS
        Get column info or column data from a table.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER ColumnName
        Optional: return data for this specific column. If omitted, lists all columns.
    .PARAMETER Limit
        Max data rows when reading a specific column. Default: 100.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelTableColumn -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -AsJson
    .EXAMPLE
        Get-ExcelTableColumn -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -ColumnName "Revenue" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [string]$ColumnName,
        [int]$Limit = 100,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        # List all columns
        $cols = @()
        foreach ($col in $lo.ListColumns) {
            $cols += @{
                name  = $col.Name
                index = $col.Index
                range = $col.Range.Address($false, $false)
            }
        }
        $result = @{
            status    = 'ok'
            tableName = $TableName
            count     = $cols.Count
            columns   = $cols
        }
    } else {
        # Read specific column data
        $col = $lo.ListColumns.Item($ColumnName)
        $data = @()
        if ($null -ne $lo.DataBodyRange) {
            $colRange = $col.DataBodyRange
            $totalRows = $colRange.Rows.Count
            $rowsToRead = [Math]::Min($totalRows, $Limit)

            if ($rowsToRead -eq 1) {
                $data += (ConvertTo-ExcelSafeValue $colRange.Value2)
            } else {
                $readRange = $colRange.Resize($rowsToRead, 1)
                $raw = $readRange.Value2
                for ($r = 1; $r -le $rowsToRead; $r++) {
                    $data += (ConvertTo-ExcelSafeValue $raw[$r, 1])
                }
            }
        }

        $result = @{
            status     = 'ok'
            tableName  = $TableName
            columnName = $ColumnName
            totalRows  = if ($null -ne $lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
            rowCount   = $data.Count
            data       = $data
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelTableTotals {
    <#
    .SYNOPSIS
        Configure the totals row for a table.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the table.
    .PARAMETER TableName
        Name of the table.
    .PARAMETER ShowTotals
        Show or hide the totals row.
    .PARAMETER Calculations
        Hashtable mapping column name to calculation type:
        none, average, count, countnums, max, min, sum, stddev, var.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelTableTotals -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -TableName "Sales" -ShowTotals $true -Calculations @{Revenue="sum"; Quantity="count"} -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TableName,
        [bool]$ShowTotals = $true,
        [hashtable]$Calculations,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $lo  = $ws.ListObjects.Item($TableName)

    $lo.ShowTotals = $ShowTotals

    if ($ShowTotals -and $null -ne $Calculations) {
        foreach ($colName in $Calculations.Keys) {
            $calcType = $Calculations[$colName].ToLower()
            if ($script:XL_TOTALS_CALC.ContainsKey($calcType)) {
                $col = $lo.ListColumns.Item($colName)
                $col.TotalsCalculation = $script:XL_TOTALS_CALC[$calcType]
            } else {
                Write-Warning "Unknown totals calculation '$calcType' for column '$colName'. Valid: $($script:XL_TOTALS_CALC.Keys -join ', ')"
            }
        }
    }

    $result = @{
        status     = 'ok'
        tableName  = $TableName
        showTotals = $lo.ShowTotals
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
