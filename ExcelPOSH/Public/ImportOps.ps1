# Public/ImportOps.ps1 — CSV and text file import

function Import-ExcelCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$CsvPath,
        [string]$StartCell = 'A1',
        [ValidateSet('comma','tab','semicolon','pipe','space')]
        [string]$Delimiter = 'comma',
        [switch]$HasHeaders,
        [ValidateSet('doubleQuote','singleQuote','none')]
        [string]$TextQualifier = 'doubleQuote',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }
    $fullPath = (Resolve-Path $CsvPath).Path

    $cell = $ws.Range($StartCell)
    $qt   = $ws.QueryTables.Add("TEXT;$fullPath", $cell)

    $qt.TextFileParseType = 1  # xlDelimited

    # Reset all delimiter flags
    $qt.TextFileCommaDelimiter     = $false
    $qt.TextFileTabDelimiter       = $false
    $qt.TextFileSemicolonDelimiter = $false
    $qt.TextFileSpaceDelimiter     = $false

    switch ($Delimiter) {
        'comma'     { $qt.TextFileCommaDelimiter     = $true }
        'tab'       { $qt.TextFileTabDelimiter       = $true }
        'semicolon' { $qt.TextFileSemicolonDelimiter = $true }
        'pipe'      { $qt.TextFileOtherDelimiter     = '|'   }
        'space'     { $qt.TextFileSpaceDelimiter     = $true }
    }

    $qualMap = @{ doubleQuote = 1; singleQuote = 2; none = 0 }
    $qt.TextFileTextQualifier = $qualMap[$TextQualifier]

    $qt.Refresh($false)  # synchronous

    $rowCount = $qt.ResultRange.Rows.Count
    $colCount = $qt.ResultRange.Columns.Count

    # Clean up query table connection to leave just the data
    $qt.Delete()

    $result = @{
        status     = 'imported'
        source     = $fullPath
        start_cell = $StartCell
        rows       = $rowCount
        columns    = $colCount
        delimiter  = $Delimiter
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Import-ExcelText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$TextPath,
        [string]$StartCell = 'A1',
        [ValidateSet('delimited','fixedWidth')]
        [string]$ParseType = 'delimited',
        [string]$Delimiter = ',',
        [int[]]$FieldWidths,
        [ValidateSet('doubleQuote','singleQuote','none')]
        [string]$TextQualifier = 'doubleQuote',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if (-not (Test-Path $TextPath)) {
        throw "Text file not found: $TextPath"
    }
    $fullPath = (Resolve-Path $TextPath).Path

    $cell = $ws.Range($StartCell)
    $qt   = $ws.QueryTables.Add("TEXT;$fullPath", $cell)

    switch ($ParseType) {
        'delimited' {
            $qt.TextFileParseType    = 1  # xlDelimited
            $qt.TextFileOtherDelimiter = $Delimiter
        }
        'fixedWidth' {
            $qt.TextFileParseType = 2  # xlFixedWidth
            if ($FieldWidths) {
                $qt.TextFileFixedWidthSettings = $FieldWidths
            }
        }
    }

    $qualMap = @{ doubleQuote = 1; singleQuote = 2; none = 0 }
    $qt.TextFileTextQualifier = $qualMap[$TextQualifier]

    $qt.Refresh($false)  # synchronous

    $rowCount = $qt.ResultRange.Rows.Count
    $colCount = $qt.ResultRange.Columns.Count

    # Clean up query table connection to leave just the data
    $qt.Delete()

    $result = @{
        status     = 'imported'
        source     = $fullPath
        start_cell = $StartCell
        rows       = $rowCount
        columns    = $colCount
        parse_type = $ParseType
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Split-ExcelColumn {
    <#
    .SYNOPSIS
        Split a column of delimited or fixed-width text into multiple columns (Text to Columns).
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Range
        Single-column source range to split (e.g. "A1:A100").
    .PARAMETER ParseType
        delimited or fixedWidth.
    .PARAMETER Delimiter
        Delimiter to split on when ParseType is delimited: tab, semicolon, comma, space, other.
    .PARAMETER OtherChar
        Custom delimiter character (used when Delimiter is 'other').
    .PARAMETER Destination
        Optional top-left cell for the result (defaults to in-place).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Split-ExcelColumn -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:A100" -Delimiter comma -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [ValidateSet('delimited','fixedWidth')][string]$ParseType = 'delimited',
        [ValidateSet('tab','semicolon','comma','space','other')][string]$Delimiter = 'comma',
        [string]$OtherChar,
        [string]$Destination,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $dataType = [int]$script:XL_TEXT_PARSE_TYPE[$ParseType.ToLower()]
    $dest = if ([string]::IsNullOrWhiteSpace($Destination)) { [System.Reflection.Missing]::Value } else { $ws.Range($Destination) }

    $tab = $semi = $comma = $space = $other = $false
    $otherCharVal = [System.Reflection.Missing]::Value
    switch ($Delimiter) {
        'tab'       { $tab = $true }
        'semicolon' { $semi = $true }
        'comma'     { $comma = $true }
        'space'     { $space = $true }
        'other'     { $other = $true; $otherCharVal = $OtherChar }
    }

    # TextToColumns(Destination, DataType, TextQualifier=1(xlDoubleQuote), ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar)
    $rng.TextToColumns($dest, $dataType, 1, $false, $tab, $semi, $comma, $space, $other, $otherCharVal)

    $result = @{
        status    = 'ok'
        range     = $Range
        parseType = $ParseType
        delimiter = $Delimiter
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Import-ExcelRecordset {
    <#
    .SYNOPSIS
        Run a SQL query and dump the results into a range via Range.CopyFromRecordset (ADO).
    .DESCRIPTION
        Fastest way to load database query results into Excel. Requires an ADO-compatible
        connection string. Returns status='error' if the connection or query fails.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Destination
        Top-left cell for the results (e.g. "A1").
    .PARAMETER ConnectionString
        ADO/OLEDB/ODBC connection string.
    .PARAMETER Query
        SQL SELECT statement.
    .PARAMETER MaxRows
        Optional row cap.
    .PARAMETER MaxColumns
        Optional column cap.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Import-ExcelRecordset -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Destination "A1" -ConnectionString "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\db.accdb" -Query "SELECT * FROM Customers" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$ConnectionString,
        [Parameter(Mandatory)][string]$Query,
        [int]$MaxRows,
        [int]$MaxColumns,
        [switch]$AsJson
    )

    $app  = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws   = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $cell = $ws.Range($Destination)

    $conn = $null
    $rs   = $null
    try {
        $conn = New-Object -ComObject 'ADODB.Connection'
        $rs   = New-Object -ComObject 'ADODB.Recordset'
        $conn.Open($ConnectionString)
        $rs.Open($Query, $conn)

        if ($MaxRows -gt 0 -and $MaxColumns -gt 0) {
            $copied = $cell.CopyFromRecordset($rs, $MaxRows, $MaxColumns)
        } elseif ($MaxRows -gt 0) {
            $copied = $cell.CopyFromRecordset($rs, $MaxRows)
        } else {
            $copied = $cell.CopyFromRecordset($rs)
        }

        $result = @{
            status      = 'ok'
            destination = $Destination
            rowsCopied  = $copied
        }
    } catch {
        $result = @{
            status = 'error'
            error  = "CopyFromRecordset failed: $($_.Exception.Message)"
        }
    } finally {
        if ($null -ne $rs)   { try { if ($rs.State -ne 0)   { $rs.Close() } }   catch {}; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs) }   catch {} }
        if ($null -ne $conn) { try { if ($conn.State -ne 0) { $conn.Close() } } catch {}; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn) } catch {} }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
