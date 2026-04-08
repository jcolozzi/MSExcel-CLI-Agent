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
