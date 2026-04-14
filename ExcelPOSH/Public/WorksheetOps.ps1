# Public/WorksheetOps.ps1 — Worksheet management and cell/range operations

function Get-ExcelWorksheet {
    <#
    .SYNOPSIS
        List all worksheets in the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
        $ur = $ws.UsedRange
        $sheets += @{
            name       = $ws.Name
            index      = $ws.Index
            visible    = $ws.Visible -eq -1
            usedRange  = $ur.Address($false, $false)
            rowCount   = $ur.Rows.Count
            colCount   = $ur.Columns.Count
        }
    }

    $result = @{
        status = 'ok'
        sheets = $sheets
        count  = $sheets.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelWorksheet {
    <#
    .SYNOPSIS
        Add a new worksheet to the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name for the new worksheet.
    .PARAMETER After
        Name of sheet to insert after. If omitted, adds at the end.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -SheetName "Summary" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$After,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    if ([string]::IsNullOrWhiteSpace($After)) {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    } else {
        $afterSheet = $wb.Worksheets.Item($After)
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $afterSheet)
    }
    $ws.Name = $SheetName

    $result = @{
        status = 'created'
        sheet  = $SheetName
        index  = $ws.Index
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelWorksheet {
    <#
    .SYNOPSIS
        Delete a worksheet from the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet to delete.
    .PARAMETER Confirm
        Must be $true to confirm deletion.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -SheetName "OldSheet" -Confirm:$true -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    if ($wb.Worksheets.Count -le 1) {
        throw "Cannot delete the only worksheet in the workbook."
    }

    $ws = $wb.Worksheets.Item($SheetName)
    $ws.Delete()

    $result = @{
        status  = 'deleted'
        sheet   = $SheetName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Rename-ExcelWorksheet {
    <#
    .SYNOPSIS
        Rename a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Current name of the worksheet.
    .PARAMETER NewName
        New name for the worksheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Rename-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -NewName "Data" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$NewName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $ws.Name = $NewName

    $result = @{
        status  = 'renamed'
        oldName = $SheetName
        newName = $NewName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Copy-ExcelWorksheet {
    <#
    .SYNOPSIS
        Copy a worksheet within the same workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet to copy.
    .PARAMETER NewName
        Name for the copied worksheet.
    .PARAMETER After
        Name of sheet to place the copy after. Defaults to last sheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -SheetName "Template" -NewName "Jan2026" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$NewName,
        [string]$After,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if ([string]::IsNullOrWhiteSpace($After)) {
        $ws.Copy([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    } else {
        $ws.Copy([System.Reflection.Missing]::Value, $wb.Worksheets.Item($After))
    }

    # The new sheet is the active sheet after copy
    $newSheet = $wb.ActiveSheet
    if (-not [string]::IsNullOrWhiteSpace($NewName)) {
        $newSheet.Name = $NewName
    }

    $result = @{
        status  = 'copied'
        source  = $SheetName
        newName = $newSheet.Name
        index   = $newSheet.Index
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Move-ExcelWorksheet {
    <#
    .SYNOPSIS
        Move a worksheet to a different position in the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet to move.
    .PARAMETER Before
        Name of sheet to move before.
    .PARAMETER After
        Name of sheet to move after.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Move-ExcelWorksheet -WorkbookPath "C:\data.xlsx" -SheetName "Summary" -Before "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Before,
        [string]$After,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if (-not [string]::IsNullOrWhiteSpace($Before)) {
        $ws.Move($wb.Worksheets.Item($Before))
    } elseif (-not [string]::IsNullOrWhiteSpace($After)) {
        $ws.Move([System.Reflection.Missing]::Value, $wb.Worksheets.Item($After))
    } else {
        throw "Specify either -Before or -After parameter."
    }

    $result = @{
        status   = 'moved'
        sheet    = $SheetName
        newIndex = $ws.Index
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelRange {
    <#
    .SYNOPSIS
        Read values from a cell or range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Cell or range address (e.g. "A1", "A1:D10", "B:B").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D10" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $rows = $rng.Rows.Count
    $cols = $rng.Columns.Count
    $addr = $rng.Address($false, $false)

    # Single cell
    if ($rows -eq 1 -and $cols -eq 1) {
        $result = @{
            status  = 'ok'
            range   = $addr
            value   = (ConvertTo-ExcelSafeValue $rng.Value2)
            formula = $rng.Formula
        }
        Format-ExcelOutput -Data $result -AsJson:$AsJson
        return
    }

    # Multi-cell: Value2 returns a 1-based 2D array for COM
    $raw = $rng.Value2
    $data = @()
    for ($r = 1; $r -le $rows; $r++) {
        $row = @()
        for ($c = 1; $c -le $cols; $c++) {
            $row += (ConvertTo-ExcelSafeValue $raw[$r, $c])
        }
        $data += , $row
    }

    $result = @{
        status  = 'ok'
        range   = $addr
        rows    = $rows
        columns = $cols
        data    = $data
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelRange {
    <#
    .SYNOPSIS
        Write values to a cell or range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Cell or range address (e.g. "A1", "A1:D3").
    .PARAMETER Value
        Value to write. Scalar for single cell, 2D array @(@(1,2),@(3,4)) for range.
    .PARAMETER Formula
        Formula to write (e.g. "=SUM(A1:A10)"). Overrides -Value.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1" -Value "Hello" -AsJson
    .EXAMPLE
        Set-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:B2" -Value @(@(1,2),@(3,4)) -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        $Value,
        [string]$Formula,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    if (-not [string]::IsNullOrWhiteSpace($Formula)) {
        $rng.Formula = $Formula
    } elseif ($null -ne $Value) {
        if ($Value -is [array] -and $Value.Count -gt 0 -and $Value[0] -is [array]) {
            # 2D array — convert to COM-compatible 2D array
            $rows = $Value.Count
            $cols = $Value[0].Count
            $comArray = [System.Array]::CreateInstance([object], @($rows, $cols))
            for ($r = 0; $r -lt $rows; $r++) {
                for ($c = 0; $c -lt $cols; $c++) {
                    $comArray[$r, $c] = $Value[$r][$c]
                }
            }
            $rng.Value2 = $comArray
        } else {
            $rng.Value2 = $Value
        }
    }

    $result = @{
        status = 'ok'
        range  = $rng.Address($false, $false)
        rows   = $rng.Rows.Count
        cols   = $rng.Columns.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Clear-ExcelRange {
    <#
    .SYNOPSIS
        Clear cells in a range (contents, formats, comments, or all).
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Cell or range address to clear.
    .PARAMETER ClearType
        What to clear: all, contents, formats, comments. Default: contents.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Clear-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D10" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [ValidateSet('all','contents','formats','comments')]
        [string]$ClearType = 'contents',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $method = $script:XL_CLEAR_TYPE[$ClearType]
    $rng.$method()

    $result = @{
        status    = 'cleared'
        range     = $rng.Address($false, $false)
        clearType = $ClearType
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelUsedRange {
    <#
    .SYNOPSIS
        Get the used range info for a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelUsedRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
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
    $ur  = $ws.UsedRange

    $result = @{
        status    = 'ok'
        sheet     = $SheetName
        address   = $ur.Address($false, $false)
        firstRow  = $ur.Row
        firstCol  = $ur.Column
        rows      = $ur.Rows.Count
        columns   = $ur.Columns.Count
        cellCount = $ur.Rows.Count * $ur.Columns.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Find-ExcelValue {
    <#
    .SYNOPSIS
        Search for a value in a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER SearchText
        Text or value to search for.
    .PARAMETER LookIn
        Where to look: values, formulas, comments. Default: values.
    .PARAMETER MatchCase
        Case-sensitive search.
    .PARAMETER MaxResults
        Maximum number of matches to return. Default: 50.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-ExcelValue -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SearchText "Error" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SearchText,
        [ValidateSet('values','formulas','comments')]
        [string]$LookIn = 'values',
        [switch]$MatchCase,
        [int]$MaxResults = 50,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    # xlValues=-4163, xlFormulas=-4123, xlComments=-4144
    $lookInConst = switch ($LookIn) {
        'values'   { -4163 }
        'formulas' { -4123 }
        'comments' { -4144 }
    }

    $matches = @()
    $found = $ws.Cells.Find($SearchText, [System.Reflection.Missing]::Value,
        $lookInConst, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        $MatchCase.IsPresent)

    if ($null -ne $found) {
        $firstAddr = $found.Address($false, $false)
        do {
            $matches += @{
                address = $found.Address($false, $false)
                value   = (ConvertTo-ExcelSafeValue $found.Value2)
                row     = $found.Row
                column  = $found.Column
            }
            if ($matches.Count -ge $MaxResults) { break }
            $found = $ws.Cells.FindNext($found)
        } while ($null -ne $found -and $found.Address($false, $false) -ne $firstAddr)
    }

    $result = @{
        status     = 'ok'
        searchText = $SearchText
        sheet      = $SheetName
        matchCount = $matches.Count
        matches    = $matches
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelNamedRange {
    <#
    .SYNOPSIS
        List named ranges in the workbook or a specific sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Optional: limit to names scoped to this sheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelNamedRange -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $names = @()
    foreach ($n in $wb.Names) {
        $entry = @{
            name    = $n.Name
            refersTo = $n.RefersTo
            visible = $n.Visible
        }
        # If filtering by sheet, check scope
        if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
            if ($n.RefersTo -notlike "*$SheetName*") { continue }
        }
        $names += $entry
    }

    $result = @{
        status = 'ok'
        count  = $names.Count
        names  = $names
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelNamedRange {
    <#
    .SYNOPSIS
        Create a named range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER Name
        Name for the range.
    .PARAMETER SheetName
        Worksheet containing the range.
    .PARAMETER Range
        Cell or range address (e.g. "A1:D10").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelNamedRange -WorkbookPath "C:\data.xlsx" -Name "SalesData" -SheetName "Sheet1" -Range "A1:D100" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $null = $wb.Names.Add($Name, $rng)

    $result = @{
        status   = 'created'
        name     = $Name
        refersTo = "=$SheetName!$($rng.Address())"
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelNamedRange {
    <#
    .SYNOPSIS
        Delete a named range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER Name
        Name of the range to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelNamedRange -WorkbookPath "C:\data.xlsx" -Name "OldRange" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $wb.Names.Item($Name).Delete()

    $result = @{
        status = 'deleted'
        name   = $Name
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

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
