# Public/ClipboardOps.ps1 — Copy/paste ranges and find & replace

function Copy-ExcelRange {
    <#
    .SYNOPSIS
        Copy a range to another location (values, formulas, or everything).
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Source worksheet name.
    .PARAMETER SourceRange
        Address of the range to copy (e.g. "A1:C10").
    .PARAMETER DestinationSheet
        Target worksheet name. Defaults to the same sheet as SourceRange.
    .PARAMETER DestinationRange
        Top-left cell to paste to (e.g. "E1").
    .PARAMETER PasteType
        What to paste: all, values, formats, or formulas. Default is 'all'.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SourceRange "A1:C10" -DestinationRange "E1" -AsJson
    .EXAMPLE
        Copy-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SourceRange "A1:C10" -DestinationSheet "Sheet2" -DestinationRange "A1" -PasteType values -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceRange,
        [string]$DestinationSheet,
        [Parameter(Mandatory)][string]$DestinationRange,
        [ValidateSet('all','values','formats','formulas')]
        [string]$PasteType = 'all',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $srcRng = $ws.Range($SourceRange)

    $destSheet = if ([string]::IsNullOrWhiteSpace($DestinationSheet)) {
        $ws
    } else {
        $wb.Worksheets.Item($DestinationSheet)
    }
    $destRng = $destSheet.Range($DestinationRange)

    if ($PasteType -eq 'all') {
        $srcRng.Copy($destRng)
    } else {
        $pasteConst = switch ($PasteType) {
            'values'   { -4163 }
            'formats'  { -4122 }
            'formulas' { -4123 }
        }
        $srcRng.Copy()
        $destRng.PasteSpecial($pasteConst)
    }

    $app.CutCopyMode = $false

    $result = @{
        status      = 'copied'
        source      = $SourceRange
        destination = $DestinationRange
        pasteType   = $PasteType
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Replace-ExcelValue {
    <#
    .SYNOPSIS
        Find and replace values in a range or entire used range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER SearchText
        Text to find.
    .PARAMETER ReplaceText
        Replacement text.
    .PARAMETER Range
        Optional range address to limit the replacement. Defaults to UsedRange.
    .PARAMETER MatchCase
        Perform a case-sensitive replacement.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Replace-ExcelValue -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SearchText "foo" -ReplaceText "bar" -AsJson
    .EXAMPLE
        Replace-ExcelValue -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SearchText "old" -ReplaceText "new" -Range "A1:D50" -MatchCase -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SearchText,
        [Parameter(Mandatory)][string]$ReplaceText,
        [string]$Range,
        [switch]$MatchCase,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $rng = if ([string]::IsNullOrWhiteSpace($Range)) {
        $ws.UsedRange
    } else {
        $ws.Range($Range)
    }

    $rng.Replace(
        $SearchText,
        $ReplaceText,
        [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value,
        $MatchCase.IsPresent
    )

    $rangeAddr = $rng.Address($false, $false)

    $result = @{
        status      = 'replaced'
        searchText  = $SearchText
        replaceText = $ReplaceText
        range       = $rangeAddr
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Move-ExcelRange {
    <#
    .SYNOPSIS
        Cut and paste a range to a new location.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Source worksheet name.
    .PARAMETER SourceRange
        Address of the range to cut (e.g. "A1:C10").
    .PARAMETER DestinationSheet
        Target worksheet name. Defaults to the same sheet as SourceRange.
    .PARAMETER DestinationRange
        Top-left cell to paste to (e.g. "E1").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Move-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SourceRange "A1:C10" -DestinationRange "E1" -AsJson
    .EXAMPLE
        Move-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -SourceRange "A1:C10" -DestinationSheet "Sheet2" -DestinationRange "A1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceRange,
        [string]$DestinationSheet,
        [Parameter(Mandatory)][string]$DestinationRange,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $srcRng = $ws.Range($SourceRange)

    $destSheet = if ([string]::IsNullOrWhiteSpace($DestinationSheet)) {
        $ws
    } else {
        $wb.Worksheets.Item($DestinationSheet)
    }
    $destRng = $destSheet.Range($DestinationRange)

    $srcRng.Cut($destRng)

    $result = @{
        status      = 'moved'
        source      = $SourceRange
        destination = $DestinationRange
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
