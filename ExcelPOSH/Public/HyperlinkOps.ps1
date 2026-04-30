# Public/HyperlinkOps.ps1 — Hyperlink operations

function Set-ExcelHyperlink {
    <#
    .SYNOPSIS
        Add or update a hyperlink on a cell.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Cell address to place the hyperlink (e.g. "A1").
    .PARAMETER Address
        URL or file path for the hyperlink.
    .PARAMETER SubAddress
        Optional internal link target (e.g. "Sheet2!A1").
    .PARAMETER TextToDisplay
        Optional display text for the hyperlink.
    .PARAMETER ScreenTip
        Optional screen tip shown on hover.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelHyperlink -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1" -Address "https://example.com" -TextToDisplay "Example" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$Address,
        [string]$SubAddress,
        [string]$TextToDisplay,
        [string]$ScreenTip,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $missing = [System.Reflection.Missing]::Value

    $subAddr   = if ([string]::IsNullOrWhiteSpace($SubAddress))    { $missing } else { $SubAddress }
    $tip       = if ([string]::IsNullOrWhiteSpace($ScreenTip))     { $missing } else { $ScreenTip }
    $display   = if ([string]::IsNullOrWhiteSpace($TextToDisplay)) { $missing } else { $TextToDisplay }

    $ws.Hyperlinks.Add($ws.Range($Range), $Address, $subAddr, $tip, $display)

    $result = @{
        status  = 'added'
        cell    = $Range
        address = $Address
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelHyperlink {
    <#
    .SYNOPSIS
        Get hyperlinks from a range or entire sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Optional cell or range address. If omitted, returns all hyperlinks on the sheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelHyperlink -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    .EXAMPLE
        Get-ExcelHyperlink -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:A10" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if ([string]::IsNullOrWhiteSpace($Range)) {
        $hyperlinks = $ws.Hyperlinks
    } else {
        $hyperlinks = $ws.Range($Range).Hyperlinks
    }

    $list = @()
    foreach ($h in $hyperlinks) {
        $list += @{
            cell          = $h.Range.Address($false, $false)
            address       = $h.Address
            subAddress    = $h.SubAddress
            textToDisplay = $h.TextToDisplay
            screenTip     = $h.ScreenTip
        }
    }

    $result = @{
        status     = 'ok'
        count      = $list.Count
        hyperlinks = $list
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelHyperlink {
    <#
    .SYNOPSIS
        Remove hyperlinks from a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Range
        Cell or range address to remove hyperlinks from.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelHyperlink -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1" -AsJson
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

    $ws.Range($Range).Hyperlinks.Delete()

    $result = @{
        status = 'removed'
        range  = $Range
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
