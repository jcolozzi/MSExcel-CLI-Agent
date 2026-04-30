# Public/FormattingOps.ps1 — Cell formatting: font, color, borders, number format, alignment

function Set-ExcelCellFormat {
    <#
    .SYNOPSIS
        Apply font and fill formatting to a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address.
    .PARAMETER Bold
        Set bold.
    .PARAMETER Italic
        Set italic.
    .PARAMETER Underline
        Set underline.
    .PARAMETER FontSize
        Font size in points.
    .PARAMETER FontName
        Font name (e.g. "Calibri", "Arial").
    .PARAMETER FontColor
        Font color as RGB integer or hex string (e.g. "#FF0000" for red).
    .PARAMETER FillColor
        Fill/background color as RGB integer or hex string.
    .PARAMETER FillPattern
        Fill pattern: none, solid, gray50, gray75, gray25. Default: solid (when FillColor set).
    .PARAMETER BorderStyle
        Border line style: continuous, dash, dot, double, none.
    .PARAMETER BorderWeight
        Border weight: hairline, thin, medium, thick.
    .PARAMETER BorderEdges
        Which borders to set: all, left, right, top, bottom, outline. Default: outline.
    .PARAMETER BorderColor
        Border color as RGB integer or hex string.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelCellFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D1" -Bold $true -FillColor "#4472C4" -FontColor "#FFFFFF" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Nullable[bool]]$Bold,
        [Nullable[bool]]$Italic,
        [Nullable[bool]]$Underline,
        [Nullable[double]]$FontSize,
        [string]$FontName,
        [string]$FontColor,
        [string]$FillColor,
        [string]$FillPattern,
        [string]$BorderStyle,
        [string]$BorderWeight,
        [ValidateSet('all','left','right','top','bottom','outline')]
        [string]$BorderEdges = 'outline',
        [string]$BorderColor,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $changes = @()

    # Font properties
    if ($null -ne $Bold)      { $rng.Font.Bold = $Bold; $changes += 'bold' }
    if ($null -ne $Italic)    { $rng.Font.Italic = $Italic; $changes += 'italic' }
    if ($null -ne $Underline) { $rng.Font.Underline = if ($Underline) { 2 } else { -4142 }; $changes += 'underline' }  # xlUnderlineStyleSingle=2, xlNone=-4142
    if ($null -ne $FontSize)  { $rng.Font.Size = $FontSize; $changes += 'fontSize' }
    if (-not [string]::IsNullOrWhiteSpace($FontName)) { $rng.Font.Name = $FontName; $changes += 'fontName' }

    # Font color
    if (-not [string]::IsNullOrWhiteSpace($FontColor)) {
        $rng.Font.Color = (ConvertTo-RGBColor $FontColor)
        $changes += 'fontColor'
    }

    # Fill
    if (-not [string]::IsNullOrWhiteSpace($FillColor)) {
        $patternKey = if ([string]::IsNullOrWhiteSpace($FillPattern)) { 'solid' } else { $FillPattern.ToLower() }
        $rng.Interior.Pattern = $script:XL_PATTERN[$patternKey]
        $rng.Interior.Color = (ConvertTo-RGBColor $FillColor)
        $changes += 'fill'
    }

    # Borders
    if (-not [string]::IsNullOrWhiteSpace($BorderStyle) -or -not [string]::IsNullOrWhiteSpace($BorderWeight)) {
        $edgeKeys = switch ($BorderEdges) {
            'all'     { @('left','right','top','bottom','insideh','insidev') }
            'outline' { @('left','right','top','bottom') }
            default   { @($BorderEdges) }
        }
        foreach ($edge in $edgeKeys) {
            $borderIdx = $script:XL_BORDER_INDEX[$edge]
            $border = $rng.Borders.Item($borderIdx)
            if (-not [string]::IsNullOrWhiteSpace($BorderStyle)) {
                $border.LineStyle = $script:XL_LINE_STYLE[$BorderStyle.ToLower()]
            }
            if (-not [string]::IsNullOrWhiteSpace($BorderWeight)) {
                $border.Weight = $script:XL_BORDER_WEIGHT[$BorderWeight.ToLower()]
            }
            if (-not [string]::IsNullOrWhiteSpace($BorderColor)) {
                $border.Color = (ConvertTo-RGBColor $BorderColor)
            }
        }
        $changes += 'borders'
    }

    $result = @{
        status  = 'formatted'
        range   = $rng.Address($false, $false)
        changes = $changes
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelNumberFormat {
    <#
    .SYNOPSIS
        Apply number formatting to a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address.
    .PARAMETER NumberFormat
        Excel number format string (e.g. "#,##0.00", "0%", "mm/dd/yyyy", "@" for text).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelNumberFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B2:B100" -NumberFormat "#,##0.00" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$NumberFormat,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $rng.NumberFormat = $NumberFormat

    $result = @{
        status       = 'formatted'
        range        = $rng.Address($false, $false)
        numberFormat = $NumberFormat
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelColumnWidth {
    <#
    .SYNOPSIS
        Set column width.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Column
        Column letter(s) or range (e.g. "A", "B:D").
    .PARAMETER Width
        Width in character units. Use 0 for auto-fit.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelColumnWidth -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Column "A:D" -Width 0 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Column,
        [Parameter(Mandatory)][double]$Width,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    # Build column range
    $colRange = if ($Column -match '^\d+$') {
        $ws.Columns.Item([int]$Column)
    } else {
        $ws.Columns.Item($Column)
    }

    if ($Width -eq 0) {
        $colRange.AutoFit() | Out-Null
        $action = 'autofit'
    } else {
        $colRange.ColumnWidth = $Width
        $action = 'set'
    }

    $result = @{
        status = $action
        column = $Column
        width  = $colRange.ColumnWidth
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelRowHeight {
    <#
    .SYNOPSIS
        Set row height.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Row
        Row number or range (e.g. 1, "1:5").
    .PARAMETER Height
        Height in points. Use 0 for auto-fit.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelRowHeight -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Row "1" -Height 30 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Row,
        [Parameter(Mandatory)][double]$Height,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $rowRange = $ws.Rows.Item($Row)

    if ($Height -eq 0) {
        $rowRange.AutoFit() | Out-Null
        $action = 'autofit'
    } else {
        $rowRange.RowHeight = $Height
        $action = 'set'
    }

    $result = @{
        status = $action
        row    = $Row
        height = $rowRange.RowHeight
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelAlignment {
    <#
    .SYNOPSIS
        Set text alignment for a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address.
    .PARAMETER Horizontal
        Horizontal alignment: general, left, center, right, fill, justify, distributed, centeracross.
    .PARAMETER Vertical
        Vertical alignment: top, center, bottom, justify, distributed.
    .PARAMETER WrapText
        Enable/disable text wrapping.
    .PARAMETER MergeCells
        Merge or unmerge cells.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelAlignment -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D1" -Horizontal "center" -Vertical "center" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [string]$Horizontal,
        [string]$Vertical,
        [Nullable[bool]]$WrapText,
        [Nullable[bool]]$MergeCells,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $changes = @()

    if (-not [string]::IsNullOrWhiteSpace($Horizontal)) {
        $hKey = $Horizontal.ToLower()
        if ($script:XL_HALIGN.ContainsKey($hKey)) {
            $rng.HorizontalAlignment = $script:XL_HALIGN[$hKey]
            $changes += 'horizontal'
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Vertical)) {
        $vKey = $Vertical.ToLower()
        if ($script:XL_VALIGN.ContainsKey($vKey)) {
            $rng.VerticalAlignment = $script:XL_VALIGN[$vKey]
            $changes += 'vertical'
        }
    }

    if ($null -ne $WrapText) {
        $rng.WrapText = $WrapText
        $changes += 'wrapText'
    }

    if ($null -ne $MergeCells) {
        $rng.MergeCells = $MergeCells
        $changes += 'mergeCells'
    }

    $result = @{
        status  = 'formatted'
        range   = $rng.Address($false, $false)
        changes = $changes
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

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

# ═══════════════════════════════════════════════════════════════════════════
# PRIVATE HELPER: Color conversion
# ═══════════════════════════════════════════════════════════════════════════

function ConvertTo-RGBColor {
    <#
    .SYNOPSIS
        Convert hex color string or integer to Excel RGB long value.
        Excel uses BGR order internally: (Blue * 65536) + (Green * 256) + Red
    #>
    param([Parameter(Mandatory)]$Color)

    if ($Color -is [int] -or $Color -is [long] -or $Color -is [double]) {
        return [long]$Color
    }

    if ($Color -is [string]) {
        $hex = $Color.TrimStart('#')
        if ($hex.Length -eq 6) {
            $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
            $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
            $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
            return [long]($b * 65536 + $g * 256 + $r)  # BGR for Excel
        }
    }

    return [long]$Color
}
