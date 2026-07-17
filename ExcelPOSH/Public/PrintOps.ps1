# Public/PrintOps.ps1 — Page setup, print area, PDF export

function Set-ExcelPageSetup {
    <#
    .SYNOPSIS
        Configure page setup for a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER Orientation
        Page orientation: portrait or landscape.
    .PARAMETER PaperSize
        Paper size: letter, legal, a4, a3.
    .PARAMETER LeftMargin
        Left margin in inches.
    .PARAMETER RightMargin
        Right margin in inches.
    .PARAMETER TopMargin
        Top margin in inches.
    .PARAMETER BottomMargin
        Bottom margin in inches.
    .PARAMETER HeaderLeft
        Left header text.
    .PARAMETER HeaderCenter
        Center header text.
    .PARAMETER HeaderRight
        Right header text.
    .PARAMETER FooterLeft
        Left footer text.
    .PARAMETER FooterCenter
        Center footer text.
    .PARAMETER FooterRight
        Right footer text.
    .PARAMETER FitToPagesWide
        Number of pages wide to fit.
    .PARAMETER FitToPagesTall
        Number of pages tall to fit.
    .PARAMETER PrintArea
        Print area range (e.g. "A1:G50"). Pass empty string to clear.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelPageSetup -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Orientation landscape -PaperSize letter -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [ValidateSet('portrait','landscape')][string]$Orientation,
        [ValidateSet('letter','legal','a4','a3')][string]$PaperSize,
        [double]$LeftMargin,
        [double]$RightMargin,
        [double]$TopMargin,
        [double]$BottomMargin,
        [string]$HeaderLeft,
        [string]$HeaderCenter,
        [string]$HeaderRight,
        [string]$FooterLeft,
        [string]$FooterCenter,
        [string]$FooterRight,
        [int]$FitToPagesWide,
        [int]$FitToPagesTall,
        [string]$PrintArea,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $ps  = $ws.PageSetup

    $changes = @()

    $orientationMap = @{ 'portrait' = 1; 'landscape' = 2 }
    $paperSizeMap   = @{ 'letter' = 1; 'legal' = 5; 'a4' = 9; 'a3' = 8 }

    if ($PSBoundParameters.ContainsKey('Orientation')) {
        $ps.Orientation = $orientationMap[$Orientation]
        $changes += "Orientation=$Orientation"
    }

    if ($PSBoundParameters.ContainsKey('PaperSize')) {
        $ps.PaperSize = $paperSizeMap[$PaperSize]
        $changes += "PaperSize=$PaperSize"
    }

    if ($PSBoundParameters.ContainsKey('LeftMargin')) {
        $ps.LeftMargin = $LeftMargin * 72
        $changes += "LeftMargin=$LeftMargin"
    }

    if ($PSBoundParameters.ContainsKey('RightMargin')) {
        $ps.RightMargin = $RightMargin * 72
        $changes += "RightMargin=$RightMargin"
    }

    if ($PSBoundParameters.ContainsKey('TopMargin')) {
        $ps.TopMargin = $TopMargin * 72
        $changes += "TopMargin=$TopMargin"
    }

    if ($PSBoundParameters.ContainsKey('BottomMargin')) {
        $ps.BottomMargin = $BottomMargin * 72
        $changes += "BottomMargin=$BottomMargin"
    }

    if ($PSBoundParameters.ContainsKey('HeaderLeft')) {
        $ps.LeftHeader = $HeaderLeft
        $changes += "HeaderLeft=$HeaderLeft"
    }

    if ($PSBoundParameters.ContainsKey('HeaderCenter')) {
        $ps.CenterHeader = $HeaderCenter
        $changes += "HeaderCenter=$HeaderCenter"
    }

    if ($PSBoundParameters.ContainsKey('HeaderRight')) {
        $ps.RightHeader = $HeaderRight
        $changes += "HeaderRight=$HeaderRight"
    }

    if ($PSBoundParameters.ContainsKey('FooterLeft')) {
        $ps.LeftFooter = $FooterLeft
        $changes += "FooterLeft=$FooterLeft"
    }

    if ($PSBoundParameters.ContainsKey('FooterCenter')) {
        $ps.CenterFooter = $FooterCenter
        $changes += "FooterCenter=$FooterCenter"
    }

    if ($PSBoundParameters.ContainsKey('FooterRight')) {
        $ps.RightFooter = $FooterRight
        $changes += "FooterRight=$FooterRight"
    }

    if ($PSBoundParameters.ContainsKey('FitToPagesWide') -or $PSBoundParameters.ContainsKey('FitToPagesTall')) {
        $ps.Zoom = $false
        if ($PSBoundParameters.ContainsKey('FitToPagesWide')) {
            $ps.FitToPagesWide = $FitToPagesWide
            $changes += "FitToPagesWide=$FitToPagesWide"
        }
        if ($PSBoundParameters.ContainsKey('FitToPagesTall')) {
            $ps.FitToPagesTall = $FitToPagesTall
            $changes += "FitToPagesTall=$FitToPagesTall"
        }
    }

    if ($PSBoundParameters.ContainsKey('PrintArea')) {
        $ps.PrintArea = $PrintArea
        $changes += "PrintArea=$PrintArea"
    }

    $result = @{
        status  = 'configured'
        sheet   = $SheetName
        changes = $changes
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelPageSetup {
    <#
    .SYNOPSIS
        Get current page setup properties for a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the worksheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelPageSetup -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
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
    $ps  = $ws.PageSetup

    $setup = @{
        orientation    = $ps.Orientation
        paperSize      = $ps.PaperSize
        leftMargin     = $ps.LeftMargin / 72
        rightMargin    = $ps.RightMargin / 72
        topMargin      = $ps.TopMargin / 72
        bottomMargin   = $ps.BottomMargin / 72
        headerLeft     = $ps.LeftHeader
        headerCenter   = $ps.CenterHeader
        headerRight    = $ps.RightHeader
        footerLeft     = $ps.LeftFooter
        footerCenter   = $ps.CenterFooter
        footerRight    = $ps.RightFooter
        printArea      = $ps.PrintArea
        zoom           = $ps.Zoom
        fitToPagesWide = $ps.FitToPagesWide
        fitToPagesTall = $ps.FitToPagesTall
    }

    $result = @{
        status = 'ok'
        sheet  = $SheetName
        setup  = $setup
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Export-ExcelToPdf {
    <#
    .SYNOPSIS
        Export a worksheet or entire workbook to PDF.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER OutputPath
        File path for the output PDF.
    .PARAMETER SheetName
        Name of the worksheet to export. If omitted, exports the entire workbook.
    .PARAMETER Quality
        PDF quality: standard or minimum.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-ExcelToPdf -WorkbookPath "C:\data.xlsx" -OutputPath "C:\report.pdf" -AsJson
    .EXAMPLE
        Export-ExcelToPdf -WorkbookPath "C:\data.xlsx" -OutputPath "C:\sheet1.pdf" -SheetName "Sheet1" -Quality minimum -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$SheetName,
        [ValidateSet('PDF','XPS')][string]$Format = 'PDF',
        [ValidateSet('standard','minimum')][string]$Quality = 'standard',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $qualityMap = @{ 'standard' = 0; 'minimum' = 1 }
    $qualityConst = $qualityMap[$Quality]
    $formatConst  = [int]$script:XL_FIXED_FORMAT[$Format.ToLower()]

    if ($SheetName) {
        $ws = $wb.Worksheets.Item($SheetName)
        $ws.ExportAsFixedFormat($formatConst, $OutputPath, $qualityConst)
    } else {
        $wb.ExportAsFixedFormat($formatConst, $OutputPath, $qualityConst)
    }

    $result = @{
        status = 'exported'
        path   = $OutputPath
        format = $Format.ToLower()
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Send-ExcelPrint {
    <#
    .SYNOPSIS
        Send a worksheet directly to a printer.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet to print.
    .PARAMETER Copies
        Number of copies (default 1).
    .PARAMETER FromPage
        First page to print (optional).
    .PARAMETER ToPage
        Last page to print (optional).
    .PARAMETER Preview
        Show print preview instead of printing.
    .PARAMETER Printer
        Optional active printer name (e.g. "HP LaserJet on Ne01:").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Send-ExcelPrint -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Copies 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [int]$Copies = 1,
        [int]$FromPage,
        [int]$ToPage,
        [switch]$Preview,
        [string]$Printer,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $from       = if ($PSBoundParameters.ContainsKey('FromPage')) { $FromPage } else { [System.Reflection.Missing]::Value }
    $to         = if ($PSBoundParameters.ContainsKey('ToPage'))   { $ToPage }   else { [System.Reflection.Missing]::Value }
    $printerVal = if ([string]::IsNullOrWhiteSpace($Printer))      { [System.Reflection.Missing]::Value } else { $Printer }

    # PrintOut(From, To, Copies, Preview, ActivePrinter)
    $ws.PrintOut($from, $to, $Copies, $Preview.IsPresent, $printerVal)

    $result = @{
        status  = 'ok'
        sheet   = $SheetName
        copies  = $Copies
        preview = $Preview.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
