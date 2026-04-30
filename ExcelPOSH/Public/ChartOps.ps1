# Public/ChartOps.ps1 — Chart creation, modification, and export

function New-ExcelChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceRange,
        [ValidateSet('ColumnClustered','ColumnStacked','BarClustered','BarStacked','Line','LineMarkers','Pie','Doughnut','Area','AreaStacked','XYScatter','XYScatterLines','Radar','Surface3D','Bubble')]
        [string]$ChartType = 'ColumnClustered',
        [string]$ChartName,
        [string]$Title,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = 375,
        [double]$Height = 225,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $typeMap = @{
        ColumnClustered = 51
        ColumnStacked   = 52
        BarClustered    = 57
        BarStacked      = 58
        Line            = 4
        LineMarkers     = 65
        Pie             = 5
        Doughnut        = -4120
        Area            = 1
        AreaStacked     = 76
        XYScatter       = -4169
        XYScatterLines  = 74
        Radar           = -4151
        Surface3D       = 84
        Bubble          = 15
    }

    $chartObj = $ws.ChartObjects().Add($Left, $Top, $Width, $Height)
    $chart    = $chartObj.Chart
    $chart.SetSourceData($ws.Range($SourceRange))
    $chart.ChartType = $typeMap[$ChartType]

    if ($ChartName) { $chartObj.Name = $ChartName }
    if ($Title) {
        $chart.HasTitle       = $true
        $chart.ChartTitle.Text = $Title
    }

    $result = @{
        status       = 'created'
        name         = $chartObj.Name
        chart_type   = $ChartType
        source_range = $SourceRange
        left         = $chartObj.Left
        top          = $chartObj.Top
        width        = $chartObj.Width
        height       = $chartObj.Height
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$ChartName,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $charts = @()

    if ($ChartName) {
        $co    = $ws.ChartObjects($ChartName)
        $chart = $co.Chart
        $info  = @{
            name       = $co.Name
            left       = $co.Left
            top        = $co.Top
            width      = $co.Width
            height     = $co.Height
            chart_type = [int]$chart.ChartType
            has_title  = [bool]$chart.HasTitle
        }
        if ($chart.HasTitle) { $info.title_text = $chart.ChartTitle.Text }
        $charts += $info
    } else {
        foreach ($co in $ws.ChartObjects()) {
            $chart = $co.Chart
            $info  = @{
                name       = $co.Name
                left       = $co.Left
                top        = $co.Top
                width      = $co.Width
                height     = $co.Height
                chart_type = [int]$chart.ChartType
                has_title  = [bool]$chart.HasTitle
            }
            if ($chart.HasTitle) { $info.title_text = $chart.ChartTitle.Text }
            $charts += $info
        }
    }

    $result = @{
        status = 'ok'
        charts = $charts
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [ValidateSet('ColumnClustered','ColumnStacked','BarClustered','BarStacked','Line','LineMarkers','Pie','Doughnut','Area','AreaStacked','XYScatter','XYScatterLines','Radar','Surface3D','Bubble')]
        [string]$ChartType,
        [string]$Title,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$SourceRange,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $typeMap = @{
        ColumnClustered = 51
        ColumnStacked   = 52
        BarClustered    = 57
        BarStacked      = 58
        Line            = 4
        LineMarkers     = 65
        Pie             = 5
        Doughnut        = -4120
        Area            = 1
        AreaStacked     = 76
        XYScatter       = -4169
        XYScatterLines  = 74
        Radar           = -4151
        Surface3D       = 84
        Bubble          = 15
    }

    $co    = $ws.ChartObjects($ChartName)
    $chart = $co.Chart
    $changes = @()

    if ($PSBoundParameters.ContainsKey('ChartType')) {
        $chart.ChartType = $typeMap[$ChartType]
        $changes += 'chart_type'
    }
    if ($PSBoundParameters.ContainsKey('Title')) {
        $chart.HasTitle        = $true
        $chart.ChartTitle.Text = $Title
        $changes += 'title'
    }
    if ($PSBoundParameters.ContainsKey('Left')) {
        $co.Left = $Left
        $changes += 'left'
    }
    if ($PSBoundParameters.ContainsKey('Top')) {
        $co.Top = $Top
        $changes += 'top'
    }
    if ($PSBoundParameters.ContainsKey('Width')) {
        $co.Width = $Width
        $changes += 'width'
    }
    if ($PSBoundParameters.ContainsKey('Height')) {
        $co.Height = $Height
        $changes += 'height'
    }
    if ($PSBoundParameters.ContainsKey('SourceRange')) {
        $chart.SetSourceData($ws.Range($SourceRange))
        $changes += 'source_range'
    }

    $result = @{
        status  = 'modified'
        name    = $co.Name
        changes = $changes
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Export-ExcelChart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][string]$OutputPath,
        [int]$Width,
        [int]$Height,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $co    = $ws.ChartObjects($ChartName)
    $chart = $co.Chart

    if ($PSBoundParameters.ContainsKey('Width'))  { $co.Width  = $Width }
    if ($PSBoundParameters.ContainsKey('Height')) { $co.Height = $Height }

    $chart.Export($OutputPath)

    $result = @{
        status = 'exported'
        name   = $co.Name
        path   = $OutputPath
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelChartSeries {
    <#
    .SYNOPSIS  Modify a chart data series properties.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Worksheet containing the chart.
    .PARAMETER ChartName     Name of the chart object.
    .PARAMETER SeriesIndex   1-based index of the series to modify.
    .PARAMETER Name          Series display name.
    .PARAMETER Values        Range address for series values.
    .PARAMETER XValues       Range address for category (X) values.
    .PARAMETER FillColor     Hex fill color (e.g. "#FF0000").
    .PARAMETER LineColor     Hex line color.
    .PARAMETER LineWeight    Line weight in points.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Set-ExcelChartSeries -WorkbookPath C:\data.xlsx -SheetName Sheet1 -ChartName "Chart 1" -SeriesIndex 1 -Name "Revenue" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][int]$SeriesIndex,
        [string]$Name,
        [string]$Values,
        [string]$XValues,
        [string]$FillColor,
        [string]$LineColor,
        [double]$LineWeight,
        [switch]$AsJson
    )

    $app    = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart  = $ws.ChartObjects($ChartName).Chart
    $series = $chart.SeriesCollection($SeriesIndex)

    if ($PSBoundParameters.ContainsKey('Name'))      { $series.Name    = $Name }
    if ($PSBoundParameters.ContainsKey('Values'))     { $series.Values  = $ws.Range($Values) }
    if ($PSBoundParameters.ContainsKey('XValues'))    { $series.XValues = $ws.Range($XValues) }
    if ($PSBoundParameters.ContainsKey('FillColor'))  { $series.Format.Fill.ForeColor.RGB = ConvertTo-RGBColor $FillColor }
    if ($PSBoundParameters.ContainsKey('LineColor'))  { $series.Format.Line.ForeColor.RGB = ConvertTo-RGBColor $LineColor }
    if ($PSBoundParameters.ContainsKey('LineWeight')) { $series.Format.Line.Weight = $LineWeight }

    $result = @{ status = 'ok'; chart = $ChartName; seriesIndex = $SeriesIndex }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelChartAxis {
    <#
    .SYNOPSIS  Configure a chart axis (title, scale, number format).
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Worksheet containing the chart.
    .PARAMETER ChartName     Name of the chart object.
    .PARAMETER AxisType      Category, Value, or SeriesAxis.
    .PARAMETER AxisGroup     Primary or Secondary axis group.
    .PARAMETER Title         Axis title text.
    .PARAMETER MinimumScale  Minimum axis value.
    .PARAMETER MaximumScale  Maximum axis value.
    .PARAMETER NumberFormat  Number format string for tick labels.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Set-ExcelChartAxis -WorkbookPath C:\data.xlsx -SheetName Sheet1 -ChartName "Chart 1" -AxisType Value -Title "Revenue ($)" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)]
        [ValidateSet('Category','Value','SeriesAxis')]
        [string]$AxisType,
        [ValidateSet('Primary','Secondary')]
        [string]$AxisGroup = 'Primary',
        [string]$Title,
        [double]$MinimumScale,
        [double]$MaximumScale,
        [string]$NumberFormat,
        [switch]$AsJson
    )

    $app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart = $ws.ChartObjects($ChartName).Chart

    $axType  = [int]$script:XL_AXIS_TYPE[$AxisType.ToLower()]
    $axGroup = [int]$script:XL_AXIS_GROUP[$AxisGroup.ToLower()]
    $axis    = $chart.Axes($axType, $axGroup)

    if ($PSBoundParameters.ContainsKey('Title')) {
        $axis.HasTitle       = $true
        $axis.AxisTitle.Text = $Title
    }
    if ($PSBoundParameters.ContainsKey('MinimumScale')) { $axis.MinimumScale = $MinimumScale }
    if ($PSBoundParameters.ContainsKey('MaximumScale')) { $axis.MaximumScale = $MaximumScale }
    if ($PSBoundParameters.ContainsKey('NumberFormat')) { $axis.TickLabels.NumberFormat = $NumberFormat }

    $result = @{ status = 'ok'; chart = $ChartName; axisType = $AxisType; axisGroup = $AxisGroup }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelChartLegend {
    <#
    .SYNOPSIS  Show/hide and position the chart legend.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Worksheet containing the chart.
    .PARAMETER ChartName     Name of the chart object.
    .PARAMETER Show          Show or hide the legend.
    .PARAMETER Position      Legend position: Bottom, Corner, Left, Right, Top.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Set-ExcelChartLegend -WorkbookPath C:\data.xlsx -SheetName Sheet1 -ChartName "Chart 1" -Show $true -Position Bottom -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][bool]$Show,
        [ValidateSet('Bottom','Corner','Left','Right','Top')]
        [string]$Position,
        [switch]$AsJson
    )

    $app   = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws    = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart = $ws.ChartObjects($ChartName).Chart

    $chart.HasLegend = $Show
    if ($Show -and $PSBoundParameters.ContainsKey('Position')) {
        $chart.Legend.Position = [int]$script:XL_LEGEND_POSITION[$Position.ToLower()]
    }

    $result = @{ status = 'ok'; chart = $ChartName; hasLegend = $Show }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelChartDataLabels {
    <#
    .SYNOPSIS  Configure data labels on a chart series.
    .PARAMETER WorkbookPath   Path to the Excel workbook.
    .PARAMETER SheetName      Worksheet containing the chart.
    .PARAMETER ChartName      Name of the chart object.
    .PARAMETER SeriesIndex    1-based index of the series.
    .PARAMETER ShowValue      Show data values.
    .PARAMETER ShowCategory   Show category names.
    .PARAMETER ShowPercentage Show percentages (pie/doughnut charts).
    .PARAMETER NumberFormat   Number format string for labels.
    .PARAMETER Position       Label position: Center, InsideBase, InsideEnd, OutsideEnd, BestFit.
    .PARAMETER AsJson         Return JSON string.
    .EXAMPLE   Set-ExcelChartDataLabels -WorkbookPath C:\data.xlsx -SheetName Sheet1 -ChartName "Chart 1" -SeriesIndex 1 -ShowValue $true -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ChartName,
        [Parameter(Mandatory)][int]$SeriesIndex,
        [bool]$ShowValue,
        [bool]$ShowCategory,
        [bool]$ShowPercentage,
        [string]$NumberFormat,
        [ValidateSet('Center','InsideBase','InsideEnd','OutsideEnd','BestFit')]
        [string]$Position,
        [switch]$AsJson
    )

    $app    = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $chart  = $ws.ChartObjects($ChartName).Chart
    $series = $chart.SeriesCollection($SeriesIndex)

    $series.HasDataLabels = $true
    $labels = $series.DataLabels

    if ($PSBoundParameters.ContainsKey('ShowValue'))      { $labels.ShowValue        = $ShowValue }
    if ($PSBoundParameters.ContainsKey('ShowCategory'))   { $labels.ShowCategoryName = $ShowCategory }
    if ($PSBoundParameters.ContainsKey('ShowPercentage')) { $labels.ShowPercentage   = $ShowPercentage }
    if ($PSBoundParameters.ContainsKey('NumberFormat'))   { $labels.NumberFormat     = $NumberFormat }
    if ($PSBoundParameters.ContainsKey('Position')) {
        $posMap = @{ Center = 0; InsideBase = 3; InsideEnd = 1; OutsideEnd = 2; BestFit = 5 }
        $labels.Position = $posMap[$Position]
    }

    $result = @{ status = 'ok'; chart = $ChartName; seriesIndex = $SeriesIndex }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
