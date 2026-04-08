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
