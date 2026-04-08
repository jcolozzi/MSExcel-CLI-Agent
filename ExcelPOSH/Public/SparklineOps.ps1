# Public/SparklineOps.ps1 — Sparkline creation, query, and removal

function Add-ExcelSparkline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$LocationRange,
        [Parameter(Mandatory)][string]$DataRange,
        [ValidateSet('line','column','winloss')]
        [string]$SparklineType = 'line',
        [switch]$ShowHighPoint,
        [switch]$ShowLowPoint,
        [switch]$ShowFirstPoint,
        [switch]$ShowLastPoint,
        [switch]$ShowNegativePoints,
        [switch]$ShowMarkers,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $typeMap = @{ line = 1; column = 2; winloss = 3 }

    $locRange = $ws.Range($LocationRange)
    $group = $locRange.SparklineGroups.Add($typeMap[$SparklineType], $DataRange)

    if ($ShowHighPoint)      { $group.Points.HighPoint.Visible = $true }
    if ($ShowLowPoint)       { $group.Points.LowPoint.Visible = $true }
    if ($ShowFirstPoint)     { $group.Points.FirstPoint.Visible = $true }
    if ($ShowLastPoint)      { $group.Points.LastPoint.Visible = $true }
    if ($ShowNegativePoints) { $group.Points.NegativePoints.Visible = $true }
    if ($ShowMarkers)        { $group.Points.Markers.Visible = $true }

    $result = @{
        status     = 'added'
        location   = $LocationRange
        data_range = $DataRange
        type       = $SparklineType
        count      = $locRange.Cells.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelSparkline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$CellAddress,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if ($CellAddress) {
        $cell = $ws.Range($CellAddress)
        if ($cell.SparklineGroups.Count -gt 0) {
            $sg = $cell.SparklineGroups.Item(1)
            $info = @{
                type            = $sg.Type
                location        = $sg.Location.Address($false, $false)
                source_data     = $sg.SourceData
                count           = $sg.Count
                high_point      = $sg.Points.HighPoint.Visible
                low_point       = $sg.Points.LowPoint.Visible
                first_point     = $sg.Points.FirstPoint.Visible
                last_point      = $sg.Points.LastPoint.Visible
                negative_points = $sg.Points.NegativePoints.Visible
                markers         = $sg.Points.Markers.Visible
            }
            $result = @{
                status          = 'found'
                sparkline_groups = @($info)
            }
        } else {
            $result = @{
                status  = 'none'
                message = 'No sparkline in cell'
            }
        }
    } else {
        $ur = $ws.UsedRange
        $seen = @{}
        $allGroups = @()
        foreach ($cell in $ur.Cells) {
            if ($cell.SparklineGroups.Count -gt 0) {
                $sg  = $cell.SparklineGroups.Item(1)
                $loc = $sg.Location.Address($false, $false)
                if (-not $seen.ContainsKey($loc)) {
                    $seen[$loc] = $true
                    $allGroups += @{
                        type        = $sg.Type
                        location    = $loc
                        source_data = $sg.SourceData
                        count       = $sg.Count
                    }
                }
            }
        }
        $result = @{
            status           = if ($allGroups.Count -gt 0) { 'found' } else { 'none' }
            sparkline_groups = $allGroups
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelSparkline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$CellAddress,
        [switch]$AsJson
    )
    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $range = $ws.Range($CellAddress)
    $range.SparklineGroups.Clear()

    $result = @{
        status = 'removed'
        range  = $CellAddress
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
