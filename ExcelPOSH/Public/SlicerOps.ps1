# Public/SlicerOps.ps1 — Slicers and Timelines (Workbook.SlicerCaches) automation
# Slicer/timeline COM APIs are fragile and version-dependent; all COM is wrapped in try/catch
# and returns status='unsupported' on failure rather than throwing.

function New-ExcelSlicer {
    <#
    .SYNOPSIS
        Create a slicer for a PivotTable or table (ListObject).
    .DESCRIPTION
        Adds a SlicerCache to the workbook via Workbook.SlicerCaches.Add2 bound to the given
        source field, then adds a visible slicer from that cache onto the worksheet. The source
        is resolved as a PivotTable first and falls back to a ListObject (table). Slicer COM APIs
        are version-dependent; on failure this returns status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet that hosts the source object and the new slicer.
    .PARAMETER SourceName
        Name of the source PivotTable or table (ListObject).
    .PARAMETER SourceField
        Name of the field to slice on.
    .PARAMETER Name
        Optional internal name for the slicer.
    .PARAMETER Caption
        Optional visible caption for the slicer.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelSlicer -WorkbookPath C:\data.xlsx -SheetName PivotSheet -SourceName "SalesPivot" -SourceField "Region" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string]$SourceField,
        [string]$Name,
        [string]$Caption,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    try {
        $source = try {
            $ws.PivotTables($SourceName)
        } catch {
            $ws.ListObjects($SourceName)
        }
        $cache  = $wb.SlicerCaches.Add2($source, $SourceField)
        $miss   = [System.Reflection.Missing]::Value
        $slName = if ([string]::IsNullOrWhiteSpace($Name)) { $miss } else { $Name }
        $slCap  = if ([string]::IsNullOrWhiteSpace($Caption)) { $miss } else { $Caption }
        $slicer = $cache.Slicers.Add($ws, $miss, $slName, $slCap, 10, 10, 150, 200)
        $result = @{
            status      = 'ok'
            cache       = $cache.Name
            slicer      = $slicer.Name
            sourceField = $SourceField
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelSlicer {
    <#
    .SYNOPSIS
        List all slicer caches in a workbook and their slicers.
    .DESCRIPTION
        Enumerates Workbook.SlicerCaches. For each cache it reports the cache name, the source
        field, and every visible slicer (name and caption). Slicer COM APIs are version-dependent;
        on failure this returns status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelSlicer -WorkbookPath C:\data.xlsx -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $out = @()
        foreach ($c in $wb.SlicerCaches) {
            $slicers = @()
            foreach ($s in $c.Slicers) {
                $slicers += @{
                    name    = $s.Name
                    caption = (ConvertTo-ExcelSafeValue $s.Caption)
                }
            }
            $out += @{
                name        = $c.Name
                sourceField = (ConvertTo-ExcelSafeValue $c.SourceName)
                slicers     = $slicers
            }
        }
        $result = @{
            status       = 'ok'
            count        = $out.Count
            slicerCaches = $out
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelSlicer {
    <#
    .SYNOPSIS
        Update properties of a slicer (caption, columns, size).
    .DESCRIPTION
        Resolves the slicer cache by name via Workbook.SlicerCaches.Item and operates on its
        first slicer. Only properties whose parameters were supplied are changed. Slicer COM
        APIs are version-dependent; on failure this returns status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the slicer cache to modify.
    .PARAMETER Caption
        New visible caption for the slicer. Omit to leave unchanged.
    .PARAMETER NumberOfColumns
        Number of columns of buttons in the slicer. Omit to leave unchanged.
    .PARAMETER Width
        New slicer width in points. Omit to leave unchanged.
    .PARAMETER Height
        New slicer height in points. Omit to leave unchanged.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelSlicer -WorkbookPath C:\data.xlsx -Name "Slicer_Region" -Caption "Region" -NumberOfColumns 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [string]$Caption,
        [int]$NumberOfColumns,
        [double]$Width,
        [double]$Height,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $cache   = $wb.SlicerCaches.Item($Name)
        $slicer  = $cache.Slicers.Item(1)
        $changed = @()
        if ($PSBoundParameters.ContainsKey('Caption'))         { $slicer.Caption = $Caption; $changed += 'caption' }
        if ($PSBoundParameters.ContainsKey('NumberOfColumns')) { $slicer.NumberOfColumns = $NumberOfColumns; $changed += 'numberOfColumns' }
        if ($PSBoundParameters.ContainsKey('Width'))           { $slicer.Width = $Width; $changed += 'width' }
        if ($PSBoundParameters.ContainsKey('Height'))          { $slicer.Height = $Height; $changed += 'height' }
        $result = @{
            status  = 'ok'
            name    = $Name
            updated = $changed
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelSlicer {
    <#
    .SYNOPSIS
        Delete a slicer cache (and its slicers) from a workbook.
    .DESCRIPTION
        Calls Workbook.SlicerCaches.Item(Name).Delete(), which removes the cache and every
        visible slicer bound to it. Slicer COM APIs are version-dependent; on failure this
        returns status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the slicer cache to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelSlicer -WorkbookPath C:\data.xlsx -Name "Slicer_Region" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $wb.SlicerCaches.Item($Name).Delete()
        $result = @{
            status  = 'ok'
            name    = $Name
            deleted = $true
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelTimeline {
    <#
    .SYNOPSIS
        Create a timeline slicer for a PivotTable date field.
    .DESCRIPTION
        Adds a timeline-type SlicerCache via Workbook.SlicerCaches.Add2 with the timeline type
        constant. The source must be a PivotTable and the field must be a date field. The Add2
        signature for timelines is version-dependent; on failure this returns status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet containing the source PivotTable.
    .PARAMETER SourceName
        Name of the source PivotTable.
    .PARAMETER DateField
        Name of the date field to build the timeline on.
    .PARAMETER Name
        Optional internal name for the timeline slicer cache.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelTimeline -WorkbookPath C:\data.xlsx -SheetName PivotSheet -SourceName "SalesPivot" -DateField "OrderDate" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string]$DateField,
        [string]$Name,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    try {
        $pivot  = $ws.PivotTables($SourceName)
        $cache  = $wb.SlicerCaches.Add2($pivot, $DateField, $Name, $false, [int]$script:XL_SLICER_TYPE.timeline)
        $result = @{
            status    = 'ok'
            cache     = $cache.Name
            dateField = $DateField
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelTimelineRange {
    <#
    .SYNOPSIS
        Set the visible/filtered date range of a timeline slicer.
    .DESCRIPTION
        Resolves the timeline slicer cache by name and calls its TimelineState.SetFilterDateRange
        with the start and end dates. TimelineState is version-dependent; on failure this returns
        status='unsupported'.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the timeline slicer cache to filter.
    .PARAMETER StartDate
        Start of the date range.
    .PARAMETER EndDate
        End of the date range.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelTimelineRange -WorkbookPath C:\data.xlsx -Name "NativeTimeline_OrderDate" -StartDate 2024-01-01 -EndDate 2024-12-31 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $cache = $wb.SlicerCaches.Item($Name)
        $cache.TimelineState.SetFilterDateRange($StartDate, $EndDate)
        $result = @{
            status = 'ok'
            name   = $Name
            start  = $StartDate.ToString('o')
            end    = $EndDate.ToString('o')
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Slicers/timelines unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
