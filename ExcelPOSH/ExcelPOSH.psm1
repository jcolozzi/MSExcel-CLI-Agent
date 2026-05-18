<#
.SYNOPSIS
    Excel-POSH — PowerShell Excel Workbook Automation

.DESCRIPTION
    Provides COM automation of Microsoft Excel workbooks (.xlsx/.xlsm/.xlsb/.xls).
    84 public functions for workbook, worksheet, table, formatting, metadata,
    filter/sort, conditional format, data validation, view, hyperlink, clipboard,
    print, image/shape, pivot table, chart, import, and sparkline operations.
    No MCP server needed — AI agents call functions directly via terminal.

    Usage:
        Import-Module .\ExcelPOSH\ExcelPOSH.psd1 -Force
        Get-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D10" -AsJson
        Close-ExcelWorkbook                           # release COM

.NOTES
    Requires: Windows + Microsoft Excel (Microsoft 365)
    PowerShell: 5.1+ or PowerShell 7+
#>

# ═══════════════════════════════════════════════════════════════════════════
# CONSTANTS & TYPE MAPS
# ═══════════════════════════════════════════════════════════════════════════

# xlFileFormat — file format constants for SaveAs
$script:XL_FILE_FORMAT = @{
    xlsx   = 51     # xlOpenXMLWorkbook
    xlsm   = 52     # xlOpenXMLWorkbookMacroEnabled
    xlsb   = 50     # xlExcel12
    xls    = 56     # xlExcel8
    csv    = 6      # xlCSV
    txt    = -4158  # xlCurrentPlatformText
}

# ExportAsFixedFormat type
$script:XL_FIXED_FORMAT = @{
    pdf  = 0   # xlTypePDF
    xps  = 1   # xlTypeXPS
}

# Border weight
$script:XL_BORDER_WEIGHT = @{
    hairline = 1      # xlHairline
    thin     = 2      # xlThin
    medium   = -4138  # xlMedium
    thick    = 4      # xlThick
}

# Border line style
$script:XL_LINE_STYLE = @{
    continuous   = 1      # xlContinuous
    dash         = -4115  # xlDash
    dashdot      = 4      # xlDashDot
    dashdotdot   = 5      # xlDashDotDot
    dot          = -4118  # xlDot
    double       = -4119  # xlDouble
    none         = -4142  # xlLineStyleNone
    slantdashdot = 13     # xlSlantDashDot
}

# Border index
$script:XL_BORDER_INDEX = @{
    left     = 7   # xlEdgeLeft
    right    = 10  # xlEdgeRight
    top      = 8   # xlEdgeTop
    bottom   = 9   # xlEdgeBottom
    insideh  = 12  # xlInsideHorizontal
    insidev  = 11  # xlInsideVertical
    diagdown = 5   # xlDiagonalDown
    diagup   = 6   # xlDiagonalUp
}

# Horizontal alignment
$script:XL_HALIGN = @{
    general      = 1      # xlGeneral
    left         = -4131  # xlLeft
    center       = -4108  # xlCenter
    right        = -4152  # xlRight
    fill         = 5      # xlFill
    justify      = -4130  # xlJustify
    distributed  = -4117  # xlDistributed
    centeracross = 7      # xlCenterAcrossSelection
}

# Vertical alignment
$script:XL_VALIGN = @{
    top         = -4160  # xlTop
    center      = -4108  # xlCenter
    bottom      = -4107  # xlBottom
    justify     = -4130  # xlJustify
    distributed = -4117  # xlDistributed
}

# Interior pattern
$script:XL_PATTERN = @{
    none   = -4142  # xlNone
    solid  = 1      # xlSolid
    gray50 = -4125  # xlGray50
    gray75 = -4126  # xlGray75
    gray25 = -4124  # xlGray25
}

# ListObject totals calculation
$script:XL_TOTALS_CALC = @{
    none      = 0    # xlTotalsCalculationNone
    average   = 2    # xlTotalsCalculationAverage
    count     = 3    # xlTotalsCalculationCount
    countnums = 4    # xlTotalsCalculationCountNums
    max       = 6    # xlTotalsCalculationMax
    min       = 5    # xlTotalsCalculationMin
    sum       = 109  # xlTotalsCalculationSum
    stddev    = 7    # xlTotalsCalculationStdDev
    var       = 8    # xlTotalsCalculationVar
}

# Sort order
$script:XL_SORT_ORDER = @{
    ascending  = 1  # xlAscending
    descending = 2  # xlDescending
}

# Delete shift direction
$script:XL_DELETE_SHIFT = @{
    up   = -4162  # xlShiftUp
    left = -4159  # xlShiftToLeft
}

# Insert shift direction
$script:XL_INSERT_SHIFT = @{
    down  = -4121  # xlShiftDown
    right = -4161  # xlShiftToRight
}

# Clear type for Clear-ExcelRange
$script:XL_CLEAR_TYPE = @{
    all      = 'Clear'
    contents = 'ClearContents'
    formats  = 'ClearFormats'
    comments = 'ClearComments'
}

# xlSourceType — ListObject source
$script:XL_SOURCE_TYPE = @{
    range    = 1  # xlSrcRange
    external = 2  # xlSrcExternal
}

# xlYesNoGuess — header detection
$script:XL_YES_NO_GUESS = @{
    guess = 0  # xlGuess
    yes   = 1  # xlYes
    no    = 2  # xlNo
}

# xlCalculation — calculation mode
$script:XL_CALCULATION = @{
    automatic     = -4105  # xlCalculationAutomatic
    manual        = -4135  # xlCalculationManual
    semiautomatic = 2      # xlCalculationSemiautomatic
}

# xlCellType — SpecialCells type
$script:XL_CELL_TYPE = @{
    blanks    = 4      # xlCellTypeBlanks
    constants = 2      # xlCellTypeConstants
    formulas  = -4123  # xlCellTypeFormulas
    lastcell  = 11     # xlCellTypeLastCell
    visible   = 12     # xlCellTypeVisible
    comments  = -4144  # xlCellTypeComments
}

# xlFilterAction — AdvancedFilter action
$script:XL_FILTER_ACTION = @{
    filterinplace = 1  # xlFilterInPlace
    copytorange   = 2  # xlFilterCopy
}

# xlAxisType — chart axis type
$script:XL_AXIS_TYPE = @{
    category   = 1  # xlCategory
    value      = 2  # xlValue
    seriesaxis = 3  # xlSeriesAxis
}

# xlAxisGroup — chart axis group
$script:XL_AXIS_GROUP = @{
    primary   = 1  # xlPrimary
    secondary = 2  # xlSecondary
}

# xlLegendPosition — chart legend position
$script:XL_LEGEND_POSITION = @{
    bottom = -4107  # xlLegendPositionBottom
    corner = 2      # xlLegendPositionCorner
    left   = -4131  # xlLegendPositionLeft
    right  = -4152  # xlLegendPositionRight
    top    = -4160  # xlLegendPositionTop
}

# ═══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════

$script:ExcelSession = @{
    App          = $null   # COM Excel.Application object
    WorkbookPath = $null   # Currently open workbook path (resolved)
    OwnsApp      = $false  # $true when we created the COM instance; controls Quit() on exit
}

# ═══════════════════════════════════════════════════════════════════════════
# DOT-SOURCE ALL SUB-FILES
# ═══════════════════════════════════════════════════════════════════════════

# Private helpers first (session, utilities)
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# Public domain files (workbook, worksheet, table, formatting, metadata)
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# ═══════════════════════════════════════════════════════════════════════════
# CLEANUP ON EXIT
# ═══════════════════════════════════════════════════════════════════════════

Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($null -ne $script:ExcelSession -and $null -ne $script:ExcelSession.App) {
        # Only restore settings and Quit if we created the instance
        if ($script:ExcelSession.OwnsApp) {
            try { $script:ExcelSession.App.ScreenUpdating = $true } catch {}
            try { $script:ExcelSession.App.Calculation = -4105 } catch {}   # xlCalculationAutomatic
            try { $script:ExcelSession.App.EnableEvents = $true } catch {}
            try { $script:ExcelSession.App.DisplayAlerts = $false } catch {}
            try { $script:ExcelSession.App.Quit() } catch {}
        }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelSession.App) } catch {}
        $script:ExcelSession.App          = $null
        $script:ExcelSession.WorkbookPath = $null
        $script:ExcelSession.OwnsApp      = $false
    }
} | Out-Null

# ═══════════════════════════════════════════════════════════════════════════
# LOADED
# ═══════════════════════════════════════════════════════════════════════════
Write-Host 'ExcelPOSH module loaded. Use Close-ExcelWorkbook to release COM when done.' -ForegroundColor Cyan
