# Public/FilterSortOps.ps1 — AutoFilter and sorting operations

function Set-ExcelAutoFilter {
    <#
    .SYNOPSIS
        Enable or apply AutoFilter on a range in a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to apply the filter to.
    .PARAMETER Range
        Range address (e.g. "A1:D10"). Defaults to UsedRange if omitted.
    .PARAMETER Column
        Column number within the range to filter (1-based).
    .PARAMETER Criteria1
        First filter criteria value.
    .PARAMETER Operator
        Filter operator: xlAnd or xlOr. Maps to Excel constants.
    .PARAMETER Criteria2
        Second filter criteria value (used with Operator).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelAutoFilter -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    .EXAMPLE
        Set-ExcelAutoFilter -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:E20" -Column 2 -Criteria1 ">100" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Range,
        [int]$Column,
        [string]$Criteria1,
        [string]$Operator,
        [string]$Criteria2,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if ([string]::IsNullOrWhiteSpace($Range)) {
        $rng = $ws.UsedRange
    } else {
        $rng = $ws.Range($Range)
    }

    $rangeAddr = $rng.Address($false, $false)
    $filtered  = $false

    if ($Column -gt 0 -and -not [string]::IsNullOrWhiteSpace($Criteria1)) {
        # Map operator string to Excel constant
        $operatorMap = @{
            'xlAnd' = 1
            'xlOr'  = 2
        }

        if (-not [string]::IsNullOrWhiteSpace($Operator) -and -not [string]::IsNullOrWhiteSpace($Criteria2)) {
            $opVal = if ($operatorMap.ContainsKey($Operator)) { $operatorMap[$Operator] } else { 1 }
            $rng.AutoFilter($Column, $Criteria1, $opVal, $Criteria2)
        } else {
            $rng.AutoFilter($Column, $Criteria1)
        }
        $filtered = $true
    } else {
        # Toggle AutoFilter on/off
        $rng.AutoFilter()
        $filtered = $ws.AutoFilterMode
    }

    $result = @{
        status   = 'ok'
        range    = $rangeAddr
        filtered = $filtered
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelAutoFilter {
    <#
    .SYNOPSIS
        Clear all filters from a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to clear filters from.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelAutoFilter -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
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

    if ($ws.AutoFilterMode) {
        $ws.AutoFilterMode = $false
    }

    $result = @{
        status = 'removed'
        sheet  = $SheetName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Sort-ExcelRange {
    <#
    .SYNOPSIS
        Sort a range by one or two keys.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet containing the range.
    .PARAMETER Range
        Range address to sort (e.g. "A1:D20").
    .PARAMETER SortKey1
        Column letter or range for the primary sort key (e.g. "A1").
    .PARAMETER Order1
        Sort order for key 1: ascending (default) or descending.
    .PARAMETER SortKey2
        Optional column letter or range for a secondary sort key.
    .PARAMETER Order2
        Sort order for key 2: ascending (default) or descending.
    .PARAMETER Header
        Whether the range has a header row. Default: $true.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Sort-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D20" -SortKey1 "B1" -AsJson
    .EXAMPLE
        Sort-ExcelRange -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:D20" -SortKey1 "A1" -Order1 "ascending" -SortKey2 "C1" -Order2 "descending" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$SortKey1,
        [string]$Order1 = 'ascending',
        [string]$SortKey2,
        [string]$Order2 = 'ascending',
        [bool]$Header = $true,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $rng = $ws.Range($Range)

    # xlAscending = 1, xlDescending = 2
    $order1Val = if ($Order1 -eq 'descending') { 2 } else { 1 }

    # xlYes = 1, xlNo = 2
    $headerVal = if ($Header) { 1 } else { 2 }

    $keysUsed = @($SortKey1)

    if (-not [string]::IsNullOrWhiteSpace($SortKey2)) {
        $order2Val = if ($Order2 -eq 'descending') { 2 } else { 1 }
        $rng.Sort(
            $ws.Range($SortKey1), $order1Val,
            $ws.Range($SortKey2), [System.Reflection.Missing]::Value, $order2Val,
            [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
            $headerVal
        )
        $keysUsed += $SortKey2
    } else {
        $rng.Sort(
            $ws.Range($SortKey1), $order1Val,
            [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
            $headerVal
        )
    }

    $result = @{
        status = 'sorted'
        range  = $rng.Address($false, $false)
        keys   = $keysUsed
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelAutoFilter {
    <#
    .SYNOPSIS
        Get current AutoFilter state for a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet to inspect.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelAutoFilter -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
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

    $active      = $ws.AutoFilterMode
    $filterRange = $null
    $filters     = @()

    if ($active -and $null -ne $ws.AutoFilter) {
        $filterRange = $ws.AutoFilter.Range.Address($false, $false)
        $filterObj   = $ws.AutoFilter.Filters
        for ($i = 1; $i -le $filterObj.Count; $i++) {
            $f = $filterObj.Item($i)
            if ($f.On) {
                $entry = @{
                    column    = $i
                    criteria1 = $null
                    operator  = $null
                    criteria2 = $null
                }
                try { $entry.criteria1 = $f.Criteria1 } catch {}
                try { $entry.operator  = $f.Operator  } catch {}
                try { $entry.criteria2 = $f.Criteria2 } catch {}
                $filters += $entry
            }
        }
    }

    $result = @{
        status      = 'ok'
        active      = $active
        filterRange = $filterRange
        filters     = $filters
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelDuplicates {
    <#
    .SYNOPSIS  Remove duplicate rows from a range.
    .PARAMETER WorkbookPath  Path to the Excel workbook.
    .PARAMETER SheetName     Target worksheet name.
    .PARAMETER Range         Range address (e.g. "A1:D100").
    .PARAMETER Columns       1-based column indices within the range to compare for duplicates.
    .PARAMETER HasHeader     Indicates the first row is a header (default $true).
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Remove-ExcelDuplicates -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D100" -Columns 1,3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][int[]]$Columns,
        [switch]$HasHeader,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $headerVal = if ($HasHeader) { [int]$script:XL_YES_NO_GUESS.yes } else { [int]$script:XL_YES_NO_GUESS.no }
    $rng.RemoveDuplicates($Columns, $headerVal)

    $result = @{
        status    = 'ok'
        range     = $Range
        hasHeader = $HasHeader.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Invoke-ExcelAdvancedFilter {
    <#
    .SYNOPSIS  Apply Excel Advanced Filter to a range.
    .PARAMETER WorkbookPath   Path to the Excel workbook.
    .PARAMETER SheetName      Target worksheet name.
    .PARAMETER Range          Source data range.
    .PARAMETER Action         FilterInPlace or CopyToRange.
    .PARAMETER CriteriaRange  Range containing filter criteria.
    .PARAMETER CopyToRange    Destination range (required for CopyToRange action).
    .PARAMETER UniqueOnly     Return only unique records.
    .PARAMETER AsJson         Return JSON string.
    .EXAMPLE   Invoke-ExcelAdvancedFilter -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D100" -Action FilterInPlace -UniqueOnly -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)]
        [ValidateSet('FilterInPlace','CopyToRange')]
        [string]$Action,
        [string]$CriteriaRange,
        [string]$CopyToRange,
        [switch]$UniqueOnly,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $actionVal = [int]$script:XL_FILTER_ACTION[$Action.ToLower()]

    $criteria = if (-not [string]::IsNullOrWhiteSpace($CriteriaRange)) { $ws.Range($CriteriaRange) } else { [System.Reflection.Missing]::Value }
    $copyTo   = if (-not [string]::IsNullOrWhiteSpace($CopyToRange))   { $ws.Range($CopyToRange)   } else { [System.Reflection.Missing]::Value }

    if ($Action -eq 'CopyToRange' -and [string]::IsNullOrWhiteSpace($CopyToRange)) {
        throw 'CopyToRange parameter is required when Action is CopyToRange.'
    }

    $rng.AdvancedFilter($actionVal, $criteria, $copyTo, $UniqueOnly.IsPresent)

    $result = @{
        status = 'ok'
        action = $Action
        unique = $UniqueOnly.IsPresent
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelSubtotal {
    <#
    .SYNOPSIS
        Insert automatic subtotals with group breaks into a range.
    .DESCRIPTION
        Wraps Range.Subtotal. The range should be sorted by the GroupBy column first.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Target worksheet name.
    .PARAMETER Range
        Data range including headers (e.g. "A1:D100").
    .PARAMETER GroupBy
        1-based column index to group by (break subtotals when this column changes).
    .PARAMETER Function
        Aggregate: average, count, countnums, max, min, product, stdev, stdevp, sum, var, varp.
    .PARAMETER TotalColumns
        1-based column indices to subtotal.
    .PARAMETER Replace
        Replace existing subtotals (default $true).
    .PARAMETER SummaryBelow
        Place summary rows below data (default $true).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelSubtotal -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1:D100" -GroupBy 1 -Function sum -TotalColumns 3,4 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][int]$GroupBy,
        [Parameter(Mandatory)]
        [ValidateSet('average','count','countnums','max','min','product','stdev','stdevp','sum','var','varp')]
        [string]$Function,
        [Parameter(Mandatory)][int[]]$TotalColumns,
        [bool]$Replace = $true,
        [bool]$SummaryBelow = $true,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $fn = [int]$script:XL_CONSOLIDATION_FN[$Function.ToLower()]
    # Subtotal(GroupBy, Function, TotalList, [Replace], [PageBreaks], [SummaryBelowData])
    $rng.Subtotal($GroupBy, $fn, $TotalColumns, $Replace, $false, $SummaryBelow)

    $result = @{
        status       = 'ok'
        range        = $Range
        groupBy      = $GroupBy
        function     = $Function
        totalColumns = $TotalColumns
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
