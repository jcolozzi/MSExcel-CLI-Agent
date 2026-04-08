# Public/ConditionalFormatOps.ps1 — Conditional formatting rules

function Add-ExcelConditionalFormat {
    <#
    .SYNOPSIS
        Add a conditional format rule to a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address (e.g. "A1:A100").
    .PARAMETER RuleType
        Type of conditional format rule.
    .PARAMETER Operator
        Comparison operator (for CellValue type).
    .PARAMETER Value1
        First comparison value (for CellValue/Top10).
    .PARAMETER Value2
        Second comparison value (for 'between'/'notbetween' operators).
    .PARAMETER Formula
        Formula string (for Expression type, e.g. "=$A1>100").
    .PARAMETER FontColor
        Font color as hex string (e.g. "#FF0000").
    .PARAMETER FillColor
        Fill/background color as hex string.
    .PARAMETER Bold
        Set font bold on the format condition.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:A100" -RuleType CellValue -Operator greater -Value1 50 -FillColor "#FFFF00" -AsJson
    .EXAMPLE
        Add-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B2:B50" -RuleType Expression -Formula "=$B2>$A2" -FontColor "#FF0000" -Bold $true -AsJson
    .EXAMPLE
        Add-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "C1:C100" -RuleType ColorScale -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)]
        [ValidateSet('CellValue','Expression','ColorScale','DataBar','IconSet','Top10','AboveAverage','DuplicateValues','UniqueValues')]
        [string]$RuleType,
        [ValidateSet('between','notbetween','equal','notequal','greater','greaterequal','less','lessequal')]
        [string]$Operator,
        [string]$Value1,
        [string]$Value2,
        [string]$Formula,
        [string]$FontColor,
        [string]$FillColor,
        [Nullable[bool]]$Bold,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    # Operator constant map: xlBetween=1, xlNotBetween=2, xlEqual=3, xlNotEqual=4,
    # xlGreater=5, xlLess=6, xlGreaterEqual=7, xlLessEqual=8
    $operatorMap = @{
        'between'      = 1
        'notbetween'   = 2
        'equal'        = 3
        'notequal'     = 4
        'greater'      = 5
        'less'         = 6
        'greaterequal' = 7
        'lessequal'    = 8
    }

    $fc = $null
    switch ($RuleType) {
        'CellValue' {
            if (-not $Operator) {
                throw "Operator is required for CellValue rule type."
            }
            $operatorConst = $operatorMap[$Operator]
            if ($Value2) {
                $fc = $rng.FormatConditions.Add(1, $operatorConst, $Value1, $Value2)   # xlCellValue=1
            } else {
                $fc = $rng.FormatConditions.Add(1, $operatorConst, $Value1)
            }
        }
        'Expression' {
            if (-not $Formula) {
                throw "Formula is required for Expression rule type."
            }
            $fc = $rng.FormatConditions.Add(2, [System.Reflection.Missing]::Value, $Formula)  # xlExpression=2
        }
        'ColorScale' {
            $fc = $rng.FormatConditions.AddColorScale(3)  # 3-color scale
        }
        'DataBar' {
            $fc = $rng.FormatConditions.AddDatabar()
        }
        'IconSet' {
            $fc = $rng.FormatConditions.AddIconSetCondition()
        }
        'Top10' {
            $fc = $rng.FormatConditions.AddTop10()
            $fc.TopBottom = 1   # xlTop10Top=1
            $fc.Rank = if ($Value1) { [int]$Value1 } else { 10 }
        }
        'AboveAverage' {
            $fc = $rng.FormatConditions.AddAboveAverage()
        }
        'DuplicateValues' {
            $fc = $rng.FormatConditions.AddUniqueValues()
            $fc.DupeUnique = 1  # xlDuplicate=1
        }
        'UniqueValues' {
            $fc = $rng.FormatConditions.AddUniqueValues()
            $fc.DupeUnique = 0  # xlUnique=0
        }
    }

    # Apply optional font/fill formatting (only for types that support .Font/.Interior)
    if ($null -ne $fc) {
        if ($FontColor) {
            $fc.Font.Color = ConvertTo-RGBColor -Color $FontColor
        }
        if ($FillColor) {
            $fc.Interior.Color = ConvertTo-RGBColor -Color $FillColor
        }
        if ($null -ne $Bold) {
            $fc.Font.Bold = $Bold
        }
    }

    $result = @{
        status    = 'added'
        range     = $Range
        ruleType  = $RuleType
        ruleCount = $rng.FormatConditions.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelConditionalFormat {
    <#
    .SYNOPSIS
        List conditional format rules on a range or entire sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Optional: cell or range address. If omitted, checks the UsedRange.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    .EXAMPLE
        Get-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:A100" -AsJson
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

    $rng = if ([string]::IsNullOrWhiteSpace($Range)) {
        $ws.UsedRange
    } else {
        $ws.Range($Range)
    }

    $rules = @()
    for ($i = 1; $i -le $rng.FormatConditions.Count; $i++) {
        $fc = $rng.FormatConditions.Item($i)

        $formula1 = $null
        $formula2 = $null
        $appliesTo = $null

        try { $formula1 = $fc.Formula1 } catch {}
        try { $formula2 = $fc.Formula2 } catch {}
        try { $appliesTo = $fc.AppliesTo.Address($false, $false) } catch {}

        $rules += @{
            index     = $i
            type      = $fc.Type
            formula1  = $formula1
            formula2  = $formula2
            appliesTo = $appliesTo
        }
    }

    $result = @{
        status = 'ok'
        count  = $rules.Count
        rules  = $rules
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelConditionalFormat {
    <#
    .SYNOPSIS
        Remove conditional format rules from a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address.
    .PARAMETER RuleIndex
        1-based index of the specific rule to remove.
    .PARAMETER All
        Remove all conditional format rules on the range.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:A100" -RuleIndex 2 -AsJson
    .EXAMPLE
        Remove-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1:A100" -All -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [int]$RuleIndex,
        [switch]$All,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $countBefore = $rng.FormatConditions.Count

    if ($All) {
        $rng.FormatConditions.Delete()
        $removedCount = $countBefore
    } elseif ($RuleIndex -gt 0) {
        $rng.FormatConditions.Item($RuleIndex).Delete()
        $removedCount = 1
    } else {
        throw "Specify -RuleIndex or -All to remove conditional format rules."
    }

    $result = @{
        status       = 'removed'
        range        = $Range
        removedCount = $removedCount
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Clear-ExcelConditionalFormat {
    <#
    .SYNOPSIS
        Remove ALL conditional formatting from an entire sheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Clear-ExcelConditionalFormat -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
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

    $ws.Cells.FormatConditions.Delete()

    $result = @{
        status = 'cleared'
        sheet  = $SheetName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
