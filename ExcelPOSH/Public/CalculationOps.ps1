# Public/CalculationOps.ps1 — Performance controls, calculation, formula evaluation

function Set-ExcelPerformanceMode {
    <#
    .SYNOPSIS
        Control Excel performance settings (ScreenUpdating, Calculation, EnableEvents).
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER ScreenUpdating
        Enable/disable screen redraw.
    .PARAMETER Calculation
        Calculation mode: Automatic, Manual, SemiAutomatic.
    .PARAMETER EnableEvents
        Enable/disable event processing.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelPerformanceMode -WorkbookPath C:\data.xlsx -ScreenUpdating $false -Calculation Manual -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [bool]$ScreenUpdating,
        [ValidateSet('Automatic','Manual','SemiAutomatic')]
        [string]$Calculation,
        [bool]$EnableEvents,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    # Capture previous state for caller restoration
    $calcReverse = @{ -4105 = 'Automatic'; -4135 = 'Manual'; 2 = 'SemiAutomatic' }
    $previous = @{
        screenUpdating = [bool]$app.ScreenUpdating
        calculation    = $calcReverse[[int]$app.Calculation]
        enableEvents   = [bool]$app.EnableEvents
    }

    if ($PSBoundParameters.ContainsKey('ScreenUpdating')) {
        $app.ScreenUpdating = $ScreenUpdating
    }
    if ($PSBoundParameters.ContainsKey('Calculation')) {
        $app.Calculation = [int]$script:XL_CALCULATION[$Calculation.ToLower()]
    }
    if ($PSBoundParameters.ContainsKey('EnableEvents')) {
        $app.EnableEvents = $EnableEvents
    }

    # Read back current state
    $calcName = $calcReverse[[int]$app.Calculation]
    if (-not $calcName) { $calcName = [string]$app.Calculation }

    $result = @{
        status         = 'ok'
        previous       = $previous
        screenUpdating = [bool]$app.ScreenUpdating
        calculation    = $calcName
        enableEvents   = [bool]$app.EnableEvents
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Invoke-ExcelCalculate {
    <#
    .SYNOPSIS
        Trigger Excel recalculation.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Full
        Force full recalculation of all open workbooks.
    .PARAMETER SheetName
        Calculate a specific worksheet only.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-ExcelCalculate -WorkbookPath C:\data.xlsx -Full -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$Full,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    if ($Full) {
        $app.CalculateFull()
        $scope = 'full'
    } elseif (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $ws = $app.ActiveWorkbook.Worksheets.Item($SheetName)
        $ws.Calculate()
        $scope = "sheet:$SheetName"
    } else {
        $app.Calculate()
        $scope = 'workbook'
    }

    $result = @{
        status = 'ok'
        scope  = $scope
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Invoke-ExcelFunction {
    <#
    .SYNOPSIS
        Call an Excel WorksheetFunction method.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER FunctionName
        Name of the WorksheetFunction method (e.g. "SumIf", "VLookup", "CountIf").
    .PARAMETER Arguments
        Array of arguments — scalars or range address strings.
        Range strings like "A1:A10" or "Sheet1!B2:B50" are auto-resolved to COM Range objects.
    .PARAMETER SheetName
        Default worksheet for resolving bare range addresses (e.g. "A1:A10").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-ExcelFunction -WorkbookPath C:\data.xlsx -FunctionName "Sum" -Arguments @("Sheet1!A1:A10") -AsJson
    .EXAMPLE
        Invoke-ExcelFunction -WorkbookPath C:\data.xlsx -FunctionName "VLookup" -Arguments @("SearchVal","Sheet1!A1:D100",3,$false) -SheetName Sheet1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$FunctionName,
        [object[]]$Arguments,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    # Default worksheet for bare range refs
    $defaultWs = if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $wb.Worksheets.Item($SheetName)
    } else {
        $wb.ActiveSheet
    }

    # Resolve range references in arguments
    $resolvedArgs = @()
    foreach ($arg in $Arguments) {
        if ($arg -is [string] -and $arg -match "^'?(.+?)'?\!([A-Z]+\d+.*)$") {
            # Sheet-qualified reference: "Sheet1!A1:A10"
            $sheetRef = $matches[1]
            $cellRef  = $matches[2]
            $resolvedArgs += $wb.Worksheets.Item($sheetRef).Range($cellRef)
        } elseif ($arg -is [string] -and $arg -match '^[A-Z]+\d+(:[A-Z]+\d+)?$') {
            # Bare range reference: "A1:A10"
            $resolvedArgs += $defaultWs.Range($arg)
        } else {
            $resolvedArgs += $arg
        }
    }

    try {
        $wf     = $app.WorksheetFunction
        $result = $wf.GetType().InvokeMember(
            $FunctionName,
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null,
            $wf,
            $resolvedArgs
        )
    } catch {
        throw "WorksheetFunction.$FunctionName failed: $_"
    }

    $output = @{
        status   = 'ok'
        function = $FunctionName
        result   = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}

function Invoke-ExcelEvaluate {
    <#
    .SYNOPSIS
        Evaluate an Excel formula expression without writing to a cell.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Expression
        Excel formula expression (without leading =). E.g. "SUM(A1:A10)".
    .PARAMETER SheetName
        Worksheet context for cell references.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-ExcelEvaluate -WorkbookPath C:\data.xlsx -Expression "SUM(A1:A10)" -SheetName Sheet1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Expression,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $ws     = $app.ActiveWorkbook.Worksheets.Item($SheetName)
        $result = $ws.Evaluate($Expression)
    } else {
        $result = $app.Evaluate($Expression)
    }

    # Handle COM error values
    if ($result -is [int] -and $result -gt 2000) {
        $errorMap = @{
            2000 = '#NULL!'
            2007 = '#DIV/0!'
            2015 = '#VALUE!'
            2023 = '#REF!'
            2029 = '#NAME?'
            2036 = '#NUM!'
            2042 = '#N/A'
        }
        if ($errorMap.ContainsKey($result)) { $result = $errorMap[$result] }
    }

    $output = @{
        status     = 'ok'
        expression = $Expression
        result     = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}
