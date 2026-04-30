# Public/DataValidationOps.ps1 — Data validation rules

function Set-ExcelDataValidation {
    <#
    .SYNOPSIS
        Add data validation to a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address to apply validation to.
    .PARAMETER ValidationType
        Type of validation: List, WholeNumber, Decimal, Date, Time, TextLength, Custom.
    .PARAMETER Operator
        Comparison operator: between, notbetween, equal, notequal, greater, greaterequal, less, lessequal.
    .PARAMETER Formula1
        Primary constraint value. For List type, comma-separated values or a range reference.
    .PARAMETER Formula2
        Secondary constraint value (used with between/notbetween operators).
    .PARAMETER InputTitle
        Title for the input prompt shown when the cell is selected.
    .PARAMETER InputMessage
        Body text for the input prompt.
    .PARAMETER ErrorTitle
        Title for the error alert.
    .PARAMETER ErrorMessage
        Body text for the error alert.
    .PARAMETER ErrorStyle
        Error alert style: stop, warning, information. Default: stop.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelDataValidation -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B2:B100" -ValidationType List -Formula1 "Yes,No,N/A" -AsJson
    .EXAMPLE
        Set-ExcelDataValidation -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "C2:C100" -ValidationType WholeNumber -Operator between -Formula1 "1" -Formula2 "100" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)]
        [ValidateSet('List','WholeNumber','Decimal','Date','Time','TextLength','Custom')]
        [string]$ValidationType,
        [ValidateSet('between','notbetween','equal','notequal','greater','greaterequal','less','lessequal')]
        [string]$Operator,
        [Parameter(Mandatory)][string]$Formula1,
        [string]$Formula2,
        [string]$InputTitle,
        [string]$InputMessage,
        [string]$ErrorTitle,
        [string]$ErrorMessage,
        [ValidateSet('stop','warning','information')]
        [string]$ErrorStyle = 'stop',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    # xlValidate constants
    $typeMap = @{
        'List'       = 3
        'WholeNumber'= 1
        'Decimal'    = 2
        'Date'       = 4
        'Time'       = 5
        'TextLength' = 6
        'Custom'     = 7
    }

    # xlOperator constants
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

    # xlErrorStyle constants
    $errorStyleMap = @{
        'stop'        = 1
        'warning'     = 2
        'information' = 3
    }

    $typeConst = $typeMap[$ValidationType]

    # Delete any existing validation on the range
    try { $rng.Validation.Delete() } catch { }

    # Add validation
    $missing = [System.Reflection.Missing]::Value
    if ($ValidationType -eq 'List') {
        $rng.Validation.Add($typeConst, $missing, $missing, $Formula1)
    } else {
        $operatorConst = if (-not [string]::IsNullOrWhiteSpace($Operator)) { $operatorMap[$Operator] } else { $missing }
        $f2 = if (-not [string]::IsNullOrWhiteSpace($Formula2)) { $Formula2 } else { $missing }
        $rng.Validation.Add($typeConst, $operatorConst, $missing, $Formula1, $f2)
    }

    # Set input message properties
    if (-not [string]::IsNullOrWhiteSpace($InputTitle))   { $rng.Validation.InputTitle   = $InputTitle }
    if (-not [string]::IsNullOrWhiteSpace($InputMessage)) { $rng.Validation.InputMessage = $InputMessage }

    # Set error alert properties
    if (-not [string]::IsNullOrWhiteSpace($ErrorTitle))   { $rng.Validation.ErrorTitle   = $ErrorTitle }
    if (-not [string]::IsNullOrWhiteSpace($ErrorMessage)) { $rng.Validation.ErrorMessage = $ErrorMessage }
    $rng.Validation.ErrorStyle = $errorStyleMap[$ErrorStyle]

    $result = @{
        status         = 'added'
        range          = $Range
        validationType = $ValidationType
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelDataValidation {
    <#
    .SYNOPSIS
        Get validation rules on a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address to inspect.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelDataValidation -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B2" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    try {
        $val = $rng.Validation
        $valType = $val.Type

        $result = @{
            status       = 'ok'
            hasValidation = $true
            range        = $Range
            type         = $valType
            operator     = $val.Operator
            formula1     = $val.Formula1
            formula2     = $val.Formula2
            inputTitle   = $val.InputTitle
            inputMessage = $val.InputMessage
            errorTitle   = $val.ErrorTitle
            errorMessage = $val.ErrorMessage
            errorStyle   = $val.ErrorStyle
        }
    } catch {
        $result = @{
            status        = 'ok'
            hasValidation = $false
            range         = $Range
        }
    }

    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelDataValidation {
    <#
    .SYNOPSIS
        Remove validation from a range.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell or range address to remove validation from.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelDataValidation -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "B2:B100" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $rng = $ws.Range($Range)

    $rng.Validation.Delete()

    $result = @{
        status = 'removed'
        range  = $Range
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
