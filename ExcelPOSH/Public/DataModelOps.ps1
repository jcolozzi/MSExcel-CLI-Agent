# Public/DataModelOps.ps1 — Power Pivot / Data Model (Workbook.Model) automation
# Requires Excel with Power Pivot / Data Model support. The Workbook.Model object may not
# exist on all editions; read/refresh paths return status='unsupported' when it is absent.

function Get-ExcelDataModel {
    <#
    .SYNOPSIS
        Inspect the Power Pivot Data Model (Workbook.Model) of a workbook.
    .DESCRIPTION
        Enumerates the Data Model's tables, measures, and relationship count via
        Workbook.Model. Returns status='unsupported' on editions without the Data Model
        object model rather than throwing.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelDataModel -WorkbookPath C:\data.xlsx -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $model = $wb.Model
        $tables = @()
        foreach ($t in $model.ModelTables) {
            $tables += @{ name = $t.Name }
        }
        $measures = @()
        foreach ($m in $model.ModelMeasures) {
            $measures += @{
                name    = $m.Name
                formula = $m.Formula
                table   = $m.AssociatedTable.Name
            }
        }
        $relationshipCount = $model.ModelRelationships.Count
        $result = @{
            status            = 'ok'
            tableCount        = $tables.Count
            tables            = $tables
            measureCount      = $measures.Count
            measures          = $measures
            relationshipCount = $relationshipCount
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Data Model unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelModelMeasure {
    <#
    .SYNOPSIS
        Add a DAX measure to the Power Pivot Data Model.
    .DESCRIPTION
        Calls Workbook.Model.ModelMeasures.Add(Name, AssociatedTable, Formula, FormatInformation
        [, Description]). The format is chosen from the Model's ModelFormat* objects via -FormatType.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER MeasureName
        Name of the new measure.
    .PARAMETER TableName
        Name of the Data Model table to associate the measure with.
    .PARAMETER Formula
        The DAX formula for the measure (e.g. 'SUM(Sales[Amount])').
    .PARAMETER FormatType
        Number format for the measure. One of General, Currency, DecimalNumber, WholeNumber,
        PercentageNumber, ScientificNumber, Boolean, Date. Default General.
    .PARAMETER Description
        Optional description of the measure.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelModelMeasure -WorkbookPath C:\data.xlsx -MeasureName "Total Sales" -TableName "Sales" -Formula "SUM(Sales[Amount])" -FormatType Currency -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$MeasureName,
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string]$Formula,
        [ValidateSet('General','Currency','DecimalNumber','WholeNumber','PercentageNumber','ScientificNumber','Boolean','Date')]
        [string]$FormatType = 'General',
        [string]$Description,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $model = $wb.Model
        $tbl   = $model.ModelTables.Item($TableName)
        switch ($FormatType) {
            'General'          { $fmt = $model.ModelFormatGeneral }
            'Currency'         { $fmt = $model.ModelFormatCurrency }
            'DecimalNumber'    { $fmt = $model.ModelFormatDecimalNumber }
            'WholeNumber'      { $fmt = $model.ModelFormatWholeNumber }
            'PercentageNumber' { $fmt = $model.ModelFormatPercentageNumber }
            'ScientificNumber' { $fmt = $model.ModelFormatScientificNumber }
            'Boolean'          { $fmt = $model.ModelFormatBoolean }
            'Date'             { $fmt = $model.ModelFormatDate }
        }
        if ([string]::IsNullOrWhiteSpace($Description)) {
            $null = $model.ModelMeasures.Add($MeasureName, $tbl, $Formula, $fmt)
        } else {
            $null = $model.ModelMeasures.Add($MeasureName, $tbl, $Formula, $fmt, $Description)
        }
        $result = @{
            status  = 'ok'
            measure = $MeasureName
            table   = $TableName
        }
    } catch {
        $result = @{
            status  = 'error'
            measure = $MeasureName
            error   = "Failed to add measure: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelModelMeasure {
    <#
    .SYNOPSIS
        Delete a DAX measure from the Power Pivot Data Model.
    .DESCRIPTION
        Iterates Workbook.Model.ModelMeasures and calls ModelMeasure.Delete() on the measure
        whose Name matches -MeasureName. Returns status='error' if no such measure exists.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER MeasureName
        Name of the measure to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelModelMeasure -WorkbookPath C:\data.xlsx -MeasureName "Total Sales" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$MeasureName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $found = $false
        foreach ($m in $wb.Model.ModelMeasures) {
            if ($m.Name -eq $MeasureName) {
                $m.Delete()
                $found = $true
                break
            }
        }
        if ($found) {
            $result = @{
                status  = 'ok'
                measure = $MeasureName
                deleted = $found
            }
        } else {
            $result = @{
                status  = 'error'
                measure = $MeasureName
                error   = 'Measure not found'
            }
        }
    } catch {
        $result = @{
            status  = 'error'
            measure = $MeasureName
            error   = "Failed to delete measure: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelModelRelationship {
    <#
    .SYNOPSIS
        Create a relationship between two Data Model tables.
    .DESCRIPTION
        Calls Workbook.Model.ModelRelationships.Add(ForeignKeyColumn, PrimaryKeyColumn). Columns
        are resolved via Model.ModelTables.Item(table).ModelTableColumns.Item(column).
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER ForeignKeyTable
        Name of the table on the many side (holding the foreign key).
    .PARAMETER ForeignKeyColumn
        Name of the foreign key column in ForeignKeyTable.
    .PARAMETER PrimaryKeyTable
        Name of the table on the one side (holding the primary key).
    .PARAMETER PrimaryKeyColumn
        Name of the primary key column in PrimaryKeyTable.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelModelRelationship -WorkbookPath C:\data.xlsx -ForeignKeyTable "Sales" -ForeignKeyColumn "ProductID" -PrimaryKeyTable "Products" -PrimaryKeyColumn "ID" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$ForeignKeyTable,
        [Parameter(Mandatory)][string]$ForeignKeyColumn,
        [Parameter(Mandatory)][string]$PrimaryKeyTable,
        [Parameter(Mandatory)][string]$PrimaryKeyColumn,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $model = $wb.Model
        $fk = $model.ModelTables.Item($ForeignKeyTable).ModelTableColumns.Item($ForeignKeyColumn)
        $pk = $model.ModelTables.Item($PrimaryKeyTable).ModelTableColumns.Item($PrimaryKeyColumn)
        $null = $model.ModelRelationships.Add($fk, $pk)
        $result = @{
            status = 'ok'
            from   = "$ForeignKeyTable.$ForeignKeyColumn"
            to     = "$PrimaryKeyTable.$PrimaryKeyColumn"
        }
    } catch {
        $result = @{
            status = 'error'
            error  = "Failed to add relationship: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Update-ExcelDataModel {
    <#
    .SYNOPSIS
        Refresh the entire Power Pivot Data Model.
    .DESCRIPTION
        Calls Workbook.Model.Refresh(), which reprocesses all Data Model tables from their
        sources. Returns status='unsupported' on editions without the Data Model object model.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-ExcelDataModel -WorkbookPath C:\data.xlsx -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $wb.Model.Refresh()
        $result = @{
            status = 'ok'
            scope  = 'model'
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Data Model unavailable: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
