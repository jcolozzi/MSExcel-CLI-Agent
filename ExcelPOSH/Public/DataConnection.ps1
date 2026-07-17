# Public/DataConnection.ps1 — Workbook data connection refresh / add / remove

function Update-ExcelDataConnection {
    <#
    .SYNOPSIS
        Refresh one workbook data connection, or refresh all data in the workbook.
    .DESCRIPTION
        With -Name, calls WorkbookConnection.Refresh() for that connection. Without -Name
        (or with -All), calls Workbook.RefreshAll() which refreshes all external data ranges,
        connections and PivotTables.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Optional: name of a single connection to refresh.
    .PARAMETER All
        Refresh all queries/connections/pivots in the workbook (Workbook.RefreshAll).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-ExcelDataConnection -WorkbookPath C:\data.xlsx -Name "SalesConn" -AsJson
    .EXAMPLE
        Update-ExcelDataConnection -WorkbookPath C:\data.xlsx -All -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$Name,
        [switch]$All,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        if (-not $All -and -not [string]::IsNullOrWhiteSpace($Name)) {
            $wb.Connections.Item($Name).Refresh()
            $scope = "connection:$Name"
        } else {
            $wb.RefreshAll()
            $scope = 'all'
        }
        $result = @{
            status = 'ok'
            scope  = $scope
        }
    } catch {
        $result = @{
            status = 'error'
            error  = "Failed to refresh: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelDataConnection {
    <#
    .SYNOPSIS
        Delete a data connection from a workbook.
    .DESCRIPTION
        Calls WorkbookConnection.Delete(). Deleting a connection does not remove any table
        already loaded from it; use Remove-ExcelTable for that.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the connection to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelDataConnection -WorkbookPath C:\data.xlsx -Name "SalesConn" -AsJson
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
        $wb.Connections.Item($Name).Delete()
        $result = @{
            status  = 'ok'
            name    = $Name
            deleted = $true
        }
    } catch {
        $result = @{
            status = 'error'
            name   = $Name
            error  = "Failed to delete connection: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelDataConnection {
    <#
    .SYNOPSIS
        Add a new data connection to a workbook.
    .DESCRIPTION
        Calls Workbook.Connections.Add2(Name, Description, ConnectionString, CommandText, CommandType).
        CommandType maps to XlCmdType (Cube=1, Sql=2, Table=3, Default=4, List=5). Add2 is
        available in Excel 2010+; on versions/configurations where it fails, status='error' is
        returned with the underlying error rather than throwing.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the new connection.
    .PARAMETER ConnectionString
        The connection string (e.g. 'OLEDB;Provider=SQLOLEDB;Data Source=srv;Initial Catalog=db;Integrated Security=SSPI').
    .PARAMETER CommandText
        The command text (e.g. a SQL statement or table name) executed against the connection.
    .PARAMETER Description
        Optional description of the connection.
    .PARAMETER CommandType
        How CommandText is interpreted: Default, Sql, Table, Cube, or List. Defaults to Sql.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelDataConnection -WorkbookPath C:\data.xlsx -Name "SalesConn" -ConnectionString "OLEDB;Provider=SQLOLEDB;Data Source=srv;Initial Catalog=db;Integrated Security=SSPI" -CommandText "SELECT * FROM Sales" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ConnectionString,
        [Parameter(Mandatory)][string]$CommandText,
        [string]$Description,
        [ValidateSet('Default','Sql','Table','Cube','List')][string]$CommandType = 'Sql',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $cmdMap = @{ Cube = 1; Sql = 2; Table = 3; Default = 4; List = 5 }
        $null = $wb.Connections.Add2($Name, $Description, $ConnectionString, $CommandText, $cmdMap[$CommandType])
        $result = @{
            status  = 'ok'
            name    = $Name
            created = $true
        }
    } catch {
        $result = @{
            status = 'error'
            name   = $Name
            error  = "Failed to add connection: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
