# Public/PowerQueryOps.ps1 — Power Query (Get & Transform) automation via Workbook.Queries
# Requires Excel 2016+ (Workbook.Queries collection). Queries are authored as M formula strings.

function Get-ExcelPowerQuery {
    <#
    .SYNOPSIS
        List Power Query (M) queries in a workbook, or get one by name.
    .DESCRIPTION
        Enumerates the Workbook.Queries collection (Excel 2016+). Each query exposes its
        name, description, and Power Query M formula. Returns status='unsupported' on
        Excel versions without the Queries object model.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Optional: return only the query with this name.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelPowerQuery -WorkbookPath C:\data.xlsx -AsJson
    .EXAMPLE
        Get-ExcelPowerQuery -WorkbookPath C:\data.xlsx -Name "SalesData" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$Name,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $queries = @()
        foreach ($q in $wb.Queries) {
            if (-not [string]::IsNullOrWhiteSpace($Name) -and $q.Name -ne $Name) { continue }
            $queries += @{
                name        = $q.Name
                description = (ConvertTo-ExcelSafeValue $q.Description)
                formula     = $q.Formula
            }
        }
        $result = @{
            status  = 'ok'
            count   = $queries.Count
            queries = $queries
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Power Query object model unavailable (requires Excel 2016+): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function New-ExcelPowerQuery {
    <#
    .SYNOPSIS
        Create a new Power Query from a Power Query M formula string.
    .DESCRIPTION
        Calls Workbook.Queries.Add(Name, Formula, Description). The new query is created
        as connection-only. Use Import-ExcelPowerQueryToTable to load its results to a sheet,
        or Update-ExcelPowerQuery to refresh it.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the new query.
    .PARAMETER Formula
        The Power Query M formula (e.g. 'let Source = Csv.Document(File.Contents("C:\d.csv")) in Source').
    .PARAMETER Description
        Optional description of the query.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelPowerQuery -WorkbookPath C:\data.xlsx -Name "Load" -Formula 'let Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content] in Source' -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Formula,
        [string]$Description,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        if ([string]::IsNullOrWhiteSpace($Description)) {
            $null = $wb.Queries.Add($Name, $Formula)
        } else {
            $null = $wb.Queries.Add($Name, $Formula, $Description)
        }
        $result = @{
            status  = 'ok'
            name    = $Name
            created = $true
        }
    } catch {
        $result = @{
            status = 'error'
            name   = $Name
            error  = "Failed to add query (requires Excel 2016+): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelPowerQuery {
    <#
    .SYNOPSIS
        Update the M formula and/or description of an existing Power Query.
    .DESCRIPTION
        WorkbookQuery.Formula is read/write, so the M is updated in place (no delete/re-add).
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the query to update.
    .PARAMETER Formula
        New Power Query M formula. Omit to leave unchanged.
    .PARAMETER Description
        New description. Omit to leave unchanged.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelPowerQuery -WorkbookPath C:\data.xlsx -Name "Load" -Formula 'let Source = 1 in Source' -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$Name,
        [string]$Formula,
        [string]$Description,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    try {
        $q = $wb.Queries.Item($Name)
        $changed = @()
        if ($PSBoundParameters.ContainsKey('Formula')) {
            $q.Formula = $Formula
            $changed += 'formula'
        }
        if ($PSBoundParameters.ContainsKey('Description')) {
            $q.Description = $Description
            $changed += 'description'
        }
        $result = @{
            status  = 'ok'
            name    = $Name
            updated = $changed
        }
    } catch {
        $result = @{
            status = 'error'
            name   = $Name
            error  = "Failed to update query: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelPowerQuery {
    <#
    .SYNOPSIS
        Delete a Power Query from a workbook.
    .DESCRIPTION
        Calls WorkbookQuery.Delete(). Deleting a query does not remove any table already
        loaded from it; use Remove-ExcelTable / Remove-ExcelDataConnection for that.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Name of the query to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelPowerQuery -WorkbookPath C:\data.xlsx -Name "Load" -AsJson
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
        $wb.Queries.Item($Name).Delete()
        $result = @{
            status  = 'ok'
            name    = $Name
            deleted = $true
        }
    } catch {
        $result = @{
            status = 'error'
            name   = $Name
            error  = "Failed to delete query: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Update-ExcelPowerQuery {
    <#
    .SYNOPSIS
        Refresh one Power Query, or refresh all data in the workbook.
    .DESCRIPTION
        With -Name, calls WorkbookQuery.Refresh() for that query. Without -Name (or with -All),
        calls Workbook.RefreshAll() which refreshes all external data ranges and PivotTables.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER Name
        Optional: name of a single query to refresh.
    .PARAMETER All
        Refresh all queries/connections/pivots in the workbook (Workbook.RefreshAll).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Update-ExcelPowerQuery -WorkbookPath C:\data.xlsx -Name "Load" -AsJson
    .EXAMPLE
        Update-ExcelPowerQuery -WorkbookPath C:\data.xlsx -All -AsJson
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
            $wb.Queries.Item($Name).Refresh()
            $scope = "query:$Name"
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

function Import-ExcelPowerQueryToTable {
    <#
    .SYNOPSIS
        Load a connection-only Power Query's results into a worksheet table (ListObject).
    .DESCRIPTION
        Creates a ListObject bound to the query via the Microsoft.Mashup.OleDb.1 provider and
        refreshes it. This path depends on the Power Query mashup engine; on failure it returns
        status='unsupported' with the underlying error rather than throwing.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet that will host the loaded table.
    .PARAMETER QueryName
        Name of an existing Power Query (see New-ExcelPowerQuery / Get-ExcelPowerQuery).
    .PARAMETER Destination
        Top-left cell for the table (default "A1").
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Import-ExcelPowerQueryToTable -WorkbookPath C:\data.xlsx -SheetName Sheet1 -QueryName "Load" -Destination "A1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$QueryName,
        [string]$Destination = 'A1',
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    try {
        $connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook;Location=$QueryName;Extended Properties=`"`""
        $lo = $ws.ListObjects.Add(
            [int]$script:XL_SOURCE_TYPE.external,
            $connString,
            $false,
            [int]$script:XL_YES_NO_GUESS.yes,
            $ws.Range($Destination)
        )
        $lo.QueryTable.CommandType = 2   # xlCmdSql
        $lo.QueryTable.CommandText = "SELECT * FROM [$QueryName]"
        $lo.QueryTable.Refresh($false)
        $result = @{
            status = 'ok'
            sheet  = $SheetName
            query  = $QueryName
            table  = $lo.Name
        }
    } catch {
        $result = @{
            status = 'unsupported'
            query  = $QueryName
            error  = "Load-to-table via mashup OLEDB failed: $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
