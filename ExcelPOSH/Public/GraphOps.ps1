# Public/GraphOps.ps1 — Excel workbook dependency + data-relationship graph export

# ── xlConnectionType (subset) ──
$script:XL_CONNECTION_TYPE = @{
    1 = 'OLEDB'; 2 = 'ODBC'; 3 = 'XMLMAP'; 4 = 'TEXT'; 5 = 'WEB'
    6 = 'MODEL'; 7 = 'WORKSHEET'; 8 = 'NOSOURCE'; 9 = 'DATAFEED'
}
# ── xlSlicerCacheType ──
$script:XL_SLICER_CACHE_TYPE = @{ 1 = 'Slicer'; 2 = 'Timeline' }

# ──────────────────────────────────────────────────────────────────────
#  Internal scan helpers (not exported)
# ──────────────────────────────────────────────────────────────────────

function Get-ExcelColumnStats {
    <#
    .SYNOPSIS
        Read a ListObject's data and compute per-column value statistics for PK/FK inference.
    .OUTPUTS
        Array of [ordered]@{ name; distinct(HashSet[string]); total; blanks; isPk }
    #>
    param($ListObject, [int]$MaxRows = 5000)

    $result = New-Object 'System.Collections.Generic.List[object]'
    $colNames = @()
    foreach ($c in $ListObject.ListColumns) { $colNames += [string]$c.Name }
    $colCount = $colNames.Count
    if ($colCount -eq 0) { return $result }

    $dbr = $null
    try { $dbr = $ListObject.DataBodyRange } catch {}
    if ($null -eq $dbr) {
        foreach ($n in $colNames) {
            $result.Add([ordered]@{ name = $n; distinct = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase); total = 0; blanks = 0; isPk = $false })
        }
        return $result
    }

    $rowCount = [int]$dbr.Rows.Count
    $readRows = [Math]::Min($rowCount, $MaxRows)
    # NOTE: assign the Range with a plain statement — `$x = if(){}else{$dbr}` would emit the
    # COM Range to the output stream, which PowerShell enumerates into cells (breaking Value2).
    $readRange = $dbr
    if ($rowCount -gt $MaxRows) { $readRange = $dbr.Resize($MaxRows, $colCount) }
    $vals = $readRange.Value2

    # Prepare per-column accumulators
    $dist = @()
    $blanks = @(0) * $colCount
    for ($c = 0; $c -lt $colCount; $c++) {
        $dist += , ([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase))
    }

    if ($readRows -eq 1 -and $colCount -eq 1) {
        # Value2 is a scalar
        $v = $vals
        if ($null -eq $v -or ([string]$v).Trim() -eq '') { $blanks[0]++ } else { [void]$dist[0].Add([string]$v) }
    }
    else {
        for ($r = 1; $r -le $readRows; $r++) {
            for ($c = 1; $c -le $colCount; $c++) {
                $v = $vals[$r, $c]
                if ($null -eq $v -or ([string]$v).Trim() -eq '') { $blanks[$c - 1]++ }
                else { [void]$dist[$c - 1].Add([string]$v) }
            }
        }
    }

    for ($c = 0; $c -lt $colCount; $c++) {
        $isPk = ($blanks[$c] -eq 0) -and ($readRows -gt 0) -and ($dist[$c].Count -eq $readRows)
        $result.Add([ordered]@{
                name     = $colNames[$c]
                distinct = $dist[$c]
                total    = $readRows
                blanks   = $blanks[$c]
                isPk     = $isPk
            })
    }
    return $result
}

function Resolve-ExcelFormulaTarget {
    <#
    .SYNOPSIS
        Resolve structural references in a formula to graph node ids.
    .OUTPUTS
        Array of node-id strings.
    #>
    param(
        [string]$Formula,
        [hashtable]$RefLookup,      # lowercased name -> nodeId (tables, names, sheets)
        [regex]$NamesRegex,         # bare defined-name matcher (may be $null)
        [hashtable]$GraphState,
        [string]$SelfId
    )

    $ids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ref in (Get-ExcelFormulaReference -Formula $Formula)) {
        $key = $ref.name.ToLowerInvariant()
        if ($ref.kind -eq 'external') {
            $extId = Get-GraphObjectId -Group 'external' -Name $ref.name
            if (-not $GraphState.NodeIndex.ContainsKey($extId)) {
                Add-GraphNode -GraphState $GraphState -Id $extId -Label $ref.name -Group 'external' -Layer 'structure' -Meta @{ kind = 'external workbook' } | Out-Null
            }
            [void]$ids.Add($extId)
        }
        elseif ($RefLookup.ContainsKey($key)) {
            [void]$ids.Add($RefLookup[$key])
        }
    }

    if ($null -ne $NamesRegex) {
        foreach ($m in $NamesRegex.Matches($Formula)) {
            $key = $m.Value.ToLowerInvariant()
            if ($RefLookup.ContainsKey($key)) { [void]$ids.Add($RefLookup[$key]) }
        }
    }

    if ($SelfId) { [void]$ids.Remove($SelfId) }
    return @($ids)
}

# ──────────────────────────────────────────────────────────────────────
#  Export-ExcelGraph
# ──────────────────────────────────────────────────────────────────────

function Export-ExcelGraph {
    <#
    .SYNOPSIS
        Build a dependency + data-relationship graph of an Excel workbook and output
        graph.json plus an interactive HTML viewer.

    .DESCRIPTION
        Scans worksheets, tables (ListObjects), named ranges, PivotTables, charts, connections,
        Power Query queries, the Data Model (Power Pivot), slicers/timelines, and VBA modules.

        Builds two graph layers surfaced by a Structure/Data toggle in the viewer:
          * STRUCTURE — object containment and reference edges (formulas, chart/pivot sources,
            connections, queries, slicers, VBA code references).
          * DATA — the actual data relationships: Data Model foreign keys, lookup-formula
            relationships, value-overlap inferred foreign keys, and primary-key detection.

        Outputs graph.json and an embedded vis.js index.html.

    .PARAMETER WorkbookPath
        Path to the Excel workbook (.xlsx/.xlsm/.xlsb/.xls).

    .PARAMETER OutDir
        Output directory for graph.json and index.html. Default: .\excel-graph-out

    .PARAMETER FormulaMode
        Formula-edge granularity. None = no formula edges; Aggregate = sheet-level reference
        edges; Cell = per formula-cell nodes (capped); Both = aggregate + cell. Default: Aggregate.

    .PARAMETER MaxFormulaCells
        Cap on per-cell formula nodes created in Cell/Both mode. Default: 500.

    .PARAMETER FkOverlapThreshold
        Minimum value-overlap ratio (0-1) to emit an inferred foreign key. Default: 0.85.

    .PARAMETER MaxOverlapRows
        Max rows sampled per table for PK/FK value inference. Default: 5000.

    .PARAMETER DisableDataGraph
        Skip the data-relationship layer (Data Model FKs, lookups, value-overlap, PK detection).

    .PARAMETER DisableVbaHeuristics
        Skip VBA module analysis. VBA analysis requires 'Trust access to the VBA project object
        model'; it degrades to a warning if the project is not accessible.

    .PARAMETER IncludeEmptyModules
        Include VBA modules that contain no code (e.g. the empty ThisWorkbook / SheetN document
        modules). By default these are skipped to reduce noise.

    .PARAMETER SkipViewerCopy
        Skip generating the embedded HTML viewer (index.html).

    .PARAMETER PassThru
        Return the graph object to the pipeline (in addition to writing files).

    .PARAMETER Quiet
        Suppress Write-Progress output.

    .EXAMPLE
        Export-ExcelGraph -WorkbookPath C:\Sales.xlsx

    .EXAMPLE
        $g = Export-ExcelGraph -WorkbookPath C:\Sales.xlsm -FormulaMode Both -PassThru
        $g.nodes | Where-Object group -eq 'table'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$WorkbookPath,

        [string]$OutDir = '.\excel-graph-out',

        [ValidateSet('None', 'Aggregate', 'Cell', 'Both')]
        [string]$FormulaMode = 'Aggregate',

        [int]$MaxFormulaCells = 500,

        [ValidateRange(0.0, 1.0)]
        [double]$FkOverlapThreshold = 0.85,

        [int]$MaxOverlapRows = 5000,

        [switch]$DisableDataGraph,

        [switch]$DisableVbaHeuristics,

        [switch]$IncludeEmptyModules,

        [switch]$SkipViewerCopy,

        [switch]$PassThru,

        [switch]$Quiet
    )

    $resolved = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw [System.IO.FileNotFoundException]::new("Workbook file not found: $resolved", $resolved)
    }

    if (-not (Test-Path -LiteralPath $OutDir)) { $null = New-Item -ItemType Directory -Force -Path $OutDir }
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path

    $app = Connect-ExcelWorkbook -WorkbookPath $resolved
    $wb  = Get-ExcelGraphTargetWorkbook -App $app -ResolvedPath $resolved

    $gs = New-GraphState
    $doData = -not $DisableDataGraph

    # Lookups built during scan
    $refLookup = @{}                # lowercased name -> nodeId (for formula resolution)
    $definedNames = New-Object 'System.Collections.Generic.List[string]'
    $columnStatsByTable = @{}       # tableName -> column-stats list (for FK/PK inference)
    $tableNodeIdByName = @{}        # tableName (lc) -> table nodeId

    function _AddRef([string]$name, [string]$nodeId) {
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        $k = $name.ToLowerInvariant()
        if (-not $refLookup.ContainsKey($k)) { $refLookup[$k] = $nodeId }
    }

    try {
        # ── Workbook node ──
        $wbName = [string]$wb.Name
        $wbNodeId = Get-GraphObjectId -Group 'workbook' -Name $wbName
        Add-GraphNode -GraphState $gs -Id $wbNodeId -Label $wbName -Group 'workbook' -Layer 'structure' -Meta @{ path = $resolved } | Out-Null

        # ── Worksheets ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning worksheets...' -PercentComplete 5 }
        foreach ($ws in $wb.Worksheets) {
            $sheetName = [string]$ws.Name
            $sheetId = Get-GraphObjectId -Group 'sheet' -Name $sheetName
            $visible = try { [int]$ws.Visible -eq -1 } catch { $true }   # xlSheetVisible = -1
            Add-GraphNode -GraphState $gs -Id $sheetId -Label $sheetName -Group 'sheet' -Layer 'structure' -Meta @{ visible = $visible } | Out-Null
            Add-GraphEdge -GraphState $gs -From $wbNodeId -To $sheetId -Label 'contains' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $sheetName -NodeId $sheetId -Group 'sheet'
            _AddRef $sheetName $sheetId
        }

        # ── Tables (ListObjects) ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning tables...' -PercentComplete 15 }
        foreach ($ws in $wb.Worksheets) {
            $sheetName = [string]$ws.Name
            $sheetId = Get-GraphObjectId -Group 'sheet' -Name $sheetName
            foreach ($lo in $ws.ListObjects) {
                $tableName = [string]$lo.Name
                $tableId = Get-GraphObjectId -Group 'table' -Name $tableName
                $colNames = @(); foreach ($c in $lo.ListColumns) { $colNames += [string]$c.Name }
                $rowCount = try { if ($null -ne $lo.DataBodyRange) { [int]$lo.DataBodyRange.Rows.Count } else { 0 } } catch { 0 }
                Add-GraphNode -GraphState $gs -Id $tableId -Label $tableName -Group 'table' -Layer 'both' -Meta @{
                    sheet       = $sheetName
                    columns     = ($colNames -join ', ')
                    columnCount = $colNames.Count
                    rowCount    = $rowCount
                } | Out-Null
                Add-GraphEdge -GraphState $gs -From $sheetId -To $tableId -Label 'contains' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null
                Register-GraphNameTarget -GraphState $gs -Name $tableName -NodeId $tableId -Group 'table' -IsDataObject
                _AddRef $tableName $tableId
                $tableNodeIdByName[$tableName.ToLowerInvariant()] = $tableId

                if ($doData) {
                    $stats = Get-ExcelColumnStats -ListObject $lo -MaxRows $MaxOverlapRows
                    $columnStatsByTable[$tableName] = $stats
                    foreach ($cs in $stats) {
                        $colId = Get-GraphColumnId -OwnerName $tableName -ColumnName $cs.name
                        Add-GraphNode -GraphState $gs -Id $colId -Label $cs.name -Group 'column' -Layer 'data' -Meta @{
                            table        = $tableName
                            isPrimaryKey = $cs.isPk
                            distinct     = $cs.distinct.Count
                            rows         = $cs.total
                        } | Out-Null
                        Add-GraphEdge -GraphState $gs -From $tableId -To $colId -Label 'column' -Kind 'column-owner' -Arrows 'to' -Layer 'data' | Out-Null
                    }

                    # Lookup-formula inference from calculated table columns
                    foreach ($c in $lo.ListColumns) {
                        $f = $null
                        try {
                            $cell = $c.DataBodyRange.Cells.Item(1, 1)
                            if ($cell.HasFormula) { $f = [string]$cell.Formula }
                        } catch {}
                        if ($f -and (Test-ExcelLookupFormula -Formula $f)) {
                            foreach ($ref in (Get-ExcelFormulaReference -Formula $f)) {
                                if ($ref.kind -ne 'table') { continue }
                                $tk = $ref.name.ToLowerInvariant()
                                if ($tableNodeIdByName.ContainsKey($tk) -and $tableNodeIdByName[$tk] -ne $tableId) {
                                    Add-GraphEdge -GraphState $gs -From $tableId -To $tableNodeIdByName[$tk] -Label 'lookup' -Kind 'lookup-fk' -Arrows 'to' -Layer 'data' -Meta @{
                                        via     = $c.Name
                                        formula = (Get-GraphPreviewText -Text $f -MaxLength 120)
                                    } | Out-Null
                                }
                            }
                        }
                    }
                }
            }
        }

        # ── Named ranges ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning named ranges...' -PercentComplete 30 }
        foreach ($nm in $wb.Names) {
            $nmName = try { [string]$nm.Name } catch { continue }
            if ([string]::IsNullOrWhiteSpace($nmName)) { continue }
            $localName = $nmName -replace '^.*!', ''   # strip sheet-scope prefix
            if ($localName -match '^(Print_Area|Print_Titles|_FilterDatabase)$' -or $localName -like '_xl*') { continue }
            $isVisible = try { [bool]$nm.Visible } catch { $true }
            if (-not $isVisible) { continue }

            $refersTo = try { [string]$nm.RefersTo } catch { '' }
            $nameId = Get-GraphObjectId -Group 'name' -Name $localName
            Add-GraphNode -GraphState $gs -Id $nameId -Label $localName -Group 'name' -Layer 'structure' -Meta @{
                refersTo = (Get-GraphPreviewText -Text $refersTo -MaxLength 120)
            } | Out-Null
            Register-GraphNameTarget -GraphState $gs -Name $localName -NodeId $nameId -Group 'name'
            _AddRef $localName $nameId
            $definedNames.Add($localName)
        }

        # ── Data Model (Power Pivot) ──
        $modelOk = $false
        $model = $null
        try { $model = $wb.Model; if ($null -ne $model) { $modelOk = $true } } catch {}
        if ($modelOk) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning data model...' -PercentComplete 40 }
            try {
                foreach ($mt in $model.ModelTables) {
                    $mtName = [string]$mt.Name
                    $mtId = Get-GraphObjectId -Group 'modeltable' -Name $mtName
                    $mcols = @(); try { foreach ($mc in $mt.ModelTableColumns) { $mcols += [string]$mc.Name } } catch {}
                    Add-GraphNode -GraphState $gs -Id $mtId -Label $mtName -Group 'modeltable' -Layer 'both' -Meta @{
                        columns     = ($mcols -join ', ')
                        columnCount = $mcols.Count
                    } | Out-Null
                    if ($doData) {
                        foreach ($mcName in $mcols) {
                            $colId = Get-GraphColumnId -OwnerName $mtName -ColumnName $mcName
                            Add-GraphNode -GraphState $gs -Id $colId -Label $mcName -Group 'column' -Layer 'data' -Meta @{ table = $mtName } | Out-Null
                            Add-GraphEdge -GraphState $gs -From $mtId -To $colId -Label 'column' -Kind 'column-owner' -Arrows 'to' -Layer 'data' | Out-Null
                        }
                    }
                }
            } catch { Add-GraphWarning -GraphState $gs -Code 'ModelTableScanFailed' -Message ("Data Model table scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'modeltable' } }

            try {
                foreach ($ms in $model.ModelMeasures) {
                    $msName = [string]$ms.Name
                    $msId = Get-GraphObjectId -Group 'measure' -Name $msName
                    $assoc = try { [string]$ms.AssociatedTable.Name } catch { '' }
                    $formula = try { [string]$ms.Formula } catch { '' }
                    Add-GraphNode -GraphState $gs -Id $msId -Label $msName -Group 'measure' -Layer 'structure' -Meta @{
                        table   = $assoc
                        formula = (Get-GraphPreviewText -Text $formula -MaxLength 160)
                    } | Out-Null
                    if ($assoc) {
                        $mtId = Get-GraphObjectId -Group 'modeltable' -Name $assoc
                        if ($gs.NodeIndex.ContainsKey($mtId)) {
                            Add-GraphEdge -GraphState $gs -From $mtId -To $msId -Label 'measure' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null
                        }
                    }
                }
            } catch { Add-GraphWarning -GraphState $gs -Code 'ModelMeasureScanFailed' -Message ("Data Model measure scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'measure' } }

            # Data Model relationships (explicit FKs) — data layer
            if ($doData) {
                try {
                    foreach ($rel in $model.ModelRelationships) {
                        $fkTable = try { [string]$rel.ForeignKeyTable.Name } catch { '' }
                        $fkCol   = try { [string]$rel.ForeignKeyColumn.Name } catch { '' }
                        $pkTable = try { [string]$rel.PrimaryKeyTable.Name } catch { '' }
                        $pkCol   = try { [string]$rel.PrimaryKeyColumn.Name } catch { '' }
                        $active  = try { [bool]$rel.Active } catch { $true }
                        if (-not $fkTable -or -not $pkTable) { continue }

                        $fkColId = Get-GraphColumnId -OwnerName $fkTable -ColumnName $fkCol
                        $pkColId = Get-GraphColumnId -OwnerName $pkTable -ColumnName $pkCol
                        if (-not $gs.NodeIndex.ContainsKey($fkColId)) {
                            Add-GraphNode -GraphState $gs -Id $fkColId -Label $fkCol -Group 'column' -Layer 'data' -Meta @{ table = $fkTable } | Out-Null
                        }
                        if (-not $gs.NodeIndex.ContainsKey($pkColId)) {
                            Add-GraphNode -GraphState $gs -Id $pkColId -Label $pkCol -Group 'column' -Layer 'data' -Meta @{ table = $pkTable; isPrimaryKey = $true } | Out-Null
                        }
                        Add-GraphEdge -GraphState $gs -From $fkColId -To $pkColId -Label 'FK' -Kind 'datamodel-fk' -Arrows 'to' -Layer 'data' -Meta @{
                            fkTable = $fkTable; pkTable = $pkTable; active = $active
                        } | Out-Null
                    }
                } catch { Add-GraphWarning -GraphState $gs -Code 'ModelRelScanFailed' -Message ("Data Model relationship scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'modeltable' } }
            }
        }
        else {
            Add-GraphWarning -GraphState $gs -Code 'DataModelUnsupported' -Message 'Workbook.Model is unavailable (no Data Model or unsupported edition).' -Meta @{ group = 'modeltable' }
        }

        # ── Connections ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning connections...' -PercentComplete 50 }
        try {
            foreach ($conn in $wb.Connections) {
                $connName = [string]$conn.Name
                $connId = Get-GraphObjectId -Group 'connection' -Name $connName
                $ctype = try { [int]$conn.Type } catch { -1 }
                $ctypeName = if ($script:XL_CONNECTION_TYPE.ContainsKey($ctype)) { $script:XL_CONNECTION_TYPE[$ctype] } else { "Type$ctype" }
                $desc = try { [string]$conn.Description } catch { '' }
                Add-GraphNode -GraphState $gs -Id $connId -Label $connName -Group 'connection' -Layer 'structure' -Meta @{
                    type = $ctypeName; description = $desc
                } | Out-Null
                # Connection destination ranges -> tables it feeds
                try {
                    foreach ($rng in $conn.Ranges) {
                        $lo = $null; try { $lo = $rng.ListObject } catch {}
                        if ($null -ne $lo) {
                            $tId = Get-GraphObjectId -Group 'table' -Name ([string]$lo.Name)
                            if ($gs.NodeIndex.ContainsKey($tId)) {
                                Add-GraphEdge -GraphState $gs -From $connId -To $tId -Label 'feeds' -Kind 'connection-feeds' -Arrows 'to' -Layer 'structure' | Out-Null
                            }
                        }
                    }
                } catch {}
            }
        } catch { Add-GraphWarning -GraphState $gs -Code 'ConnectionScanFailed' -Message ("Connection scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'connection' } }

        # ── Power Query queries ──
        try {
            foreach ($q in $wb.Queries) {
                $qName = [string]$q.Name
                $qId = Get-GraphObjectId -Group 'query' -Name $qName
                $qFormula = try { [string]$q.Formula } catch { '' }
                Add-GraphNode -GraphState $gs -Id $qId -Label $qName -Group 'query' -Layer 'structure' -Meta @{
                    mHash   = Get-GraphTextHash -Text $qFormula
                    mPreview = Get-GraphPreviewText -Text $qFormula -MaxLength 200
                } | Out-Null
                # Power Query creates a connection named "Query - <name>"
                $connId = Get-GraphObjectId -Group 'connection' -Name ("Query - $qName")
                if ($gs.NodeIndex.ContainsKey($connId)) {
                    Add-GraphEdge -GraphState $gs -From $qId -To $connId -Label 'loads' -Kind 'query-loads' -Arrows 'to' -Layer 'structure' | Out-Null
                }
            }
        } catch { Add-GraphWarning -GraphState $gs -Code 'QueryScanFailed' -Message ("Power Query scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'query' } }

        # ── PivotTables ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning pivot tables...' -PercentComplete 60 }
        foreach ($ws in $wb.Worksheets) {
            $sheetName = [string]$ws.Name
            $sheetId = Get-GraphObjectId -Group 'sheet' -Name $sheetName
            foreach ($pt in $ws.PivotTables()) {
                $ptName = [string]$pt.Name
                $ptId = Get-GraphObjectId -Group 'pivot' -Name $ptName
                $srcData = try { [string]$pt.SourceData } catch { '' }
                Add-GraphNode -GraphState $gs -Id $ptId -Label $ptName -Group 'pivot' -Layer 'structure' -Meta @{
                    sheet = $sheetName; source = (Get-GraphPreviewText -Text $srcData -MaxLength 120)
                } | Out-Null
                Add-GraphEdge -GraphState $gs -From $sheetId -To $ptId -Label 'contains' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null

                # Source resolution
                $linked = $false
                if ($srcData) {
                    foreach ($ref in (Get-ExcelFormulaReference -Formula $srcData)) {
                        $k = $ref.name.ToLowerInvariant()
                        if ($refLookup.ContainsKey($k)) {
                            Add-GraphEdge -GraphState $gs -From $ptId -To $refLookup[$k] -Label 'source' -Kind 'pivot-source' -Arrows 'to' -Layer 'structure' | Out-Null
                            $linked = $true
                        }
                    }
                }
                try {
                    $cache = $pt.PivotCache()
                    $wc = $null; try { $wc = $cache.WorkbookConnection } catch {}
                    if ($null -ne $wc) {
                        $connId = Get-GraphObjectId -Group 'connection' -Name ([string]$wc.Name)
                        if ($gs.NodeIndex.ContainsKey($connId)) {
                            Add-GraphEdge -GraphState $gs -From $ptId -To $connId -Label 'source' -Kind 'pivot-source' -Arrows 'to' -Layer 'structure' | Out-Null
                            $linked = $true
                        }
                    }
                } catch {}
                if (-not $linked -and $srcData) {
                    Add-GraphWarning -GraphState $gs -Code 'PivotSourceUnresolved' -Message ("PivotTable '{0}' source not resolved: {1}" -f $ptName, (Get-GraphPreviewText -Text $srcData -MaxLength 80)) -Meta @{ owner = $ptName; group = 'pivot' }
                }
            }
        }

        # ── Charts (ChartObjects + chart sheets) ──
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning charts...' -PercentComplete 70 }
        $chartScan = {
            param($chartObj, $ownerSheetId, $ownerSheet)
            $chName = try { [string]$chartObj.Name } catch { 'Chart' }
            $chId = Get-GraphObjectId -Group 'chart' -Name ($ownerSheet + '!' + $chName)
            Add-GraphNode -GraphState $gs -Id $chId -Label $chName -Group 'chart' -Layer 'structure' -Meta @{ sheet = $ownerSheet } | Out-Null
            if ($ownerSheetId) { Add-GraphEdge -GraphState $gs -From $ownerSheetId -To $chId -Label 'contains' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null }
            try {
                foreach ($ser in $chartObj.SeriesCollection()) {
                    $sf = try { [string]$ser.Formula } catch { '' }
                    foreach ($tid in (Resolve-ExcelFormulaTarget -Formula $sf -RefLookup $refLookup -NamesRegex $null -GraphState $gs -SelfId $chId)) {
                        Add-GraphEdge -GraphState $gs -From $chId -To $tid -Label 'source' -Kind 'chart-source' -Arrows 'to' -Layer 'structure' | Out-Null
                    }
                }
            } catch {}
        }
        foreach ($ws in $wb.Worksheets) {
            $sheetName = [string]$ws.Name
            $sheetId = Get-GraphObjectId -Group 'sheet' -Name $sheetName
            foreach ($co in $ws.ChartObjects()) {
                try { & $chartScan $co.Chart $sheetId $sheetName } catch {}
            }
        }
        try {
            foreach ($chSheet in $wb.Charts) {
                try { & $chartScan $chSheet $null ([string]$chSheet.Name) } catch {}
            }
        } catch {}

        # ── Slicers & Timelines ──
        try {
            foreach ($sc in $wb.SlicerCaches) {
                $scName = try { [string]$sc.Name } catch { 'Slicer' }
                $scType = try { [int]$sc.SlicerCacheType } catch { 1 }
                $scTypeName = if ($script:XL_SLICER_CACHE_TYPE.ContainsKey($scType)) { $script:XL_SLICER_CACHE_TYPE[$scType] } else { 'Slicer' }
                $srcField = try { [string]$sc.SourceName } catch { '' }
                $scId = Get-GraphObjectId -Group 'slicer' -Name $scName
                Add-GraphNode -GraphState $gs -Id $scId -Label $scName -Group 'slicer' -Layer 'structure' -Meta @{ type = $scTypeName; field = $srcField } | Out-Null
                try {
                    foreach ($linkedPt in $sc.PivotTables()) {
                        $ptId = Get-GraphObjectId -Group 'pivot' -Name ([string]$linkedPt.Name)
                        if ($gs.NodeIndex.ContainsKey($ptId)) {
                            Add-GraphEdge -GraphState $gs -From $scId -To $ptId -Label 'filters' -Kind 'slicer-filters' -Arrows 'to' -Layer 'structure' | Out-Null
                        }
                    }
                } catch {}
                $lo = $null; try { $lo = $sc.ListObject } catch {}
                if ($null -ne $lo) {
                    $tId = Get-GraphObjectId -Group 'table' -Name ([string]$lo.Name)
                    if ($gs.NodeIndex.ContainsKey($tId)) {
                        Add-GraphEdge -GraphState $gs -From $scId -To $tId -Label 'filters' -Kind 'slicer-filters' -Arrows 'to' -Layer 'structure' | Out-Null
                    }
                }
            }
        } catch {}

        # ── VBA modules ──
        if (-not $DisableVbaHeuristics) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Scanning VBA modules...' -PercentComplete 78 }
            $vbproj = $null
            try { $vbproj = $wb.VBProject } catch {}
            if ($null -eq $vbproj) {
                Add-GraphWarning -GraphState $gs -Code 'VbaTrustRequired' -Message "VBA project not accessible. Enable 'Trust access to the VBA project object model' in Excel Trust Center to include module analysis." -Meta @{ group = 'module' }
            }
            else {
                try {
                    foreach ($comp in $vbproj.VBComponents) {
                        $compName = [string]$comp.Name
                        $code = ''
                        try {
                            $cm = $comp.CodeModule
                            $lc = [int]$cm.CountOfLines
                            if ($lc -gt 0) { $code = [string]$cm.Lines(1, $lc) }
                        } catch {}
                        # Skip modules with no meaningful code unless requested. Workbooks with
                        # "Require Variable Declaration" auto-insert "Option Explicit" into every
                        # document module (ThisWorkbook / SheetN); treat Option/comment/blank-only as empty.
                        if (-not $IncludeEmptyModules) {
                            $meaningful = ($code -split "`r?`n" | Where-Object {
                                    $t = $_.Trim()
                                    $t -ne '' -and $t -notmatch '^(?i)Option\s' -and $t -notmatch "^'" -and $t -notmatch '^(?i)Rem(\s|$)'
                                })
                            if (-not $meaningful) { continue }
                        }
                        $modId = Get-GraphObjectId -Group 'module' -Name $compName
                        Add-GraphNode -GraphState $gs -Id $modId -Label $compName -Group 'module' -Layer 'structure' -Meta @{
                            lines = ($code -split "`n").Count
                        } | Out-Null
                        Register-GraphNameTarget -GraphState $gs -Name $compName -NodeId $modId -Group 'module'
                        if (-not [string]::IsNullOrWhiteSpace($code)) { $gs.ModuleCodeCache[$compName] = $code }
                    }
                } catch {
                    Add-GraphWarning -GraphState $gs -Code 'VbaScanFailed' -Message ("VBA module scan failed: {0}" -f $_.Exception.Message) -Meta @{ group = 'module' }
                }
            }
        }

        # ── Formula edges (structure) ──
        if ($FormulaMode -ne 'None') {
            if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Analyzing formulas...' -PercentComplete 84 }
            $namesRegex = $null
            if ($definedNames.Count -gt 0) {
                $alt = ($definedNames | Sort-Object { $_.Length } -Descending | ForEach-Object { [regex]::Escape($_) }) -join '|'
                $namesRegex = [regex]::new("(?<![\w.])(?:$alt)(?![\w.])", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            }
            $cellNodeCount = 0
            foreach ($ws in $wb.Worksheets) {
                $sheetName = [string]$ws.Name
                $sheetId = Get-GraphObjectId -Group 'sheet' -Name $sheetName
                $fCells = $null
                try { $fCells = $ws.UsedRange.SpecialCells(-4123) } catch {}   # xlCellTypeFormulas = -4123
                if ($null -eq $fCells) { continue }

                $aggTargets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($area in $fCells.Areas) {
                    foreach ($cell in $area.Cells) {
                        $f = try { [string]$cell.Formula } catch { '' }
                        if ([string]::IsNullOrWhiteSpace($f)) { continue }
                        $targets = Resolve-ExcelFormulaTarget -Formula $f -RefLookup $refLookup -NamesRegex $namesRegex -GraphState $gs -SelfId $sheetId
                        if ($targets.Count -eq 0) { continue }
                        foreach ($t in $targets) { [void]$aggTargets.Add($t) }

                        if (($FormulaMode -eq 'Cell' -or $FormulaMode -eq 'Both') -and $cellNodeCount -lt $MaxFormulaCells) {
                            $addr = try { [string]$cell.Address($false, $false) } catch { '' }
                            $cellId = Get-GraphObjectId -Group 'cell' -Name ($sheetName + '!' + $addr)
                            Add-GraphNode -GraphState $gs -Id $cellId -Label ($sheetName + '!' + $addr) -Group 'cell' -Layer 'structure' -Meta @{
                                sheet = $sheetName; formula = (Get-GraphPreviewText -Text $f -MaxLength 120)
                            } | Out-Null
                            Add-GraphEdge -GraphState $gs -From $sheetId -To $cellId -Label 'contains' -Kind 'contains' -Arrows 'to' -Layer 'structure' | Out-Null
                            foreach ($t in $targets) {
                                Add-GraphEdge -GraphState $gs -From $cellId -To $t -Label 'ref' -Kind 'formula-ref' -Arrows 'to' -Layer 'structure' | Out-Null
                            }
                            $cellNodeCount++
                        }
                    }
                }

                if ($FormulaMode -eq 'Aggregate' -or $FormulaMode -eq 'Both') {
                    foreach ($t in $aggTargets) {
                        Add-GraphEdge -GraphState $gs -From $sheetId -To $t -Label 'ref' -Kind 'formula-ref' -Arrows 'to' -Layer 'structure' | Out-Null
                    }
                }
            }
            if ($cellNodeCount -ge $MaxFormulaCells) {
                Add-GraphWarning -GraphState $gs -Code 'FormulaCellCap' -Message ("Per-cell formula nodes capped at {0}. Use -MaxFormulaCells to raise." -f $MaxFormulaCells) -Meta @{ group = 'cell' }
            }
        }

        # ── Named-range refers-to edges ──
        foreach ($nm in $definedNames) {
            $nameId = Get-GraphObjectId -Group 'name' -Name $nm
            $node = $gs.NodeIndex[$nameId]
            $refersTo = try { [string]$node.meta.refersTo } catch { '' }
            foreach ($t in (Resolve-ExcelFormulaTarget -Formula $refersTo -RefLookup $refLookup -NamesRegex $null -GraphState $gs -SelfId $nameId)) {
                Add-GraphEdge -GraphState $gs -From $nameId -To $t -Label 'refers to' -Kind 'name-refersto' -Arrows 'to' -Layer 'structure' | Out-Null
            }
        }

        # ── VBA code reference edges ──
        if (-not $DisableVbaHeuristics -and $gs.ModuleCodeCache.Count -gt 0) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Analyzing VBA code...' -PercentComplete 90 }
            Build-GraphProcIndex -GraphState $gs
            $sheetRe = [regex]::new('(?:Worksheets|Sheets)\(\s*"([^"]+)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $rangeRe = [regex]::new('Range\(\s*"([^"]+)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($modName in $gs.ModuleCodeCache.Keys) {
                $modId = Get-GraphObjectId -Group 'module' -Name $modName
                $code = $gs.ModuleCodeCache[$modName]

                foreach ($m in $sheetRe.Matches($code)) {
                    $sid = Get-GraphObjectId -Group 'sheet' -Name $m.Groups[1].Value
                    if ($gs.NodeIndex.ContainsKey($sid)) {
                        Add-GraphEdge -GraphState $gs -From $modId -To $sid -Label 'uses' -Kind 'code-ref' -Arrows 'to' -Layer 'structure' | Out-Null
                    }
                }
                foreach ($m in $rangeRe.Matches($code)) {
                    $k = $m.Groups[1].Value.ToLowerInvariant()
                    if ($refLookup.ContainsKey($k)) {
                        Add-GraphEdge -GraphState $gs -From $modId -To $refLookup[$k] -Label 'uses' -Kind 'code-ref' -Arrows 'to' -Layer 'structure' | Out-Null
                    }
                }
                # module -> module public proc calls
                if ($null -ne $gs.ProcCallRe) {
                    foreach ($m in $gs.ProcCallRe.Matches($code)) {
                        $callName = ($m.Value -replace '[\s(]', '' -replace '(?i)^call', '').ToLowerInvariant()
                        if ($gs.ProcIndex.ContainsKey($callName)) {
                            foreach ($targetModId in $gs.ProcIndex[$callName]) {
                                if ($targetModId -ne $modId) {
                                    Add-GraphEdge -GraphState $gs -From $modId -To $targetModId -Label 'calls' -Kind 'code-ref' -Arrows 'to' -Layer 'structure' | Out-Null
                                }
                            }
                        }
                    }
                }
            }
        }

        # ── Value-overlap FK inference (data) ──
        if ($doData -and $columnStatsByTable.Count -gt 1) {
            if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Inferring data relationships...' -PercentComplete 94 }
            $tableNames = @($columnStatsByTable.Keys)
            foreach ($childTable in $tableNames) {
                foreach ($childCol in $columnStatsByTable[$childTable]) {
                    if ($childCol.distinct.Count -eq 0) { continue }
                    foreach ($parentTable in $tableNames) {
                        if ($parentTable -eq $childTable) { continue }
                        foreach ($parentCol in $columnStatsByTable[$parentTable]) {
                            if (-not $parentCol.isPk) { continue }
                            if ($parentCol.distinct.Count -eq 0) { continue }
                            $inter = 0
                            foreach ($v in $childCol.distinct) { if ($parentCol.distinct.Contains($v)) { $inter++ } }
                            if ($inter -eq 0) { continue }
                            $ratio = [double]$inter / [double]$childCol.distinct.Count
                            if ($ratio -lt $FkOverlapThreshold) { continue }
                            # Don't duplicate an explicit datamodel-fk between the same columns
                            $nameMatch = ($childCol.name -eq $parentCol.name) -or ($childCol.name -like "*$($parentCol.name)")
                            $childColId = Get-GraphColumnId -OwnerName $childTable -ColumnName $childCol.name
                            $parentColId = Get-GraphColumnId -OwnerName $parentTable -ColumnName $parentCol.name
                            Add-GraphEdge -GraphState $gs -From $childColId -To $parentColId -Label 'inferred FK' -Kind 'inferred-fk' -Arrows 'to' -Layer 'data' -Meta @{
                                confidence = [Math]::Round($ratio, 2)
                                overlap    = ('{0}/{1}' -f $inter, $childCol.distinct.Count)
                                nameMatch  = $nameMatch
                            } | Out-Null
                        }
                    }
                }
            }
        }

        # ──────────────────────────────────────────────────────────────
        #  OUTPUT
        # ──────────────────────────────────────────────────────────────
        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Status 'Writing graph output...' -PercentComplete 97 }

        $graph = [pscustomobject][ordered]@{
            meta = [ordered]@{
                workbook    = $resolved
                generatedAt = [DateTime]::UtcNow.ToString('o')
                formulaMode = $FormulaMode
                dataGraph   = [bool]$doData
                stats       = [ordered]@{
                    nodeCount   = $gs.Nodes.Count
                    edgeCount   = $gs.Edges.Count
                    sheets      = @($gs.Nodes | Where-Object { $_.group -eq 'sheet' }).Count
                    tables      = @($gs.Nodes | Where-Object { $_.group -eq 'table' }).Count
                    names       = @($gs.Nodes | Where-Object { $_.group -eq 'name' }).Count
                    pivots      = @($gs.Nodes | Where-Object { $_.group -eq 'pivot' }).Count
                    charts      = @($gs.Nodes | Where-Object { $_.group -eq 'chart' }).Count
                    connections = @($gs.Nodes | Where-Object { $_.group -eq 'connection' }).Count
                    queries     = @($gs.Nodes | Where-Object { $_.group -eq 'query' }).Count
                    modelTables = @($gs.Nodes | Where-Object { $_.group -eq 'modeltable' }).Count
                    measures    = @($gs.Nodes | Where-Object { $_.group -eq 'measure' }).Count
                    slicers     = @($gs.Nodes | Where-Object { $_.group -eq 'slicer' }).Count
                    modules     = @($gs.Nodes | Where-Object { $_.group -eq 'module' }).Count
                    columns     = @($gs.Nodes | Where-Object { $_.group -eq 'column' }).Count
                    dataModelFks = @($gs.Edges | Where-Object { $_.kind -eq 'datamodel-fk' }).Count
                    lookupFks    = @($gs.Edges | Where-Object { $_.kind -eq 'lookup-fk' }).Count
                    inferredFks  = @($gs.Edges | Where-Object { $_.kind -eq 'inferred-fk' }).Count
                    warnings     = $gs.Warnings.Count
                }
                warnings = $gs.Warnings.ToArray()
            }
            nodes = $gs.Nodes.ToArray()
            edges = $gs.Edges.ToArray()
        }

        $graphPath = Join-Path $OutDir 'graph.json'
        $graphJson = $graph | ConvertTo-Json -Depth 25
        Set-Content -LiteralPath $graphPath -Value $graphJson -Encoding UTF8

        Copy-GraphViewer -DestinationFolder $OutDir -GraphJson $graphJson -Disabled:$SkipViewerCopy

        if (-not $Quiet) { Write-Progress -Activity 'Export-ExcelGraph' -Completed }
        Write-Host ('Graph written to: ' + $graphPath)
        Write-Host ('Nodes: {0}  Edges: {1}  Warnings: {2}' -f $gs.Nodes.Count, $gs.Edges.Count, $gs.Warnings.Count)

        if ($PassThru) { return $graph }
    }
    finally {
        if ($wb) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {} }
    }
}
