# Public/GraphQueryOps.ps1 — Query a previously-generated Excel workbook graph

function Import-ExcelGraph {
    <#
    .SYNOPSIS
        Load a graph.json produced by Export-ExcelGraph into an in-memory object.
    .DESCRIPTION
        Returns the parsed graph (meta, nodes, edges). Locate the file via -GraphPath, or via
        -WorkbookPath (auto-locates .\excel-graph-out\graph.json next to the workbook).
    .PARAMETER GraphPath
        Path to graph.json.
    .PARAMETER WorkbookPath
        Path to the workbook; used to auto-locate graph.json if -GraphPath is omitted.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        $g = Import-ExcelGraph -GraphPath .\excel-graph-out\graph.json
    #>
    [CmdletBinding()]
    param(
        [string]$GraphPath,
        [string]$WorkbookPath,
        [switch]$AsJson
    )

    $resolvedPath = $GraphPath
    if (-not $resolvedPath -and $WorkbookPath) {
        $wbDir = Split-Path ([System.IO.Path]::GetFullPath($WorkbookPath)) -Parent
        $candidate = Join-Path $wbDir 'excel-graph-out' 'graph.json'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $resolvedPath = $candidate }
    }
    if (-not $resolvedPath -or -not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        $searched = if ($resolvedPath) { $resolvedPath } else { '(no path provided)' }
        throw [System.IO.FileNotFoundException]::new(
            "graph.json not found. Run Export-ExcelGraph first. Searched: $searched", $searched)
    }

    $data = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Format-ExcelOutput -Data $data -AsJson:$AsJson
}

function Get-ExcelGraphQuery {
    <#
    .SYNOPSIS
        Query a previously-generated Excel workbook graph without re-scanning.
    .DESCRIPTION
        Loads graph.json (from Export-ExcelGraph) and supports:
          neighbors — direct connections to/from a node (depth 1-3)
          impact    — transitive downstream dependents
          path      — shortest path between two nodes
          orphans   — nodes with zero incoming edges
          summary   — high-level stats, top edge kinds, high-degree nodes
    .PARAMETER Action
        Query action to perform.
    .PARAMETER GraphPath
        Path to graph.json. If omitted, auto-located next to WorkbookPath.
    .PARAMETER WorkbookPath
        Path to the workbook; used to locate graph.json if GraphPath not provided.
    .PARAMETER Node
        Node name or id for neighbors/impact (e.g. 'Orders' or 'table:Orders').
    .PARAMETER Source
        Source node for the path action.
    .PARAMETER Target
        Target node for the path action.
    .PARAMETER Depth
        BFS depth for neighbors (1-3). Default: 1.
    .PARAMETER Direction
        Edge direction for neighbors: in, out, or both. Default: both.
    .PARAMETER Group
        Filter by node group for summary.
    .PARAMETER IncludeColumns
        Include column nodes and column-owner edges in results. Excluded by default.
    .PARAMETER AsJson
        Emit output as a JSON string.
    .EXAMPLE
        Get-ExcelGraphQuery -Action summary -WorkbookPath C:\Sales.xlsx
    .EXAMPLE
        Get-ExcelGraphQuery -Action neighbors -GraphPath .\excel-graph-out\graph.json -Node Orders -Depth 2
    .EXAMPLE
        Get-ExcelGraphQuery -Action impact -WorkbookPath C:\Sales.xlsx -Node table:Customers -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('neighbors', 'impact', 'path', 'orphans', 'summary')]
        [string]$Action,

        [string]$GraphPath,
        [string]$WorkbookPath,
        [string]$Node,
        [string]$Source,
        [string]$Target,

        [ValidateRange(1, 3)]
        [int]$Depth = 1,

        [ValidateSet('in', 'out', 'both')]
        [string]$Direction = 'both',

        [string]$Group,
        [switch]$IncludeColumns,
        [switch]$AsJson
    )

    $MaxResults = 200
    $skip = -not $IncludeColumns
    $graph = Resolve-GraphInput -GraphPath $GraphPath -WorkbookPath $WorkbookPath

    function _IsSkipped($node, $edge) {
        if (-not $skip) { return $false }
        if ($null -ne $node -and $node.group -ieq 'column') { return $true }
        if ($null -ne $edge -and $edge.kind -ieq 'column-owner') { return $true }
        return $false
    }

    function _MustResolve([string]$Name, [string]$ParamName) {
        if ([string]::IsNullOrWhiteSpace($Name)) { throw "'$ParamName' is required for action '$Action'." }
        $hits = Resolve-GraphNode -Graph $graph -Name $Name
        if ($hits.Count -eq 0) { throw "Node '$Name' not found. Try the full id (e.g. 'table:Orders')." }
        if ($hits.Count -gt 1) {
            $opts = ($hits | Select-Object -First 10 | ForEach-Object { '{0} ({1})' -f $_.id, $_.group }) -join ', '
            throw "'$Name' is ambiguous - matches: $opts. Use the full id."
        }
        return [string]$hits[0].id
    }

    function _FmtNode($N) {
        $out = [ordered]@{ id = $N.id; label = $N.label; group = $N.group }
        if ($null -ne $N.layer) { $out['layer'] = $N.layer }
        return [PSCustomObject]$out
    }
    function _FmtEdge($E) {
        return [PSCustomObject][ordered]@{ from = $E.from; to = $E.to; kind = $E.kind; label = $E.label }
    }

    switch ($Action) {
        'neighbors' {
            $nodeId = _MustResolve $Node 'Node'
            $resultIn = New-Object 'System.Collections.Generic.List[object]'
            $resultOut = New-Object 'System.Collections.Generic.List[object]'

            if ($Direction -ieq 'out' -or $Direction -ieq 'both') {
                $visited = New-Object 'System.Collections.Generic.HashSet[string]'; [void]$visited.Add($nodeId)
                $frontier = New-Object 'System.Collections.Generic.Queue[object]'
                $frontier.Enqueue([PSCustomObject]@{ Id = $nodeId; D = 0 })
                while ($frontier.Count -gt 0) {
                    $item = $frontier.Dequeue(); $cur = [string]$item.Id; $d = [int]$item.D
                    if ($d -ge $Depth) { continue }
                    if ($graph.OutAdj.ContainsKey($cur)) {
                        foreach ($e in $graph.OutAdj[$cur]) {
                            if (_IsSkipped $null $e) { continue }
                            $tgt = [string]$e.to
                            $fe = _FmtEdge $e; $fe | Add-Member -NotePropertyName depth -NotePropertyValue ($d + 1)
                            $resultOut.Add($fe)
                            if (-not $visited.Contains($tgt)) { [void]$visited.Add($tgt); $frontier.Enqueue([PSCustomObject]@{ Id = $tgt; D = ($d + 1) }) }
                        }
                    }
                }
            }
            if ($Direction -ieq 'in' -or $Direction -ieq 'both') {
                $visitedIn = New-Object 'System.Collections.Generic.HashSet[string]'; [void]$visitedIn.Add($nodeId)
                $frontier = New-Object 'System.Collections.Generic.Queue[object]'
                $frontier.Enqueue([PSCustomObject]@{ Id = $nodeId; D = 0 })
                while ($frontier.Count -gt 0) {
                    $item = $frontier.Dequeue(); $cur = [string]$item.Id; $d = [int]$item.D
                    if ($d -ge $Depth) { continue }
                    if ($graph.InAdj.ContainsKey($cur)) {
                        foreach ($e in $graph.InAdj[$cur]) {
                            if (_IsSkipped $null $e) { continue }
                            $src = [string]$e.from
                            $fe = _FmtEdge $e; $fe | Add-Member -NotePropertyName depth -NotePropertyValue ($d + 1)
                            $resultIn.Add($fe)
                            if (-not $visitedIn.Contains($src)) { [void]$visitedIn.Add($src); $frontier.Enqueue([PSCustomObject]@{ Id = $src; D = ($d + 1) }) }
                        }
                    }
                }
            }

            $out = [PSCustomObject][ordered]@{
                action         = 'neighbors'
                node           = (_FmtNode $graph.Nodes[$nodeId])
                depth          = $Depth
                direction      = $Direction
                incoming       = @($resultIn | Select-Object -First $MaxResults)
                outgoing       = @($resultOut | Select-Object -First $MaxResults)
                total_incoming = $resultIn.Count
                total_outgoing = $resultOut.Count
                truncated      = (($resultIn.Count -gt $MaxResults) -or ($resultOut.Count -gt $MaxResults))
            }
            return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
        }

        'impact' {
            $nodeId = _MustResolve $Node 'Node'
            $visited = New-Object 'System.Collections.Generic.HashSet[string]'; [void]$visited.Add($nodeId)
            $frontier = New-Object 'System.Collections.Generic.Queue[string]'; $frontier.Enqueue($nodeId)
            $affected = New-Object 'System.Collections.Generic.List[object]'
            while ($frontier.Count -gt 0) {
                $cur = $frontier.Dequeue()
                if ($graph.OutAdj.ContainsKey($cur)) {
                    foreach ($e in $graph.OutAdj[$cur]) {
                        if (_IsSkipped $null $e) { continue }
                        $tgt = [string]$e.to
                        if (-not $visited.Contains($tgt)) { [void]$visited.Add($tgt); $affected.Add((_FmtNode $graph.Nodes[$tgt])); $frontier.Enqueue($tgt) }
                    }
                }
                if ($affected.Count -ge $MaxResults) { break }
            }
            $byGroup = @{}
            foreach ($a in $affected) { $g = [string]$a.group; if (-not $byGroup.ContainsKey($g)) { $byGroup[$g] = New-Object 'System.Collections.Generic.List[string]' }; $byGroup[$g].Add([string]$a.label) }
            $out = [PSCustomObject][ordered]@{
                action            = 'impact'
                node              = (_FmtNode $graph.Nodes[$nodeId])
                affected_count    = $affected.Count
                affected_by_group = $byGroup
                affected          = @($affected | Select-Object -First $MaxResults)
                truncated         = ($affected.Count -ge $MaxResults)
            }
            return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
        }

        'path' {
            $srcId = _MustResolve $Source 'Source'
            $tgtId = _MustResolve $Target 'Target'
            $srcNode = _FmtNode $graph.Nodes[$srcId]
            $tgtNode = _FmtNode $graph.Nodes[$tgtId]
            if ($srcId -ieq $tgtId) {
                $out = [PSCustomObject][ordered]@{ action = 'path'; source = $srcNode; target = $tgtNode; found = $true; path = @($srcNode); edges = @(); length = 0 }
                return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
            }
            $adj = @{}
            foreach ($e in $graph.Edges) {
                if (_IsSkipped $null $e) { continue }
                $f = [string]$e.from; $t = [string]$e.to
                if (-not $adj.ContainsKey($f)) { $adj[$f] = New-Object 'System.Collections.Generic.List[object]' }
                $adj[$f].Add([PSCustomObject]@{ Neighbor = $t; Edge = $e })
                if (-not $adj.ContainsKey($t)) { $adj[$t] = New-Object 'System.Collections.Generic.List[object]' }
                $adj[$t].Add([PSCustomObject]@{ Neighbor = $f; Edge = $e })
            }
            $visited = New-Object 'System.Collections.Generic.HashSet[string]'; [void]$visited.Add($srcId)
            $parent = @{}; $frontier = New-Object 'System.Collections.Generic.Queue[string]'; $frontier.Enqueue($srcId); $found = $false
            while ($frontier.Count -gt 0 -and -not $found) {
                $cur = $frontier.Dequeue()
                if ($adj.ContainsKey($cur)) {
                    foreach ($pair in $adj[$cur]) {
                        $nb = [string]$pair.Neighbor
                        if (-not $visited.Contains($nb)) {
                            [void]$visited.Add($nb); $parent[$nb] = [PSCustomObject]@{ ParentId = $cur; Edge = $pair.Edge }
                            if ($nb -ieq $tgtId) { $found = $true; break }
                            $frontier.Enqueue($nb)
                        }
                    }
                }
            }
            if (-not $found) {
                $out = [PSCustomObject][ordered]@{ action = 'path'; source = $srcNode; target = $tgtNode; found = $false; path = @(); edges = @(); length = -1 }
                return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
            }
            $pathNodes = New-Object 'System.Collections.Generic.List[object]'; $pathEdges = New-Object 'System.Collections.Generic.List[object]'
            $cur = $tgtId
            while ($cur -ine $srcId) { $pathNodes.Add((_FmtNode $graph.Nodes[$cur])); $pInfo = $parent[$cur]; $pathEdges.Add((_FmtEdge $pInfo.Edge)); $cur = [string]$pInfo.ParentId }
            $pathNodes.Add((_FmtNode $graph.Nodes[$srcId])); $pathNodes.Reverse(); $pathEdges.Reverse()
            $out = [PSCustomObject][ordered]@{ action = 'path'; source = $srcNode; target = $tgtNode; found = $true; path = $pathNodes.ToArray(); edges = $pathEdges.ToArray(); length = $pathEdges.Count }
            return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
        }

        'orphans' {
            $orphans = New-Object 'System.Collections.Generic.List[object]'
            foreach ($nid in $graph.Nodes.Keys) {
                $n = $graph.Nodes[$nid]
                if (_IsSkipped $n $null) { continue }
                $hasIncoming = $false
                if ($graph.InAdj.ContainsKey($nid)) {
                    foreach ($e in $graph.InAdj[$nid]) { if (_IsSkipped $null $e) { continue }; $hasIncoming = $true; break }
                }
                if (-not $hasIncoming) { $orphans.Add((_FmtNode $n)) }
            }
            $byGroup = @{}
            foreach ($o in $orphans) { $g = [string]$o.group; if (-not $byGroup.ContainsKey($g)) { $byGroup[$g] = New-Object 'System.Collections.Generic.List[string]' }; $byGroup[$g].Add([string]$o.label) }
            $out = [PSCustomObject][ordered]@{
                action = 'orphans'; count = $orphans.Count; by_group = $byGroup
                orphans = @($orphans | Select-Object -First $MaxResults); truncated = ($orphans.Count -gt $MaxResults)
            }
            return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
        }

        'summary' {
            $nodeCounts = @{}
            foreach ($n in $graph.Nodes.Values) {
                if ($Group -and $n.group -ine $Group) { continue }
                $g = [string]$n.group
                if (-not $nodeCounts.ContainsKey($g)) { $nodeCounts[$g] = 0 }
                $nodeCounts[$g]++
            }
            $edgeKinds = @{}
            foreach ($e in $graph.Edges) {
                $k = [string]$e.kind
                if (-not $edgeKinds.ContainsKey($k)) { $edgeKinds[$k] = 0 }
                $edgeKinds[$k]++
            }
            # High fan-in (most incoming edges)
            $degree = @{}
            foreach ($nid in $graph.InAdj.Keys) { $degree[$nid] = $graph.InAdj[$nid].Count }
            $topFanIn = $degree.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
                $n = $graph.Nodes[$_.Key]
                [PSCustomObject][ordered]@{ id = $_.Key; label = $n.label; group = $n.group; inbound = $_.Value }
            }
            $out = [PSCustomObject][ordered]@{
                action        = 'summary'
                workbook      = $graph.Meta['workbook']
                nodeCount     = $graph.Nodes.Count
                edgeCount     = $graph.Edges.Count
                nodesByGroup  = $nodeCounts
                edgesByKind   = $edgeKinds
                topFanIn      = @($topFanIn)
            }
            return (Format-ExcelOutput -Data $out -AsJson:$AsJson)
        }
    }
}
