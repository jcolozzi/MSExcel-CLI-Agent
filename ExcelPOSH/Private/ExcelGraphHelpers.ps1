# Private/ExcelGraphHelpers.ps1 — Graph analysis helpers for Export-ExcelGraph
#
# Ported from AccessPOSH/Private/GraphHelpers.ps1 and adapted for the Excel object model.
# All functions operate on a per-invocation $GraphState hashtable (no $script: pollution),
# except the read-only VBA-builtin set below. These functions are module-private (dot-sourced
# by ExcelPOSH.psm1) so their generic names never collide with AccessPOSH when both are loaded.

# Built-in VBA / Excel function names — excluded from cross-module call detection.
$script:XL_VBA_BUILTIN_NAMES = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'asc','ascw','chr','chrw','format','instr','instrb','instrrev','join',
        'lcase','left','len','lenb','ltrim','mid','replace','right','space',
        'split','str','strcomp','strconv','strreverse','trim','rtrim','ucase','val','string',
        'cbool','cbyte','ccur','cdate','cdbl','cdec','cint','clng','clnglng','clngptr','csng','cstr','cvar','cverr',
        'isarray','isdate','isempty','iserror','ismissing','isnull','isnumeric','isobject','typename','vartype',
        'abs','atn','cos','exp','fix','int','log','rnd','round','sgn','sin','sqr','tan',
        'date','dateadd','datediff','datepart','dateserial','datevalue','day','formatdatetime',
        'hour','minute','month','monthname','now','second','time','timeserial','timevalue','timer',
        'weekday','weekdayname','year',
        'inputbox','msgbox',
        'curdir','dir','eof','filecopy','filedatetime','filelen','freefile','getattr','loc','lof','setattr',
        'array','erase','filter','lbound','ubound',
        'appactivate','beep','command','doevents','environ','sendkeys','shell',
        'error','iif','choose','switch','rgb','qbcolor',
        'callbyname','createobject','getobject','hex','oct',
        'range','cells','rows','columns','sheets','worksheets','workbooks','activesheet','activecell',
        'application','thisworkbook','activeworkbook','selection','offset','union','intersect','evaluate','worksheetfunction'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ──────────────────────────────────────────────────────────────────────
#  Graph State Factory
# ──────────────────────────────────────────────────────────────────────

function New-GraphState {
    <#
    .SYNOPSIS
        Create a fresh per-invocation graph state hashtable.
    #>
    return @{
        Nodes           = New-Object 'System.Collections.Generic.List[object]'
        Edges           = New-Object 'System.Collections.Generic.List[object]'
        NodeIndex       = @{}
        EdgeIndex       = @{}
        NameTargets     = @{}
        DataNameTargets = @{}
        KnownNames      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        Warnings        = New-Object 'System.Collections.Generic.List[object]'
        EdgeId          = 0
        ProcIndex       = @{}
        ProcCallRe      = $null
        ModuleCodeCache = @{}
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Utility Helpers
# ──────────────────────────────────────────────────────────────────────

function Get-GraphObjectId {
    param([string]$Group, [string]$Name)
    return ('{0}:{1}' -f $Group, $Name)
}

function Get-GraphColumnId {
    param([string]$OwnerName, [string]$ColumnName)
    return ('column:{0}.{1}' -f $OwnerName, $ColumnName)
}

function Remove-GraphBrackets {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $Name }
    $trimmed = $Name.Trim()
    # Strip a single wrapping pair of [] or '' (Excel sheet/name quoting)
    if ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']') -and $trimmed.Length -ge 2) {
        return $trimmed.Substring(1, $trimmed.Length - 2)
    }
    if ($trimmed.StartsWith("'") -and $trimmed.EndsWith("'") -and $trimmed.Length -ge 2) {
        return $trimmed.Substring(1, $trimmed.Length - 2)
    }
    return $trimmed
}

function Get-GraphTextHash {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-GraphPreviewText {
    param([string]$Text, [int]$MaxLength = 180)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $flat = ($Text -replace '\s+', ' ').Trim()
    if ($flat.Length -le $MaxLength) { return $flat }
    return $flat.Substring(0, $MaxLength).TrimEnd() + '...'
}

function Format-GraphMetaTitle {
    param([System.Collections.IDictionary]$Meta)
    if (-not $Meta -or $Meta.Count -eq 0) { return '' }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($key in ($Meta.Keys | Sort-Object)) {
        $value = $Meta[$key]
        if ($null -eq $value) { continue }
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            $joined = ($value | ForEach-Object { $_ }) -join ', '
            $parts.Add(("{0}: {1}" -f $key, $joined))
        }
        else {
            $parts.Add(("{0}: {1}" -f $key, $value))
        }
    }
    return ($parts -join "`n")
}

# ──────────────────────────────────────────────────────────────────────
#  Warning Helper
# ──────────────────────────────────────────────────────────────────────

function Add-GraphWarning {
    param(
        [hashtable]$GraphState,
        [string]$Code,
        [string]$Message,
        [hashtable]$Meta = @{}
    )
    $entry = [pscustomobject][ordered]@{
        code    = $Code
        message = $Message
        meta    = $Meta
    }
    $GraphState.Warnings.Add($entry)
    Write-Warning $Message
}

# ──────────────────────────────────────────────────────────────────────
#  Node & Edge Management
# ──────────────────────────────────────────────────────────────────────

function Add-GraphNode {
    param(
        [hashtable]$GraphState,
        [string]$Id,
        [string]$Label,
        [string]$Group,
        [ValidateSet('structure', 'data', 'both')]
        [string]$Layer = 'structure',
        [hashtable]$Meta = @{}
    )

    if ($GraphState.NodeIndex.ContainsKey($Id)) {
        $existing = $GraphState.NodeIndex[$Id]
        foreach ($key in $Meta.Keys) { $existing.meta[$key] = $Meta[$key] }
        if ($Label) { $existing.label = $Label }
        if ($Group) { $existing.group = $Group }
        # Promote layer to 'both' when a node is referenced by both layers
        if ($existing.layer -ne $Layer) { $existing.layer = 'both' }
        $existing.title = Format-GraphMetaTitle -Meta $existing.meta
        return $existing
    }

    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) { $metaCopy[$key] = $Meta[$key] }

    $node = [pscustomobject][ordered]@{
        id    = $Id
        label = $Label
        group = $Group
        layer = $Layer
        title = Format-GraphMetaTitle -Meta $metaCopy
        meta  = $metaCopy
    }

    $GraphState.Nodes.Add($node)
    $GraphState.NodeIndex[$Id] = $node
    return $node
}

function Add-GraphEdge {
    param(
        [hashtable]$GraphState,
        [string]$From,
        [string]$To,
        [string]$Label,
        [string]$Kind,
        [string]$Arrows = 'to',
        [ValidateSet('structure', 'data')]
        [string]$Layer = 'structure',
        [hashtable]$Meta = @{}
    )

    $metaJson = ConvertTo-Json -InputObject $Meta -Depth 8 -Compress
    $edgeKey = ($From, $To, $Kind, $Label, $Arrows, $metaJson) -join '|'
    if ($GraphState.EdgeIndex.ContainsKey($edgeKey)) { return }

    $GraphState.EdgeId += 1
    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) { $metaCopy[$key] = $Meta[$key] }

    $edge = [pscustomobject][ordered]@{
        id     = ('e{0}' -f $GraphState.EdgeId)
        from   = $From
        to     = $To
        label  = $Label
        kind   = $Kind
        arrows = $Arrows
        layer  = $Layer
        title  = Format-GraphMetaTitle -Meta $metaCopy
        meta   = $metaCopy
    }

    $GraphState.Edges.Add($edge)
    $GraphState.EdgeIndex[$edgeKey] = $edge
}

function Register-GraphNameTarget {
    param(
        [hashtable]$GraphState,
        [string]$Name,
        [string]$NodeId,
        [string]$Group,
        [switch]$IsDataObject
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    [void]$GraphState.KnownNames.Add($Name)

    if (-not $GraphState.NameTargets.ContainsKey($Name)) {
        $GraphState.NameTargets[$Name] = New-Object 'System.Collections.Generic.List[object]'
    }
    $GraphState.NameTargets[$Name].Add([pscustomobject]@{ id = $NodeId; group = $Group; name = $Name })

    if ($IsDataObject) {
        $lcName = $Name.ToLowerInvariant()
        if (-not $GraphState.DataNameTargets.ContainsKey($lcName)) {
            $GraphState.DataNameTargets[$lcName] = New-Object 'System.Collections.Generic.List[object]'
        }
        $GraphState.DataNameTargets[$lcName].Add([pscustomobject]@{ id = $NodeId; group = $Group; name = $Name })
    }
}

function Get-GraphTargetsByName {
    param([string]$Name, [hashtable]$TargetTable)

    if ([string]::IsNullOrWhiteSpace($Name)) { return @() }

    $lcName = $Name.ToLowerInvariant()
    if ($TargetTable.ContainsKey($lcName)) { return $TargetTable[$lcName].ToArray() }

    $unbracketed = (Remove-GraphBrackets -Name $Name).ToLowerInvariant()
    if ($TargetTable.ContainsKey($unbracketed)) { return $TargetTable[$unbracketed].ToArray() }

    return @()
}

# ──────────────────────────────────────────────────────────────────────
#  Excel-specific helpers
# ──────────────────────────────────────────────────────────────────────

function Get-ExcelGraphTargetWorkbook {
    <#
    .SYNOPSIS
        Resolve the target workbook by FullName (Connect-ExcelWorkbook does not always
        re-activate an already-open workbook). Falls back to ActiveWorkbook.
    #>
    param($App, [string]$ResolvedPath)
    try {
        foreach ($wb in $App.Workbooks) {
            if ($wb.FullName -eq $ResolvedPath) { return $wb }
        }
    } catch {}
    return $App.ActiveWorkbook
}

function Get-ExcelFormulaReference {
    <#
    .SYNOPSIS
        Extract structural references (tables, sheets, external workbooks) from a formula string.
    .OUTPUTS
        Array of [pscustomobject]@{ kind = 'table'|'sheet'|'external'; name = <string> }
    #>
    param([string]$Formula)

    $refs = New-Object 'System.Collections.Generic.List[object]'
    if ([string]::IsNullOrWhiteSpace($Formula)) { return $refs }

    # External workbook links: [Book.xlsx] or 'C:\path\[Book.xlsx]Sheet'!
    foreach ($m in [regex]::Matches($Formula, '\[([^\[\]]+\.xls[a-z]*)\]')) {
        $refs.Add([pscustomobject]@{ kind = 'external'; name = $m.Groups[1].Value })
    }

    # Structured table references: TableName[...]  (name token immediately before '[')
    foreach ($m in [regex]::Matches($Formula, '(?<![\w.\]])([A-Za-z_\\][\w.]*)\[')) {
        $refs.Add([pscustomobject]@{ kind = 'table'; name = $m.Groups[1].Value })
    }

    # Quoted sheet references: 'Sheet Name'!
    foreach ($m in [regex]::Matches($Formula, "'([^']+)'!")) {
        $val = $m.Groups[1].Value
        if ($val -notmatch '\.xls') { $refs.Add([pscustomobject]@{ kind = 'sheet'; name = $val }) }
    }

    # Bare sheet references: SheetName!  (not preceded by ] or ' which indicates external/quoted)
    foreach ($m in [regex]::Matches($Formula, "(?<![\w'\]!])([A-Za-z_][\w.]*)!")) {
        $refs.Add([pscustomobject]@{ kind = 'sheet'; name = $m.Groups[1].Value })
    }

    return $refs
}

$script:XL_LOOKUP_FN_RE = [regex]::new('\b(VLOOKUP|HLOOKUP|XLOOKUP|LOOKUP|INDEX|MATCH|XMATCH)\s*\(',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

function Test-ExcelLookupFormula {
    <#
    .SYNOPSIS
        True if the formula contains a lookup-family function.
    #>
    param([string]$Formula)
    if ([string]::IsNullOrWhiteSpace($Formula)) { return $false }
    return $script:XL_LOOKUP_FN_RE.IsMatch($Formula)
}

# ──────────────────────────────────────────────────────────────────────
#  Cross-Module VBA Procedure Call Index
# ──────────────────────────────────────────────────────────────────────

function Build-GraphProcIndex {
    <#
    .SYNOPSIS
        Index public VBA procedures across modules for cross-module call detection.
    #>
    param([hashtable]$GraphState)

    $procDeclRe    = '(?im)^\s*(?:Public\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'
    $privateProcRe = '(?im)^\s*Private\s+(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'

    foreach ($node in $GraphState.NodeIndex.Values) {
        if ($node.group -ne 'module') { continue }
        if (-not $GraphState.ModuleCodeCache.ContainsKey($node.label)) { continue }
        $code = $GraphState.ModuleCodeCache[$node.label]
        if ([string]::IsNullOrWhiteSpace($code)) { continue }

        $privateNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in [regex]::Matches($code, $privateProcRe)) { [void]$privateNames.Add($m.Groups[1].Value) }

        $nodeId = $node.id
        foreach ($m in [regex]::Matches($code, $procDeclRe)) {
            $procName = $m.Groups[1].Value
            $pnameLower = $procName.ToLowerInvariant()
            if ($privateNames.Contains($procName)) { continue }
            if ($script:XL_VBA_BUILTIN_NAMES.Contains($pnameLower)) { continue }
            if ($pnameLower.Length -lt 2) { continue }
            if (-not $GraphState.ProcIndex.ContainsKey($pnameLower)) {
                $GraphState.ProcIndex[$pnameLower] = New-Object 'System.Collections.Generic.List[string]'
            }
            $GraphState.ProcIndex[$pnameLower].Add($nodeId)
        }
    }

    $procNames = @($GraphState.ProcIndex.Keys | Sort-Object { $_.Length } -Descending)
    if ($procNames.Count -gt 0) {
        $escaped = $procNames | ForEach-Object { [regex]::Escape($_) }
        $alt = $escaped -join '|'
        $GraphState.ProcCallRe = [regex]::new(
            "(?<![\.\w])(?:$alt)\s*\(|\bCall\s+(?:$alt)\b",
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
}

# ──────────────────────────────────────────────────────────────────────
#  Viewer Embedding
# ──────────────────────────────────────────────────────────────────────

function Copy-GraphViewer {
    <#
    .SYNOPSIS
        Embed graph.json into the vis.js HTML viewer and write index.html.
    #>
    param(
        [string]$DestinationFolder,
        [string]$GraphJson,
        [switch]$Disabled
    )

    if ($Disabled) { return }

    # When dot-sourced from Private/, $PSScriptRoot is the Private/ folder
    $moduleRoot = Split-Path $PSScriptRoot -Parent
    $viewerSource = Join-Path $moduleRoot 'Resources' 'excel-graph-viewer.html'

    if (-not (Test-Path -LiteralPath $viewerSource)) {
        Write-Warning "Graph viewer HTML not found at $viewerSource — skipping embedded viewer."
        return
    }

    $html = Get-Content -LiteralPath $viewerSource -Raw
    $embedTag = "<script>var EMBEDDED_GRAPH = $GraphJson;</script>"
    $html = $html -replace '<!-- EMBED_GRAPH_DATA -->', $embedTag
    Set-Content -LiteralPath (Join-Path $DestinationFolder 'index.html') -Value $html -Encoding UTF8
}

# ──────────────────────────────────────────────────────────────────────
#  Graph Query Helpers (ported from AccessPOSH)
# ──────────────────────────────────────────────────────────────────────

function ConvertFrom-GraphJson {
    <#
    .SYNOPSIS
        Load a graph.json file and build adjacency lookup tables for querying.
    .OUTPUTS
        PSCustomObject: Nodes, Edges, OutAdj, InAdj, IdLookup, LabelLookup, Meta
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw [System.IO.FileNotFoundException]::new(
            "graph.json not found at: $Path. Run Export-ExcelGraph first.", $Path)
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $data = ConvertFrom-Json $raw

    $nodes = @{}
    $idLookup = @{}
    $labelLookup = @{}
    $outAdj = @{}
    $inAdj  = @{}

    $nodeList = @()
    if ($null -ne $data.nodes) { $nodeList = @($data.nodes) }
    foreach ($n in $nodeList) {
        $nid = [string]$n.id
        $nodes[$nid] = $n
        $idLookup[$nid.ToLower()] = $n
        $labelKey = ([string]$n.label).ToLower()
        if (-not $labelLookup.ContainsKey($labelKey)) {
            $labelLookup[$labelKey] = New-Object 'System.Collections.Generic.List[object]'
        }
        $labelLookup[$labelKey].Add($n)
    }

    $edgeList = @()
    if ($null -ne $data.edges) { $edgeList = @($data.edges) }
    foreach ($e in $edgeList) {
        $fromId = [string]$e.from
        $toId   = [string]$e.to
        if (-not $outAdj.ContainsKey($fromId)) { $outAdj[$fromId] = New-Object 'System.Collections.Generic.List[object]' }
        $outAdj[$fromId].Add($e)
        if (-not $inAdj.ContainsKey($toId)) { $inAdj[$toId] = New-Object 'System.Collections.Generic.List[object]' }
        $inAdj[$toId].Add($e)
    }

    $meta = @{}
    if ($null -ne $data.meta) {
        $data.meta.PSObject.Properties | ForEach-Object { $meta[$_.Name] = $_.Value }
    }

    return [PSCustomObject]@{
        Nodes = $nodes; Edges = $edgeList; OutAdj = $outAdj; InAdj = $inAdj
        IdLookup = $idLookup; LabelLookup = $labelLookup; Meta = $meta
    }
}

function Resolve-GraphNode {
    <#
    .SYNOPSIS
        Resolve a user-supplied name to graph node(s). Priority: exact id → group:name → label.
    #>
    param(
        [Parameter(Mandatory)][PSCustomObject]$Graph,
        [Parameter(Mandatory)][string]$Name
    )

    $low = $Name.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($low)) { return @() }

    if ($Graph.IdLookup.ContainsKey($low)) { return @($Graph.IdLookup[$low]) }

    $groups = @('table','sheet','name','pivot','chart','connection','query','modeltable','measure','slicer','module','column','external','workbook')
    foreach ($g in $groups) {
        $candidate = ('{0}:{1}' -f $g, $Name).ToLower()
        if ($Graph.IdLookup.ContainsKey($candidate)) { return @($Graph.IdLookup[$candidate]) }
    }

    if ($Graph.LabelLookup.ContainsKey($low)) {
        $hits = $Graph.LabelLookup[$low]
        if ($hits.Count -gt 0) { return @($hits) }
    }

    return @()
}

function Resolve-GraphInput {
    <#
    .SYNOPSIS
        Resolve -Graph / -GraphPath / -WorkbookPath to a loaded graph object.
    #>
    param(
        [PSCustomObject]$Graph,
        [string]$GraphPath,
        [string]$WorkbookPath
    )

    if ($null -ne $Graph -and $null -ne $Graph.Nodes) { return $Graph }

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

    return ConvertFrom-GraphJson -Path $resolvedPath
}
