# Tests/ExcelGraph.Tests.ps1
# Pester 5+ tests — Excel workbook grapher (Export-ExcelGraph, Import-ExcelGraph, Get-ExcelGraphQuery)
# Parameter-validation tests always run. Integration tests (Tag 'Integration') require Excel and
# the sample workbook (misc/ExcelGraph-Sample.xlsm); they are skipped when it is absent.

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop

    function Assert-Mandatory {
        param($Command, $Name)
        $p = (Get-Command $Command).Parameters[$Name]
        $p | Should -Not -BeNullOrEmpty
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }

    $script:SampleWb = (Join-Path $PSScriptRoot '..\misc\ExcelGraph-Sample.xlsm')
    $script:SampleWb = if (Test-Path $script:SampleWb) { (Resolve-Path $script:SampleWb).Path } else { $script:SampleWb }
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ─── Parameter validation ────────────────────────────────────────────────────
Describe 'Export-ExcelGraph' {
    It 'Has CmdletBinding'      { (Get-Command Export-ExcelGraph).CmdletBinding | Should -BeTrue }
    It 'WorkbookPath mandatory' { Assert-Mandatory Export-ExcelGraph WorkbookPath }
    It 'Has OutDir parameter'   { (Get-Command Export-ExcelGraph).Parameters['OutDir'] | Should -Not -BeNullOrEmpty }
    It 'Has FormulaMode ValidateSet' {
        $p = (Get-Command Export-ExcelGraph).Parameters['FormulaMode']
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.ValidValues | Should -Contain 'Aggregate'
        $vs.ValidValues | Should -Contain 'Both'
    }
    It 'Has DisableDataGraph switch'     { (Get-Command Export-ExcelGraph).Parameters['DisableDataGraph'].SwitchParameter | Should -BeTrue }
    It 'Has DisableVbaHeuristics switch' { (Get-Command Export-ExcelGraph).Parameters['DisableVbaHeuristics'].SwitchParameter | Should -BeTrue }
    It 'Has PassThru switch'             { (Get-Command Export-ExcelGraph).Parameters['PassThru'].SwitchParameter | Should -BeTrue }
}

Describe 'Import-ExcelGraph' {
    It 'Has CmdletBinding'          { (Get-Command Import-ExcelGraph).CmdletBinding | Should -BeTrue }
    It 'Has GraphPath parameter'    { (Get-Command Import-ExcelGraph).Parameters['GraphPath'] | Should -Not -BeNullOrEmpty }
    It 'Has WorkbookPath parameter' { (Get-Command Import-ExcelGraph).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty }
    It 'Has AsJson switch'          { (Get-Command Import-ExcelGraph).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Get-ExcelGraphQuery' {
    It 'Has CmdletBinding'  { (Get-Command Get-ExcelGraphQuery).CmdletBinding | Should -BeTrue }
    It 'Action mandatory'   { Assert-Mandatory Get-ExcelGraphQuery Action }
    It 'Has Action ValidateSet with expected actions' {
        $p = (Get-Command Get-ExcelGraphQuery).Parameters['Action']
        $vs = $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        foreach ($a in 'neighbors', 'impact', 'path', 'orphans', 'summary') { $vs.ValidValues | Should -Contain $a }
    }
    It 'Has Depth ValidateRange' {
        $p = (Get-Command Get-ExcelGraphQuery).Parameters['Depth']
        $p.Attributes.Where({ $_ -is [System.Management.Automation.ValidateRangeAttribute] }) | Should -Not -BeNullOrEmpty
    }
    It 'Has IncludeColumns switch' { (Get-Command Get-ExcelGraphQuery).Parameters['IncludeColumns'].SwitchParameter | Should -BeTrue }
}

# ─── Integration (requires Excel + sample workbook) ──────────────────────────
Describe 'ExcelGraph end-to-end' -Tag 'Integration' {
    BeforeAll {
        $script:skipReason = if (-not (Test-Path $script:SampleWb)) { 'Sample workbook not found' } else { $null }
        if (-not $script:skipReason) {
            $script:OutDir = Join-Path ([System.IO.Path]::GetTempPath()) ('xlgraph_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            try {
                $script:Graph = Export-ExcelGraph -WorkbookPath $script:SampleWb -OutDir $script:OutDir -FormulaMode Both -PassThru -Quiet
            }
            catch { $script:skipReason = "Export failed: $($_.Exception.Message)" }
            finally { try { Close-ExcelWorkbook | Out-Null } catch {} }
        }
    }
    AfterAll {
        if ($script:OutDir -and (Test-Path $script:OutDir)) { Remove-Item $script:OutDir -Recurse -Force -ErrorAction SilentlyContinue }
    }

    It 'writes graph.json'  { -not $script:skipReason | Should -BeTrue -Because ($script:skipReason ?? 'ok'); Test-Path (Join-Path $script:OutDir 'graph.json') | Should -BeTrue }
    It 'writes index.html'  { Test-Path (Join-Path $script:OutDir 'index.html') | Should -BeTrue }
    It 'has >= 6 table nodes'           { @($script:Graph.nodes | Where-Object group -eq 'table').Count | Should -BeGreaterOrEqual 6 }
    It 'has >= 5 data-model FK edges'   { @($script:Graph.edges | Where-Object kind -eq 'datamodel-fk').Count | Should -BeGreaterOrEqual 5 }
    It 'has lookup-fk edges'            { @($script:Graph.edges | Where-Object kind -eq 'lookup-fk').Count | Should -BeGreaterThan 0 }
    It 'has inferred-fk edges'          { @($script:Graph.edges | Where-Object kind -eq 'inferred-fk').Count | Should -BeGreaterThan 0 }
    It 'detects primary key columns'    { @($script:Graph.nodes | Where-Object { $_.group -eq 'column' -and $_.meta.isPrimaryKey }).Count | Should -BeGreaterThan 0 }
    It 'tags nodes with a layer'        { @($script:Graph.nodes | Where-Object { $_.layer -in 'structure', 'data', 'both' }).Count | Should -Be $script:Graph.nodes.Count }

    It 'summary query returns counts' {
        $s = Get-ExcelGraphQuery -Action summary -GraphPath (Join-Path $script:OutDir 'graph.json')
        $s.nodeCount | Should -BeGreaterThan 0
        $s.edgesByKind.'datamodel-fk' | Should -BeGreaterOrEqual 5
    }
    It 'neighbors query resolves a table' {
        $n = Get-ExcelGraphQuery -Action neighbors -GraphPath (Join-Path $script:OutDir 'graph.json') -Node 'table:Orders' -Depth 1
        $n.node.id | Should -Be 'table:Orders'
    }
}
