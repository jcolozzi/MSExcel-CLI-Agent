# Tests/V4ExtensionOps.Tests.ps1
# Pester 5+ tests — parameter validation for v4.0 functions added to existing Ops files
# (WorksheetOps, CalculationOps, WorkbookOps, ImportOps, FilterSortOps, FormattingOps, PrintOps, MetadataOps)

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
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ─── WorksheetOps additions ──────────────────────────────────────────────────
Describe 'Set-ExcelSheetTab' {
    It 'Has CmdletBinding' { (Get-Command Set-ExcelSheetTab).CmdletBinding | Should -BeTrue }
    It 'WorkbookPath mandatory' { Assert-Mandatory Set-ExcelSheetTab WorkbookPath }
    It 'SheetName mandatory'    { Assert-Mandatory Set-ExcelSheetTab SheetName }
    It 'Color mandatory'        { Assert-Mandatory Set-ExcelSheetTab Color }
    It 'Has AsJson switch'      { (Get-Command Set-ExcelSheetTab).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Invoke-ExcelAutoFill' {
    It 'Has CmdletBinding' { (Get-Command Invoke-ExcelAutoFill).CmdletBinding | Should -BeTrue }
    It 'SourceRange mandatory'      { Assert-Mandatory Invoke-ExcelAutoFill SourceRange }
    It 'DestinationRange mandatory' { Assert-Mandatory Invoke-ExcelAutoFill DestinationRange }
    It 'Has FillType parameter'     { (Get-Command Invoke-ExcelAutoFill).Parameters['FillType'] | Should -Not -BeNullOrEmpty }
    It 'Has AsJson switch'          { (Get-Command Invoke-ExcelAutoFill).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Set-ExcelFormula2' {
    It 'Has CmdletBinding' { (Get-Command Set-ExcelFormula2).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'   { Assert-Mandatory Set-ExcelFormula2 Range }
    It 'Formula mandatory' { Assert-Mandatory Set-ExcelFormula2 Formula }
    It 'Has AsJson switch' { (Get-Command Set-ExcelFormula2).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Get-ExcelFormulaDependencies' {
    It 'Has CmdletBinding' { (Get-Command Get-ExcelFormulaDependencies).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'   { Assert-Mandatory Get-ExcelFormulaDependencies Range }
    It 'Has Relation parameter' { (Get-Command Get-ExcelFormulaDependencies).Parameters['Relation'] | Should -Not -BeNullOrEmpty }
    It 'Has AsJson switch' { (Get-Command Get-ExcelFormulaDependencies).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Convert-ExcelToLinkedDataType' {
    It 'Has CmdletBinding' { (Get-Command Convert-ExcelToLinkedDataType).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'    { Assert-Mandatory Convert-ExcelToLinkedDataType Range }
    It 'DataType mandatory' { Assert-Mandatory Convert-ExcelToLinkedDataType DataType }
    It 'Has AsJson switch'  { (Get-Command Convert-ExcelToLinkedDataType).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Add-ExcelScenario' {
    It 'Has CmdletBinding' { (Get-Command Add-ExcelScenario).CmdletBinding | Should -BeTrue }
    It 'Name mandatory'          { Assert-Mandatory Add-ExcelScenario Name }
    It 'ChangingCells mandatory' { Assert-Mandatory Add-ExcelScenario ChangingCells }
    It 'Values mandatory'        { Assert-Mandatory Add-ExcelScenario Values }
    It 'Has AsJson switch'       { (Get-Command Add-ExcelScenario).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Get-ExcelScenario' {
    It 'Has CmdletBinding' { (Get-Command Get-ExcelScenario).CmdletBinding | Should -BeTrue }
    It 'SheetName mandatory' { Assert-Mandatory Get-ExcelScenario SheetName }
    It 'Has AsJson switch'   { (Get-Command Get-ExcelScenario).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── CalculationOps addition ─────────────────────────────────────────────────
Describe 'Invoke-ExcelGoalSeek' {
    It 'Has CmdletBinding' { (Get-Command Invoke-ExcelGoalSeek).CmdletBinding | Should -BeTrue }
    It 'TargetCell mandatory'   { Assert-Mandatory Invoke-ExcelGoalSeek TargetCell }
    It 'TargetValue mandatory'  { Assert-Mandatory Invoke-ExcelGoalSeek TargetValue }
    It 'ChangingCell mandatory' { Assert-Mandatory Invoke-ExcelGoalSeek ChangingCell }
    It 'Has AsJson switch'      { (Get-Command Invoke-ExcelGoalSeek).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── WorkbookOps addition ────────────────────────────────────────────────────
Describe 'Set-ExcelStatusBar' {
    It 'Has CmdletBinding' { (Get-Command Set-ExcelStatusBar).CmdletBinding | Should -BeTrue }
    It 'WorkbookPath mandatory' { Assert-Mandatory Set-ExcelStatusBar WorkbookPath }
    It 'Has Reset switch'  { (Get-Command Set-ExcelStatusBar).Parameters['Reset'].SwitchParameter | Should -BeTrue }
    It 'Has AsJson switch' { (Get-Command Set-ExcelStatusBar).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── ImportOps additions ─────────────────────────────────────────────────────
Describe 'Split-ExcelColumn' {
    It 'Has CmdletBinding' { (Get-Command Split-ExcelColumn).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'   { Assert-Mandatory Split-ExcelColumn Range }
    It 'Has Delimiter parameter' { (Get-Command Split-ExcelColumn).Parameters['Delimiter'] | Should -Not -BeNullOrEmpty }
    It 'Has AsJson switch' { (Get-Command Split-ExcelColumn).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Import-ExcelRecordset' {
    It 'Has CmdletBinding' { (Get-Command Import-ExcelRecordset).CmdletBinding | Should -BeTrue }
    It 'Destination mandatory'      { Assert-Mandatory Import-ExcelRecordset Destination }
    It 'ConnectionString mandatory' { Assert-Mandatory Import-ExcelRecordset ConnectionString }
    It 'Query mandatory'            { Assert-Mandatory Import-ExcelRecordset Query }
    It 'Has AsJson switch'          { (Get-Command Import-ExcelRecordset).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── FilterSortOps addition ──────────────────────────────────────────────────
Describe 'Add-ExcelSubtotal' {
    It 'Has CmdletBinding' { (Get-Command Add-ExcelSubtotal).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'        { Assert-Mandatory Add-ExcelSubtotal Range }
    It 'GroupBy mandatory'      { Assert-Mandatory Add-ExcelSubtotal GroupBy }
    It 'Function mandatory'     { Assert-Mandatory Add-ExcelSubtotal Function }
    It 'TotalColumns mandatory' { Assert-Mandatory Add-ExcelSubtotal TotalColumns }
    It 'Has AsJson switch'      { (Get-Command Add-ExcelSubtotal).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── FormattingOps additions ─────────────────────────────────────────────────
Describe 'New-ExcelStyle' {
    It 'Has CmdletBinding' { (Get-Command New-ExcelStyle).CmdletBinding | Should -BeTrue }
    It 'Name mandatory'    { Assert-Mandatory New-ExcelStyle Name }
    It 'Has NumberFormat parameter' { (Get-Command New-ExcelStyle).Parameters['NumberFormat'] | Should -Not -BeNullOrEmpty }
    It 'Has AsJson switch' { (Get-Command New-ExcelStyle).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Set-ExcelRangeStyle' {
    It 'Has CmdletBinding' { (Get-Command Set-ExcelRangeStyle).CmdletBinding | Should -BeTrue }
    It 'Range mandatory'     { Assert-Mandatory Set-ExcelRangeStyle Range }
    It 'StyleName mandatory' { Assert-Mandatory Set-ExcelRangeStyle StyleName }
    It 'Has AsJson switch'   { (Get-Command Set-ExcelRangeStyle).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Get-ExcelStyle' {
    It 'Has CmdletBinding' { (Get-Command Get-ExcelStyle).CmdletBinding | Should -BeTrue }
    It 'WorkbookPath mandatory' { Assert-Mandatory Get-ExcelStyle WorkbookPath }
    It 'Has AsJson switch' { (Get-Command Get-ExcelStyle).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── PrintOps additions ──────────────────────────────────────────────────────
Describe 'Export-ExcelToPdf (Format param)' {
    It 'Has Format parameter' { (Get-Command Export-ExcelToPdf).Parameters['Format'] | Should -Not -BeNullOrEmpty }
    It 'Format accepts PDF and XPS' {
        $vs = (Get-Command Export-ExcelToPdf).Parameters['Format'].Attributes.Where({ $_ -is [System.Management.Automation.ValidateSetAttribute] })
        $vs.ValidValues | Should -Contain 'PDF'
        $vs.ValidValues | Should -Contain 'XPS'
    }
}

Describe 'Send-ExcelPrint' {
    It 'Has CmdletBinding' { (Get-Command Send-ExcelPrint).CmdletBinding | Should -BeTrue }
    It 'SheetName mandatory' { Assert-Mandatory Send-ExcelPrint SheetName }
    It 'Has Preview switch' { (Get-Command Send-ExcelPrint).Parameters['Preview'].SwitchParameter | Should -BeTrue }
    It 'Has AsJson switch'  { (Get-Command Send-ExcelPrint).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

# ─── MetadataOps additions (threaded comments) ───────────────────────────────
Describe 'Add-ExcelThreadedComment' {
    It 'Has CmdletBinding' { (Get-Command Add-ExcelThreadedComment).CmdletBinding | Should -BeTrue }
    It 'Range mandatory' { Assert-Mandatory Add-ExcelThreadedComment Range }
    It 'Text mandatory'  { Assert-Mandatory Add-ExcelThreadedComment Text }
    It 'Has AsJson switch' { (Get-Command Add-ExcelThreadedComment).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Get-ExcelThreadedComment' {
    It 'Has CmdletBinding' { (Get-Command Get-ExcelThreadedComment).CmdletBinding | Should -BeTrue }
    It 'SheetName mandatory' { Assert-Mandatory Get-ExcelThreadedComment SheetName }
    It 'Has AsJson switch' { (Get-Command Get-ExcelThreadedComment).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Add-ExcelThreadedCommentReply' {
    It 'Has CmdletBinding' { (Get-Command Add-ExcelThreadedCommentReply).CmdletBinding | Should -BeTrue }
    It 'Range mandatory' { Assert-Mandatory Add-ExcelThreadedCommentReply Range }
    It 'Text mandatory'  { Assert-Mandatory Add-ExcelThreadedCommentReply Text }
    It 'Has AsJson switch' { (Get-Command Add-ExcelThreadedCommentReply).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}

Describe 'Remove-ExcelThreadedComment' {
    It 'Has CmdletBinding' { (Get-Command Remove-ExcelThreadedComment).CmdletBinding | Should -BeTrue }
    It 'Range mandatory' { Assert-Mandatory Remove-ExcelThreadedComment Range }
    It 'Has AsJson switch' { (Get-Command Remove-ExcelThreadedComment).Parameters['AsJson'].SwitchParameter | Should -BeTrue }
}
