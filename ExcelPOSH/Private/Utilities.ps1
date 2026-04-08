# Private/Utilities.ps1 — Value conversion and output formatting

function ConvertTo-ExcelSafeValue {
    <#
    .SYNOPSIS
        Convert COM values to PowerShell-safe types for JSON serialization.
    #>
    param([AllowNull()]$Value)

    if ($null -eq $Value)               { return $null }
    if ($Value -is [System.DBNull])     { return $null }
    if ($Value -is [System.DateTime])   { return $Value.ToString('o') }  # ISO 8601
    if ($Value -is [decimal])           { return [double]$Value }
    if ($Value -is [byte[]])            { return "<binary $($Value.Length) bytes>" }
    # Handle COM error values (e.g. #N/A, #VALUE!)
    if ($Value -is [int] -and $Value -eq -2146826246) { return '#N/A' }
    return $Value
}

function Format-ExcelOutput {
    <#
    .SYNOPSIS
        Handle -AsJson switch: convert hashtable/PSCustomObject to JSON or return as-is.
    #>
    param(
        [Parameter(Mandatory)]$Data,
        [switch]$AsJson
    )

    if ($Data -is [hashtable]) {
        $Data = [PSCustomObject]$Data
    }
    if ($AsJson) {
        return $Data | ConvertTo-Json -Depth 10 -Compress
    }
    return $Data
}
