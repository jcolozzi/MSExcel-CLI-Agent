# Public/MetadataOps.ps1 — Document properties, connections, protection, comments, tips

function Get-ExcelDocumentProperty {
    <#
    .SYNOPSIS
        Get built-in or custom document properties.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER PropertyName
        Optional: specific property name. If omitted, returns all built-in properties.
    .PARAMETER Custom
        Include custom properties instead of built-in.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelDocumentProperty -WorkbookPath "C:\data.xlsx" -AsJson
    .EXAMPLE
        Get-ExcelDocumentProperty -WorkbookPath "C:\data.xlsx" -PropertyName "Title" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$PropertyName,
        [switch]$Custom,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $propCollection = if ($Custom) { $wb.CustomDocumentProperties } else { $wb.BuiltinDocumentProperties }

    $props = @()
    if (-not [string]::IsNullOrWhiteSpace($PropertyName)) {
        try {
            $p = $propCollection.Item($PropertyName)
            $props += @{
                name  = $PropertyName
                value = (ConvertTo-ExcelSafeValue $p.Value)
                type  = $p.Type
            }
        } catch {
            $props += @{
                name  = $PropertyName
                value = $null
                error = "Property not found: $PropertyName"
            }
        }
    } else {
        foreach ($p in $propCollection) {
            try {
                $props += @{
                    name  = $p.Name
                    value = (ConvertTo-ExcelSafeValue $p.Value)
                }
            } catch {
                $props += @{
                    name  = $p.Name
                    value = $null
                    error = 'Could not read value'
                }
            }
        }
    }

    $result = @{
        status = 'ok'
        type   = if ($Custom) { 'custom' } else { 'builtin' }
        count  = $props.Count
        properties = $props
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelDocumentProperty {
    <#
    .SYNOPSIS
        Set a built-in or custom document property.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER PropertyName
        Property name.
    .PARAMETER Value
        Value to set.
    .PARAMETER Custom
        Target custom properties instead of built-in.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelDocumentProperty -WorkbookPath "C:\data.xlsx" -PropertyName "Title" -Value "Sales Report" -AsJson
    .EXAMPLE
        Set-ExcelDocumentProperty -WorkbookPath "C:\data.xlsx" -PropertyName "Department" -Value "Finance" -Custom -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)]$Value,
        [switch]$Custom,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    if ($Custom) {
        # Try to update existing, add if not found
        $found = $false
        try {
            $p = $wb.CustomDocumentProperties.Item($PropertyName)
            $p.Value = $Value
            $found = $true
        } catch {}

        if (-not $found) {
            # msoPropertyTypeString = 4
            $null = $wb.CustomDocumentProperties.Add($PropertyName, $false, 4, $Value)
        }
    } else {
        $p = $wb.BuiltinDocumentProperties.Item($PropertyName)
        $p.Value = $Value
    }

    $result = @{
        status = 'ok'
        name   = $PropertyName
        value  = $Value
        type   = if ($Custom) { 'custom' } else { 'builtin' }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelConnection {
    <#
    .SYNOPSIS
        List data connections in the workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelConnection -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $connections = @()
    foreach ($conn in $wb.Connections) {
        $entry = @{
            name        = $conn.Name
            description = $conn.Description
            type        = $conn.Type
        }
        # Try to get connection string (may not be available for all types)
        try {
            $entry['connectionString'] = $conn.ODBCConnection.Connection
        } catch {
            try {
                $entry['connectionString'] = $conn.OLEDBConnection.Connection
            } catch {
                $entry['connectionString'] = $null
            }
        }
        $connections += $entry
    }

    $result = @{
        status      = 'ok'
        count       = $connections.Count
        connections = $connections
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelProtection {
    <#
    .SYNOPSIS
        Get protection status for a worksheet or workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Optional: check specific sheet. If omitted, checks workbook-level protection.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelProtection -WorkbookPath "C:\data.xlsx" -AsJson
    .EXAMPLE
        Get-ExcelProtection -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$SheetName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    if ([string]::IsNullOrWhiteSpace($SheetName)) {
        # Workbook-level protection
        $result = @{
            status           = 'ok'
            level            = 'workbook'
            protectStructure = $wb.ProtectStructure
            protectWindows   = $wb.ProtectWindows
        }
    } else {
        # Sheet-level protection
        $ws = $wb.Worksheets.Item($SheetName)
        $result = @{
            status          = 'ok'
            level           = 'sheet'
            sheet           = $SheetName
            protectContents = $ws.ProtectContents
            protectDrawingObjects = $ws.ProtectDrawingObjects
            protectScenarios = $ws.ProtectScenarios
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelProtection {
    <#
    .SYNOPSIS
        Protect or unprotect a worksheet or workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Optional: protect specific sheet. If omitted, protects workbook structure.
    .PARAMETER Password
        Optional password for protection.
    .PARAMETER Unprotect
        Remove protection instead of adding it.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelProtection -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Password "secret" -AsJson
    .EXAMPLE
        Set-ExcelProtection -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Unprotect -Password "secret" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$SheetName,
        [string]$Password,
        [switch]$Unprotect,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $action = if ($Unprotect) { 'unprotected' } else { 'protected' }

    if ([string]::IsNullOrWhiteSpace($SheetName)) {
        # Workbook-level
        if ($Unprotect) {
            if ([string]::IsNullOrWhiteSpace($Password)) { $wb.Unprotect() }
            else { $wb.Unprotect($Password) }
        } else {
            if ([string]::IsNullOrWhiteSpace($Password)) { $wb.Protect() }
            else { $wb.Protect($Password) }
        }
        $result = @{ status = $action; level = 'workbook' }
    } else {
        # Sheet-level
        $ws = $wb.Worksheets.Item($SheetName)
        if ($Unprotect) {
            if ([string]::IsNullOrWhiteSpace($Password)) { $ws.Unprotect() }
            else { $ws.Unprotect($Password) }
        } else {
            if ([string]::IsNullOrWhiteSpace($Password)) { $ws.Protect() }
            else { $ws.Protect($Password) }
        }
        $result = @{ status = $action; level = 'sheet'; sheet = $SheetName }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelComment {
    <#
    .SYNOPSIS
        Get comments (notes) from a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Optional: specific cell. If omitted, returns all comments on the sheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelComment -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $comments = @()

    if (-not [string]::IsNullOrWhiteSpace($Range)) {
        $cell = $ws.Range($Range)
        $c = $cell.Comment
        if ($null -ne $c) {
            $comments += @{
                address = $cell.Address($false, $false)
                author  = $c.Author
                text    = $c.Text()
            }
        }
    } else {
        foreach ($c in $ws.Comments) {
            $comments += @{
                address = $c.Parent.Address($false, $false)
                author  = $c.Author
                text    = $c.Text()
            }
        }
    }

    $result = @{
        status   = 'ok'
        sheet    = $SheetName
        count    = $comments.Count
        comments = $comments
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Set-ExcelComment {
    <#
    .SYNOPSIS
        Add, update, or remove a comment on a cell.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell address for the comment.
    .PARAMETER Text
        Comment text. If empty/null with -Remove, deletes the comment.
    .PARAMETER Remove
        Remove the comment from the cell.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-ExcelComment -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1" -Text "Check this value" -AsJson
    .EXAMPLE
        Set-ExcelComment -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -Range "A1" -Remove -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [string]$Text,
        [switch]$Remove,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)
    $cell = $ws.Range($Range)

    if ($Remove) {
        if ($null -ne $cell.Comment) {
            $cell.Comment.Delete()
        }
        $result = @{
            status  = 'removed'
            address = $cell.Address($false, $false)
        }
    } else {
        # Remove existing then add fresh
        if ($null -ne $cell.Comment) {
            $cell.Comment.Delete()
        }
        $null = $cell.AddComment($Text)
        $result = @{
            status  = 'added'
            address = $cell.Address($false, $false)
            text    = $Text
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelTip {
    <#
    .SYNOPSIS
        Return a random ExcelPOSH usage tip.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelTip -AsJson
    #>
    [CmdletBinding()]
    param([switch]$AsJson)

    $tips = @(
        'Use Get-ExcelRange to read any cell/range. Value2 returns raw numeric dates — use ConvertTo-ExcelSafeValue for safety.'
        'Set-ExcelRange accepts 2D arrays: @(@(1,2),@(3,4)) writes a 2x2 block.'
        'Excel tables (ListObjects) are powerful: Get-ExcelTable, New-ExcelTable, Get-ExcelTableData.'
        'Use Set-ExcelCellFormat with hex colors: -FontColor "#FF0000" -FillColor "#FFFF00".'
        'Auto-fit columns with Set-ExcelColumnWidth -Width 0.'
        'Find-ExcelValue searches for text in values, formulas, or comments.'
        'Always call Close-ExcelWorkbook when done to release the COM lock.'
        'Use Save-ExcelWorkbook -SaveAsPath to convert between formats (xlsx, xlsm, csv).'
        'Set-ExcelTableTotals adds a totals row with sum, average, count, etc.'
        'Named ranges help with formulas: New-ExcelNamedRange, Get-ExcelNamedRange.'
        'Get-ExcelWorkbookInfo shows all sheets, read-only status, and file format.'
        'Protect sheets with Set-ExcelProtection -SheetName "Sheet1" -Password "secret".'
        'Copy-ExcelWorkbook makes backups before risky operations.'
        'Use -AsJson on any function for structured output that is easy to parse.'
        'Set-ExcelAlignment supports merge, wrap text, and all horizontal/vertical options.'
    )

    $tip = $tips | Get-Random
    $result = @{
        status = 'ok'
        tip    = $tip
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelThreadedComment {
    <#
    .SYNOPSIS
        Add a modern threaded comment to a cell (Excel 365).
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell address for the comment.
    .PARAMETER Text
        Comment text.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelThreadedComment -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1" -Text "Please verify" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$Text,
        [switch]$AsJson
    )

    $app  = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws   = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $cell = $ws.Range($Range)

    try {
        $null = $cell.AddCommentThreaded($Text)
        $result = @{
            status  = 'ok'
            address = $cell.Address($false, $false)
            text    = $Text
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Threaded comments unavailable (requires Excel 365): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelThreadedComment {
    <#
    .SYNOPSIS
        Read threaded comments (and their replies) from a worksheet.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Optional: a specific cell. If omitted, returns all threaded comments on the sheet.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelThreadedComment -WorkbookPath C:\data.xlsx -SheetName Sheet1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$Range,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws  = $app.ActiveWorkbook.Worksheets.Item($SheetName)

    $comments = @()
    try {
        $source = if (-not [string]::IsNullOrWhiteSpace($Range)) {
            $ct = $ws.Range($Range).CommentThreaded
            if ($null -ne $ct) { @($ct) } else { @() }
        } else {
            $ws.CommentsThreaded
        }

        foreach ($ct in $source) {
            if ($null -eq $ct) { continue }
            $replies = @()
            foreach ($r in $ct.Replies) {
                $replies += @{ author = $r.Author.Name; text = $r.Text() }
            }
            $comments += @{
                address    = $ct.Parent.Address($false, $false)
                author     = $ct.Author.Name
                text       = $ct.Text()
                replyCount = $ct.Replies.Count
                replies    = $replies
            }
        }
        $result = @{
            status   = 'ok'
            sheet    = $SheetName
            count    = $comments.Count
            comments = $comments
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Threaded comments unavailable (requires Excel 365): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Add-ExcelThreadedCommentReply {
    <#
    .SYNOPSIS
        Add a reply to an existing threaded comment.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell containing the threaded comment.
    .PARAMETER Text
        Reply text.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelThreadedCommentReply -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1" -Text "Verified" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [Parameter(Mandatory)][string]$Text,
        [switch]$AsJson
    )

    $app  = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws   = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $cell = $ws.Range($Range)

    try {
        $ct = $cell.CommentThreaded
        if ($null -eq $ct) {
            $result = @{ status = 'error'; error = "No threaded comment at $Range" }
        } else {
            $null = $ct.AddReply($Text)
            $result = @{
                status  = 'ok'
                address = $cell.Address($false, $false)
                reply   = $Text
            }
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Threaded comments unavailable (requires Excel 365): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelThreadedComment {
    <#
    .SYNOPSIS
        Delete a threaded comment (and its replies) from a cell.
    .PARAMETER WorkbookPath
        Path to the Excel workbook.
    .PARAMETER SheetName
        Worksheet name.
    .PARAMETER Range
        Cell containing the threaded comment.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelThreadedComment -WorkbookPath C:\data.xlsx -SheetName Sheet1 -Range "A1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$Range,
        [switch]$AsJson
    )

    $app  = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $ws   = $app.ActiveWorkbook.Worksheets.Item($SheetName)
    $cell = $ws.Range($Range)

    try {
        $ct = $cell.CommentThreaded
        if ($null -ne $ct) { $ct.Delete() }
        $result = @{
            status  = 'ok'
            address = $cell.Address($false, $false)
            deleted = ($null -ne $ct)
        }
    } catch {
        $result = @{
            status = 'unsupported'
            error  = "Threaded comments unavailable (requires Excel 365): $($_.Exception.Message)"
        }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
