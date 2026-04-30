# Public/WorkbookOps.ps1 — Workbook lifecycle: open, close, new, save, info, repair, copy

function Open-ExcelWorkbook {
    <#
    .SYNOPSIS
        Open an Excel workbook and establish a COM session.
    .PARAMETER WorkbookPath
        Path to the .xlsx/.xlsm/.xlsb/.xls file.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Open-ExcelWorkbook -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $result = @{
        status       = 'ok'
        workbook     = $wb.Name
        path         = $wb.FullName
        sheetCount   = $wb.Worksheets.Count
        activeSheet  = $wb.ActiveSheet.Name
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Close-ExcelWorkbook {
    <#
    .SYNOPSIS
        Close the Excel COM session and release the file lock.
    .DESCRIPTION
        Closes the current workbook, quits Excel, releases COM objects.
        Safe to call even if no session is open.
    .EXAMPLE
        Close-ExcelWorkbook
    #>
    [CmdletBinding()]
    param()

    if ($null -ne $script:ExcelSession.App) {
        Write-Verbose 'Closing Excel...'

        try {
            $script:ExcelSession.App.DisplayAlerts = $false
        } catch {}

        # Close all open workbooks without saving
        try {
            foreach ($wb in @($script:ExcelSession.App.Workbooks)) {
                try { $wb.Close($false) } catch {}
            }
        } catch {
            Write-Verbose "Error closing workbooks: $_"
        }

        try {
            $script:ExcelSession.App.Quit()
            Write-Verbose 'Excel quit OK'
        } catch {
            Write-Verbose "Error quitting Excel: $_"
        }

        # Brief wait then force-kill if needed
        Start-Sleep -Milliseconds 500
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelSession.App)
        } catch {}

        $script:ExcelSession.App          = $null
        $script:ExcelSession.WorkbookPath = $null
        Clear-ExcelCaches
        Write-Verbose 'Excel closed OK'
    }
}

function New-ExcelWorkbook {
    <#
    .SYNOPSIS
        Create a new empty Excel workbook.
    .PARAMETER WorkbookPath
        Full path for the new workbook file. Extension determines format (.xlsx, .xlsm, .xlsb, .xls).
    .PARAMETER SheetNames
        Optional array of sheet names. Default: one sheet named "Sheet1".
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-ExcelWorkbook -WorkbookPath "C:\data\new.xlsx"
    .EXAMPLE
        New-ExcelWorkbook -WorkbookPath "C:\data\new.xlsx" -SheetNames @("Data","Summary")
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string[]]$SheetNames,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (Test-Path -LiteralPath $resolved) {
        throw "File already exists: $resolved"
    }

    # Ensure parent directory exists
    $dir = [System.IO.Path]::GetDirectoryName($resolved)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    # Launch Excel if needed
    if ($null -eq $script:ExcelSession.App) {
        $script:ExcelSession.App = New-Object -ComObject 'Excel.Application'
        $script:ExcelSession.App.DisplayAlerts = $false
        Set-ExcelVisibleBestEffort -Visible $true
    }

    $app = $script:ExcelSession.App
    $wb  = $app.Workbooks.Add()

    # Set up sheet names if provided
    if ($null -ne $SheetNames -and $SheetNames.Count -gt 0) {
        # Add sheets as needed
        while ($wb.Worksheets.Count -lt $SheetNames.Count) {
            $null = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        }
        # Remove extras
        while ($wb.Worksheets.Count -gt $SheetNames.Count) {
            $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
        }
        # Rename
        for ($i = 0; $i -lt $SheetNames.Count; $i++) {
            $wb.Worksheets.Item($i + 1).Name = $SheetNames[$i]
        }
    }

    # Determine file format from extension
    $ext = [System.IO.Path]::GetExtension($resolved).ToLower().TrimStart('.')
    $fmt = if ($script:XL_FILE_FORMAT.ContainsKey($ext)) {
        $script:XL_FILE_FORMAT[$ext]
    } else {
        $script:XL_FILE_FORMAT['xlsx']
    }

    $wb.SaveAs($resolved, $fmt)
    $script:ExcelSession.WorkbookPath = $resolved

    $result = @{
        status     = 'created'
        path       = $resolved
        format     = $ext
        sheetCount = $wb.Worksheets.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Save-ExcelWorkbook {
    <#
    .SYNOPSIS
        Save the current workbook, optionally as a new file (Save As).
    .PARAMETER WorkbookPath
        Path to the open workbook.
    .PARAMETER SaveAsPath
        Optional new path for Save As. Extension determines format.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Save-ExcelWorkbook -WorkbookPath "C:\data.xlsx"
    .EXAMPLE
        Save-ExcelWorkbook -WorkbookPath "C:\data.xlsx" -SaveAsPath "C:\backup.xlsm"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [string]$SaveAsPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
        $wb.Save()
        $result = @{ status = 'saved'; path = $wb.FullName }
    } else {
        $resolvedNew = [System.IO.Path]::GetFullPath($SaveAsPath)
        $ext = [System.IO.Path]::GetExtension($resolvedNew).ToLower().TrimStart('.')
        $fmt = if ($script:XL_FILE_FORMAT.ContainsKey($ext)) {
            $script:XL_FILE_FORMAT[$ext]
        } else {
            $script:XL_FILE_FORMAT['xlsx']
        }
        $wb.SaveAs($resolvedNew, $fmt)
        $script:ExcelSession.WorkbookPath = $resolvedNew
        $result = @{ status = 'saved_as'; path = $resolvedNew; format = $ext }
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelWorkbookInfo {
    <#
    .SYNOPSIS
        Get metadata about the current workbook.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelWorkbookInfo -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook

    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
        $sheets += @{
            name    = $ws.Name
            index   = $ws.Index
            visible = $ws.Visible -eq -1  # xlSheetVisible = -1
        }
    }

    $result = @{
        name         = $wb.Name
        path         = $wb.FullName
        sheetCount   = $wb.Worksheets.Count
        sheets       = $sheets
        activeSheet  = $wb.ActiveSheet.Name
        readOnly     = $wb.ReadOnly
        saved        = $wb.Saved
        fileFormat   = $wb.FileFormat
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Repair-ExcelWorkbook {
    <#
    .SYNOPSIS
        Open a workbook with the CorruptLoad option to attempt repair.
    .PARAMETER WorkbookPath
        Path to the workbook to repair.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Repair-ExcelWorkbook -WorkbookPath "C:\data.xlsx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($WorkbookPath)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "Workbook file not found: $resolved"
    }

    # Close existing if open
    if ($null -ne $script:ExcelSession.App -and $script:ExcelSession.WorkbookPath -eq $resolved) {
        try {
            $script:ExcelSession.App.ActiveWorkbook.Close($false)
        } catch {}
        $script:ExcelSession.WorkbookPath = $null
    }

    # Launch Excel if needed
    if ($null -eq $script:ExcelSession.App -or -not (Test-ExcelAlive)) {
        $script:ExcelSession.App = New-Object -ComObject 'Excel.Application'
        $script:ExcelSession.App.DisplayAlerts = $false
        Set-ExcelVisibleBestEffort -Visible $true
    }

    $app = $script:ExcelSession.App
    # xlRepairFile = 2
    $null = $app.Workbooks.Open($resolved, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value,
        2)  # CorruptLoad = xlRepairFile

    $script:ExcelSession.WorkbookPath = $resolved
    $wb = $app.ActiveWorkbook

    $result = @{
        status     = 'repaired'
        path       = $wb.FullName
        sheetCount = $wb.Worksheets.Count
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Copy-ExcelWorkbook {
    <#
    .SYNOPSIS
        Copy/backup a workbook file.
    .PARAMETER WorkbookPath
        Path to the source workbook.
    .PARAMETER DestinationPath
        Path for the copy.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-ExcelWorkbook -WorkbookPath "C:\data.xlsx" -DestinationPath "C:\backup\data_backup.xlsx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [switch]$AsJson
    )

    $resolvedSrc  = [System.IO.Path]::GetFullPath($WorkbookPath)
    $resolvedDest = [System.IO.Path]::GetFullPath($DestinationPath)

    if (-not (Test-Path -LiteralPath $resolvedSrc -PathType Leaf)) {
        throw "Source workbook not found: $resolvedSrc"
    }

    # Ensure destination directory exists
    $destDir = [System.IO.Path]::GetDirectoryName($resolvedDest)
    if (-not (Test-Path -LiteralPath $destDir)) {
        $null = New-Item -Path $destDir -ItemType Directory -Force
    }

    Copy-Item -LiteralPath $resolvedSrc -Destination $resolvedDest -Force

    $result = @{
        status      = 'copied'
        source      = $resolvedSrc
        destination = $resolvedDest
        sizeBytes   = (Get-Item -LiteralPath $resolvedDest).Length
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Invoke-ExcelMacro {
    <#
    .SYNOPSIS  Run a VBA macro in the workbook.
    .PARAMETER WorkbookPath  Path to the Excel workbook (.xlsm or .xlsb).
    .PARAMETER MacroName     Full macro name (e.g. "Sheet1.MyMacro" or just "MyMacro").
    .PARAMETER Arguments     Up to 30 arguments to pass to the macro.
    .PARAMETER AsJson        Return JSON string.
    .EXAMPLE   Invoke-ExcelMacro -WorkbookPath C:\data.xlsm -MacroName "CleanData" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$MacroName,
        [ValidateCount(0, 30)][object[]]$Arguments,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath

    # Build argument list: macro name + up to 30 args
    $runArgs = @($MacroName) + @($Arguments)

    $result = $app.GetType().InvokeMember(
        'Run',
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null,
        $app,
        $runArgs
    )

    $output = @{
        status    = 'ok'
        macroName = $MacroName
        result    = $result
    }
    Format-ExcelOutput -Data $output -AsJson:$AsJson
}
