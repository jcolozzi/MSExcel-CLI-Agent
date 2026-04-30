# Private/Session.ps1 — COM session management helpers

function Test-ExcelAlive {
    <#
    .SYNOPSIS
        Best-effort COM liveness check for Excel.
    #>
    if ($null -eq $script:ExcelSession.App) { return $false }
    $alive = $false
    try {
        $null = $script:ExcelSession.App.Hwnd
        $alive = $true
    }
    catch {
        try {
            $null = $script:ExcelSession.App.Version
            $alive = $true
        }
        catch {
            $alive = $false
        }
    }
    return $alive
}

function Get-ExcelHwnd {
    <#
    .SYNOPSIS
        Get the Excel window handle.
    #>
    param([Parameter(Mandatory)]$App)
    return [long]$App.Hwnd
}

function Set-ExcelVisibleBestEffort {
    <#
    .SYNOPSIS
        Try to set Excel visibility. Never fail startup if unsupported.
    #>
    param([bool]$Visible = $true)
    if ($null -eq $script:ExcelSession.App) { return }
    try {
        $script:ExcelSession.App.Visible = $Visible
    } catch {
        Write-Verbose "Could not set Excel.Visible=$Visible (continuing): $_"
    }
}

function Clear-ExcelCaches {
    <#
    .SYNOPSIS
        Reserved for future cache clearing. Currently a no-op.
    #>
    # No caches in current scope (no VBE/control caches)
}

function Connect-ExcelWorkbook {
    <#
    .SYNOPSIS
        Internal: ensure Excel COM is running and the requested workbook is open.
        Returns the COM Application object.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WorkbookPath
    )

    $resolved = [System.IO.Path]::GetFullPath($WorkbookPath)

    # If we have an existing session, check liveness
    if ($null -ne $script:ExcelSession.App) {
        if (-not (Test-ExcelAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            $script:ExcelSession.App          = $null
            $script:ExcelSession.WorkbookPath = $null
            Clear-ExcelCaches
        }
    }

    # Launch Excel if needed
    if ($null -eq $script:ExcelSession.App) {
        Write-Verbose 'Launching Excel.Application...'
        try {
            $script:ExcelSession.App = New-Object -ComObject 'Excel.Application'
        } catch {
            throw "Failed to create Excel.Application COM object. Is Microsoft Excel installed? Error: $_"
        }
        $script:ExcelSession.App.DisplayAlerts = $false
        Set-ExcelVisibleBestEffort -Visible $true
        Write-Verbose 'Excel launched OK'
    }

    # Switch workbook if needed
    if ($script:ExcelSession.WorkbookPath -ne $resolved) {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Workbook file not found: $resolved"
        }

        # Close previous workbook
        if ($null -ne $script:ExcelSession.WorkbookPath) {
            Write-Verbose "Closing previous workbook: $($script:ExcelSession.WorkbookPath)"
            try {
                $wb = $script:ExcelSession.App.Workbooks | Where-Object {
                    $_.FullName -eq $script:ExcelSession.WorkbookPath
                }
                if ($null -ne $wb) {
                    $wb.Close($false)  # don't save changes on switch
                }
            } catch {
                Write-Verbose "Error closing previous workbook: $_"
            }
        }

        # Open new workbook
        Write-Verbose "Opening workbook: $resolved"
        try {
            $null = $script:ExcelSession.App.Workbooks.Open($resolved)
        } catch {
            throw "Failed to open workbook '$resolved': $_"
        }

        $script:ExcelSession.WorkbookPath = $resolved
        Clear-ExcelCaches
        Write-Verbose "Workbook opened: $resolved"
    }

    return $script:ExcelSession.App
}
