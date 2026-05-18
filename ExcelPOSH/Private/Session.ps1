# Private/Session.ps1 — COM session management helpers

function Get-RunningComApp {
    <#
    .SYNOPSIS
        Try to attach to an already-running COM application via the Running Object Table.
        Returns the COM object or $null.  Works on Windows PowerShell 5.1 (Desktop);
        gracefully degrades on PowerShell 7+ where [Marshal]::GetActiveObject is unavailable.
    #>
    param(
        [Parameter(Mandatory)][string]$ProgId,
        [Parameter(Mandatory)][string]$ProcessName
    )

    # Fast exit: if the host process isn't running, skip the COM probe entirely
    if (-not (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
        Write-Verbose "Get-RunningComApp: no $ProcessName process found — skipping ROT lookup."
        return $null
    }

    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        Write-Verbose "Get-RunningComApp: attached to existing $ProgId instance."
        return $app
    }
    catch [System.Management.Automation.MethodException] {
        # .NET Core / PS7 — GetActiveObject does not exist
        Write-Verbose "Get-RunningComApp: [Marshal]::GetActiveObject unavailable (PowerShell $($PSVersionTable.PSVersion)) — will create new instance."
        return $null
    }
    catch {
        # No ROT entry, or stale/dead entry
        Write-Verbose "Get-RunningComApp: could not attach to $ProgId — $($_.Exception.Message)"
        return $null
    }
}

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
        Tries to attach to an already-running Excel instance (GetObject-first)
        before creating a new one, to prevent duplicate instances.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$WorkbookPath,
        [switch]$ForceNewInstance
    )

    $resolved = [System.IO.Path]::GetFullPath($WorkbookPath)

    # If we have an existing session, check liveness
    if ($null -ne $script:ExcelSession.App) {
        if (-not (Test-ExcelAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            $script:ExcelSession.App          = $null
            $script:ExcelSession.WorkbookPath = $null
            $script:ExcelSession.OwnsApp      = $false
            Clear-ExcelCaches
        }
    }

    # Acquire Excel instance if needed (GetObject-first, then New-Object)
    if ($null -eq $script:ExcelSession.App) {
        $adopted = $false

        # Try to attach to an existing Excel instance via the ROT
        if (-not $ForceNewInstance) {
            $existing = Get-RunningComApp -ProgId 'Excel.Application' -ProcessName 'EXCEL'
            if ($null -ne $existing) {
                Write-Verbose 'Adopting existing Excel instance'
                $script:ExcelSession.App     = $existing
                $script:ExcelSession.OwnsApp = $false
                $adopted = $true
                $script:ExcelSession.App.DisplayAlerts = $false
                Set-ExcelVisibleBestEffort -Visible $true
                Write-Verbose 'Adopted existing Excel instance OK'
            }
        }

        # Fall back to creating a new instance
        if (-not $adopted) {
            Write-Verbose 'Launching Excel.Application...'
            try {
                $script:ExcelSession.App = New-Object -ComObject 'Excel.Application'
            } catch {
                throw "Failed to create Excel.Application COM object. Is Microsoft Excel installed? Error: $_"
            }
            $script:ExcelSession.OwnsApp = $true
            $script:ExcelSession.App.DisplayAlerts = $false
            Set-ExcelVisibleBestEffort -Visible $true
            Write-Verbose 'Excel launched OK'
        }
    }

    # Switch workbook if needed
    if ($script:ExcelSession.WorkbookPath -ne $resolved) {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Workbook file not found: $resolved"
        }

        # Check if the workbook is already open in this instance
        $alreadyOpen = $false
        try {
            foreach ($wb in $script:ExcelSession.App.Workbooks) {
                if ($wb.FullName -eq $resolved) {
                    Write-Verbose "Workbook already open in this instance: $resolved"
                    $alreadyOpen = $true
                    break
                }
            }
        } catch {
            Write-Verbose "Error checking open workbooks: $_"
        }

        # Close previous workbook (only if it's different from the one we're opening)
        if (-not $alreadyOpen -and $null -ne $script:ExcelSession.WorkbookPath) {
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

        # Open workbook if not already open
        if (-not $alreadyOpen) {
            Write-Verbose "Opening workbook: $resolved"
            try {
                $null = $script:ExcelSession.App.Workbooks.Open($resolved)
            } catch {
                throw "Failed to open workbook '$resolved': $_"
            }
        }

        $script:ExcelSession.WorkbookPath = $resolved
        Clear-ExcelCaches
        Write-Verbose "Workbook opened: $resolved"
    }

    return $script:ExcelSession.App
}
