# Public/ImageShapeOps.ps1 — Images, shapes, and graphic objects

function Add-ExcelImage {
    <#
    .SYNOPSIS
        Insert an image from file into a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the target worksheet.
    .PARAMETER ImagePath
        Path to the image file to insert.
    .PARAMETER CellAddress
        Anchor cell address (e.g. "B2"). Defaults to "A1".
    .PARAMETER Width
        Width in points. If omitted, keeps original image width.
    .PARAMETER Height
        Height in points. If omitted, keeps original image height.
    .PARAMETER LinkToFile
        If set, the picture is linked rather than embedded.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-ExcelImage -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -ImagePath "C:\logo.png" -CellAddress "B2" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ImagePath,
        [string]$CellAddress = 'A1',
        [double]$Width,
        [double]$Height,
        [switch]$LinkToFile,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    if (-not (Test-Path $ImagePath)) {
        throw "Image file not found: $ImagePath"
    }

    $cell    = $ws.Range($CellAddress)
    $left    = $cell.Left
    $top     = $cell.Top
    $linkVal = if ($LinkToFile) { -1 } else { 0 }   # msoTrue=-1, msoFalse=0

    $pic = $ws.Shapes.AddPicture(
        [string](Resolve-Path $ImagePath),
        $linkVal,
        -1,       # SaveWithDocument = msoTrue
        $left,
        $top,
        -1,       # OriginalWidth
        -1        # OriginalHeight
    )

    if ($PSBoundParameters.ContainsKey('Width'))  { $pic.Width  = $Width  }
    if ($PSBoundParameters.ContainsKey('Height')) { $pic.Height = $Height }

    $result = @{
        status = 'added'
        name   = $pic.Name
        left   = $pic.Left
        top    = $pic.Top
        width  = $pic.Width
        height = $pic.Height
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Get-ExcelShape {
    <#
    .SYNOPSIS
        Get info about shapes (images, AutoShapes, etc.) in a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the target worksheet.
    .PARAMETER ShapeName
        Name of a specific shape. If omitted, lists all shapes.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-ExcelShape -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [string]$ShapeName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $typeNames = @{
        1  = 'AutoShape'
        2  = 'Callout'
        3  = 'Chart'
        4  = 'Comment'
        5  = 'FreeForm'
        6  = 'Group'
        7  = 'EmbeddedOLEObject'
        8  = 'FormControl'
        9  = 'Line'
        10 = 'LinkedOLEObject'
        11 = 'LinkedPicture'
        12 = 'OLEControlObject'
        13 = 'Picture'
        14 = 'Placeholder'
        15 = 'MediaObject'
        16 = 'ContentApp'
        17 = 'Media'
        18 = 'TextBox'
        19 = 'ScriptAnchor'
        20 = 'Table'
        21 = 'Canvas'
        22 = 'Diagram'
        23 = 'Ink'
        24 = 'InkComment'
        25 = 'Model3D'
    }

    $getShapeInfo = {
        param($s)
        $typeNum  = [int]$s.Type
        $typeName = if ($typeNames.ContainsKey($typeNum)) { $typeNames[$typeNum] } else { 'Unknown' }
        $altText  = try { $s.AlternativeText } catch { '' }
        @{
            name            = $s.Name
            type            = $typeNum
            type_name       = $typeName
            left            = $s.Left
            top             = $s.Top
            width           = $s.Width
            height          = $s.Height
            visible         = [bool]($s.Visible -eq -1)
            alternativeText = $altText
        }
    }

    if ($ShapeName) {
        $shape  = $ws.Shapes.Item($ShapeName)
        $info   = & $getShapeInfo $shape
        $result = @{
            status = 'ok'
            shapes = @($info)
            count  = 1
        }
    } else {
        $shapes = @()
        foreach ($s in $ws.Shapes) {
            $shapes += & $getShapeInfo $s
        }
        $result = @{
            status = 'ok'
            shapes = $shapes
            count  = $shapes.Count
        }
    }

    Format-ExcelOutput -Data $result -AsJson:$AsJson
}

function Remove-ExcelShape {
    <#
    .SYNOPSIS
        Delete a shape by name from a worksheet.
    .PARAMETER WorkbookPath
        Path to the workbook.
    .PARAMETER SheetName
        Name of the target worksheet.
    .PARAMETER ShapeName
        Name of the shape to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-ExcelShape -WorkbookPath "C:\data.xlsx" -SheetName "Sheet1" -ShapeName "Picture 1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkbookPath,
        [Parameter(Mandatory)][string]$SheetName,
        [Parameter(Mandatory)][string]$ShapeName,
        [switch]$AsJson
    )

    $app = Connect-ExcelWorkbook -WorkbookPath $WorkbookPath
    $wb  = $app.ActiveWorkbook
    $ws  = $wb.Worksheets.Item($SheetName)

    $shape = $ws.Shapes.Item($ShapeName)
    $shape.Delete()

    $result = @{
        status = 'removed'
        name   = $ShapeName
    }
    Format-ExcelOutput -Data $result -AsJson:$AsJson
}
