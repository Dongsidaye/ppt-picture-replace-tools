param(
    [string]$OutputPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'icon_layers.png')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundedPath([float]$x, [float]$y, [float]$width, [float]$height, [float]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = $radius * 2
    $path.AddArc($x, $y, $diameter, $diameter, 180, 90)
    $path.AddArc($x + $width - $diameter, $y, $diameter, $diameter, 270, 90)
    $path.AddArc($x + $width - $diameter, $y + $height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($x, $y + $height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

$hi = 128
$bitmap = New-Object System.Drawing.Bitmap($hi, $hi, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$graphics.Clear([System.Drawing.Color]::Transparent)

# Same gradient rounded square as the filter icon so the pair reads as family.
$bgPath = New-RoundedPath 4 4 120 120 26
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Point(8, 8)),
    (New-Object System.Drawing.Point(120, 120)),
    [System.Drawing.Color]::FromArgb(255, 47, 140, 255),
    [System.Drawing.Color]::FromArgb(255, 11, 74, 162)
)
$bgPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 11, 74, 162), 4)
$graphics.FillPath($bgBrush, $bgPath)
$graphics.DrawPath($bgPen, $bgPath)

# Three stacked list bars = the grouped object list.
$barPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 7)
$barPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
foreach ($y in @(26, 52, 78)) {
    $barPath = New-RoundedPath 26 $y 66 18 7
    $graphics.DrawPath($barPen, $barPath)
}

# Type-color dots inside each bar (object categories in the panel).
foreach ($dot in @(@(34, 31, 255, 214, 64), @(34, 57, 255, 255, 255), @(34, 83, 255, 173, 51))) {
    $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($dot[0], $dot[1], $dot[2], $dot[3]))
    $graphics.FillEllipse($dotBrush, $dot[0], $dot[1], 8, 8)
}

# Orange lock badge: locking objects is this panel's signature feature.
$lockBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 245, 157, 35))
$lockBody = New-RoundedPath 78 76 34 30 8
$graphics.FillPath($lockBrush, $lockBody)
$shacklePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 6)
$shacklePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$shacklePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawArc($shacklePen, 84, 60, 22, 22, 180, 180)
$keyBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$graphics.FillEllipse($keyBrush, 91, 86, 8, 8)

$out = New-Object System.Drawing.Bitmap(32, 32, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$outGraphics = [System.Drawing.Graphics]::FromImage($out)
$outGraphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$outGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$outGraphics.DrawImage($bitmap, 0, 0, 32, 32)
$out.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$outGraphics.Dispose()
$out.Dispose()
$graphics.Dispose()
$bitmap.Dispose()
Write-Output "wrote $OutputPath"
