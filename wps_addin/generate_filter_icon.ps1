param(
    [string]$OutputPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'icon_filter.png')
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

# An orange selection frame makes the action legible at ribbon size.  The
# white corner ticks remain visible when WPS renders the 32px asset on a dark
# or light ribbon theme.
$framePath = New-RoundedPath 25 24 68 68 8
$frameBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 245, 157, 35))
$framePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 5)
$framePen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$graphics.FillPath($frameBrush, $framePath)
$graphics.DrawPath($framePen, $framePath)

$cornerPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 255, 244, 204), 4)
$cornerPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$cornerPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawLine($cornerPen, 30, 43, 30, 31)
$graphics.DrawLine($cornerPen, 30, 31, 42, 31)
$graphics.DrawLine($cornerPen, 76, 31, 88, 31)
$graphics.DrawLine($cornerPen, 88, 31, 88, 43)
$graphics.DrawLine($cornerPen, 30, 73, 30, 85)
$graphics.DrawLine($cornerPen, 30, 85, 42, 85)
$graphics.DrawLine($cornerPen, 76, 85, 88, 85)
$graphics.DrawLine($cornerPen, 88, 85, 88, 73)

# Small pointer overlay communicates “select objects” without competing with
# the orange frame.  A dark outline preserves contrast against the frame.
$pointer = New-Object System.Drawing.Drawing2D.GraphicsPath
$pointer.AddPolygon(@(
    (New-Object System.Drawing.Point(73, 67)),
    (New-Object System.Drawing.Point(105, 91)),
    (New-Object System.Drawing.Point(91, 94)),
    (New-Object System.Drawing.Point(86, 108)),
    (New-Object System.Drawing.Point(78, 105)),
    (New-Object System.Drawing.Point(83, 92)),
    (New-Object System.Drawing.Point(73, 67))
))
$pointerOutline = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, 7, 36, 75), 5)
$pointerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$graphics.FillPath($pointerBrush, $pointer)
$graphics.DrawPath($pointerOutline, $pointer)

$small = New-Object System.Drawing.Bitmap(32, 32, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$smallGraphics = [System.Drawing.Graphics]::FromImage($small)
$smallGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$smallGraphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$smallGraphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$smallGraphics.Clear([System.Drawing.Color]::Transparent)
$smallGraphics.DrawImage($bitmap, 0, 0, 32, 32)
$directory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$small.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)

foreach ($resource in @($smallGraphics, $small, $pointerOutline, $pointerBrush, $pointer, $cornerPen, $framePen, $frameBrush, $framePath, $bgPen, $bgBrush, $bgPath, $graphics, $bitmap)) {
    if ($resource -and $resource.Dispose) { $resource.Dispose() }
}
Write-Output $OutputPath
