param(
    [string]$OutputPath = (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'icon_smart_zoom.png')
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
    (New-Object System.Drawing.Point(12, 8)),
    (New-Object System.Drawing.Point(116, 120)),
    [System.Drawing.Color]::FromArgb(255, 47, 140, 255),
    [System.Drawing.Color]::FromArgb(255, 11, 74, 162)
)
$graphics.FillPath($bgBrush, $bgPath)
$graphics.DrawPath((New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 11, 74, 162), 4)), $bgPath)

$whitePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 7)
$whitePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$whitePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$whitePen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$graphics.DrawLines($whitePen, @(
    (New-Object System.Drawing.Point(50, 30)), (New-Object System.Drawing.Point(30, 30)), (New-Object System.Drawing.Point(30, 50))
))
$graphics.DrawLines($whitePen, @(
    (New-Object System.Drawing.Point(78, 30)), (New-Object System.Drawing.Point(100, 30)), (New-Object System.Drawing.Point(100, 50))
))
$graphics.DrawLines($whitePen, @(
    (New-Object System.Drawing.Point(100, 78)), (New-Object System.Drawing.Point(100, 100)), (New-Object System.Drawing.Point(78, 100))
))
$graphics.DrawLines($whitePen, @(
    (New-Object System.Drawing.Point(50, 100)), (New-Object System.Drawing.Point(30, 100)), (New-Object System.Drawing.Point(30, 78))
))

$framePath = New-RoundedPath 44 44 40 40 8
$frameBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(235, 11, 74, 162))
$framePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 6)
$graphics.FillPath($frameBrush, $framePath)
$graphics.DrawPath($framePen, $framePath)
$plusPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 5)
$plusPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$plusPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawLine($plusPen, 56, 64, 72, 64)
$graphics.DrawLine($plusPen, 64, 56, 64, 72)

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

foreach ($resource in @($smallGraphics, $small, $plusPen, $framePen, $frameBrush, $framePath, $whitePen, $bgBrush, $bgPath, $graphics, $bitmap)) {
    if ($resource -and $resource.Dispose) { $resource.Dispose() }
}
Write-Output $OutputPath
