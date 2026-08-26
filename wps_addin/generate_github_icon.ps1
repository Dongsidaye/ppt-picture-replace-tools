# Generate the GitHub project button icon used by the WPS Ribbon.
# The mark is intentionally drawn locally so the add-in never depends on a
# network request or an icon font at Ribbon load time.
Add-Type -AssemblyName System.Drawing

$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scale = 4
$canvasSize = 128

function New-RoundedRectPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 2 * $r
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $path.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-Point([float]$x, [float]$y) {
    return New-Object System.Drawing.PointF $x, $y
}

$bmp = New-Object System.Drawing.Bitmap $canvasSize, $canvasSize, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.Clear([System.Drawing.Color]::Transparent)

# GitHub's dark neutral palette gives the icon a distinct identity beside the
# add-in's blue utility icons. The blue outline keeps it legible on light WPS
# Ribbon backgrounds.
$bgRect = New-Object System.Drawing.RectangleF 0, 0, $canvasSize, $canvasSize
$bgPath = New-RoundedRectPath 2 2 124 124 24
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $bgRect, `
    ([System.Drawing.Color]::FromArgb(255, 55, 64, 72)), `
    ([System.Drawing.Color]::FromArgb(255, 13, 17, 23)), 135.0
$g.FillPath($bgBrush, $bgPath)
$bgBrush.Dispose(); $bgPath.Dispose()

$ringPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(230, 84, 174, 255)), 3
$ringPath = New-RoundedRectPath 7 7 114 114 18
$g.DrawPath($ringPen, $ringPath)
$ringPath.Dispose(); $ringPen.Dispose()

$white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
# Head and ears: a compact Octocat silhouette that remains recognizable at
# the 32px Ribbon size.
$g.FillEllipse($white, 29, 24, 70, 70)
$g.FillPolygon($white, @(
    (New-Point 31 43), (New-Point 28 15), (New-Point 53 31),
    (New-Point 75 31), (New-Point 100 15), (New-Point 97 43)
))
$g.FillEllipse($white, 37, 61, 54, 57)

# Tail, arms and feet complete the mark without using a font or external SVG
# renderer (both of which vary between WPS builds).
$tailPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 8
$tailPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$tailPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$tail = New-Object System.Drawing.Drawing2D.GraphicsPath
$tail.AddBezier(86, 89, 108, 99, 112, 80, 103, 72)
$g.DrawPath($tailPen, $tail)
$tail.Dispose(); $tailPen.Dispose()

$limbPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 7
$limbPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$limbPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$g.DrawLine($limbPen, 48, 83, 40, 104)
$g.DrawLine($limbPen, 80, 83, 88, 104)
$limbPen.Dispose()

# Face cut-outs use the same dark tone as the background, keeping the mark
# crisp when WPS scales the PNG down.
$face = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 36, 41, 47))
$g.FillEllipse($face, 46, 49, 8, 10)
$g.FillEllipse($face, 74, 49, 8, 10)
$g.FillEllipse($face, 59, 62, 10, 7)
$smilePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 36, 41, 47)), 2
$smilePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$smilePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$g.DrawArc($smilePen, 54, 59, 20, 18, 20, 140)
$smilePen.Dispose(); $face.Dispose(); $white.Dispose()

$g.Dispose()
$final = New-Object System.Drawing.Bitmap 32, 32, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$fg = [System.Drawing.Graphics]::FromImage($final)
$fg.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$fg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$fg.DrawImage($bmp, (New-Object System.Drawing.Rectangle 0, 0, 32, 32), 0, 0, $canvasSize, $canvasSize, [System.Drawing.GraphicsUnit]::Pixel)
$fg.Dispose()
$pngPath = Join-Path $outDir 'icon_github.png'
$final.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
$final.Dispose(); $bmp.Dispose()
Write-Host "wrote $pngPath"
