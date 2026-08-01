# generate_icons.ps1 - Draws the Picture Replace Tools WPS ribbon icons (32x32).
# Flat vector-style, drawn at 4x then downscaled for crispness.
Add-Type -AssemblyName System.Drawing

$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$S = 4                       # supersample factor
$W = 128                     # drawing canvas at 4x (32*4)

function New-RoundedRectPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 2 * $r
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function New-Canvas {
    $bmp = New-Object System.Drawing.Bitmap $W, $W, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    return @($bmp, $g)
}

function Draw-Background($g) {
    $rect = New-Object System.Drawing.RectangleF 0, 0, $W, $W
    $path = New-RoundedRectPath 0 0 $W $W (6.5 * $S)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, `
        ([System.Drawing.Color]::FromArgb(255, 84, 143, 255)), `
        ([System.Drawing.Color]::FromArgb(255, 27, 95, 224)), 55.0
    $g.FillPath($brush, $path)
    $brush.Dispose(); $path.Dispose()
}

function Draw-PhotoFrame($g, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $path = New-RoundedRectPath $x $y $w $h $r
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $g.FillPath($brush, $path)
    $brush.Dispose(); $path.Dispose()
}

function Draw-MountainScene($g, [float]$ox, [float]$oy, [float]$scale) {
    # Sun
    $sunB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 201, 60))
    $g.FillEllipse($sunB, $ox + 30*$scale, $oy + 2*$scale, 16*$scale, 16*$scale)
    $sunB.Dispose()
    # Far mountain (light blue)
    $far = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 175, 201, 240))
    $pts = @(
        (New-Object System.Drawing.PointF ($ox + 8*$scale), ($oy + 76*$scale)),
        (New-Object System.Drawing.PointF ($ox + 34*$scale), ($oy + 34*$scale)),
        (New-Object System.Drawing.PointF ($ox + 62*$scale), ($oy + 76*$scale))
    )
    $g.FillPolygon($far, $pts)
    $far.Dispose()
    # Near mountain (blue)
    $near = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 61, 123, 234))
    $pts2 = @(
        (New-Object System.Drawing.PointF ($ox + 28*$scale), ($oy + 76*$scale)),
        (New-Object System.Drawing.PointF ($ox + 52*$scale), ($oy + 40*$scale)),
        (New-Object System.Drawing.PointF ($ox + 78*$scale), ($oy + 76*$scale))
    )
    $g.FillPolygon($near, $pts2)
    $near.Dispose()
}

function Draw-SwapBadge($g, [float]$cx, [float]$cy, [float]$R, [float]$penW, [float]$headLen, [float]$headHalf) {
    # White badge circle
    $b = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $badgeR = $R + 7; $g.FillEllipse($b, $cx - $badgeR, $cy - $badgeR, (2*$badgeR), (2*$badgeR))
    $b.Dispose()

    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 46, 102, 214)), $penW
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $arcRect = New-Object System.Drawing.RectangleF ($cx - $R), ($cy - $R), (2*$R), (2*$R)
    $arrowB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 102, 214))
    foreach ($start in @(270, 90)) {
        $g.DrawArc($pen, $arcRect, $start, 270)
        $endAngle = ($start + 270) % 360
        $rad = [Math]::PI / 180.0
        $px = $cx + $R * [Math]::Cos($endAngle * $rad)
        $py = $cy + $R * [Math]::Sin($endAngle * $rad)
        $tx = -[Math]::Sin($endAngle * $rad)
        $ty = [Math]::Cos($endAngle * $rad)
        $pEx = $px + $tx * $headLen
        $pEy = $py + $ty * $headLen
        $b1x = $px - $tx * $headLen * 0.45 + $ty * $headHalf
        $b1y = $py - $ty * $headLen * 0.45 - $tx * $headHalf
        $b2x = $px - $tx * $headLen * 0.45 - $ty * $headHalf
        $b2y = $py - $ty * $headLen * 0.45 + $tx * $headHalf
        $tri = @(
            (New-Object System.Drawing.PointF $pEx, $pEy),
            (New-Object System.Drawing.PointF $b1x, $b1y),
            (New-Object System.Drawing.PointF $b2x, $b2y)
        )
        $g.FillPolygon($arrowB, $tri)
    }
    $pen.Dispose(); $arrowB.Dispose()
}

function Save-Scaled($bmp, $g, [string]$path) {
    $g.Dispose()
    $final = New-Object System.Drawing.Bitmap 32, 32, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $fg = [System.Drawing.Graphics]::FromImage($final)
    $fg.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $fg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $fg.DrawImage($bmp, (New-Object System.Drawing.Rectangle 0, 0, 32, 32), 0, 0, $W, $W, [System.Drawing.GraphicsUnit]::Pixel)
    $fg.Dispose()
    $final.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $final.Dispose()
    $bmp.Dispose()
    Write-Host "wrote $path"
}

# ---- icon.png: picture manager (photo frame + swap badge) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-PhotoFrame $g (24) (20) (80) (84) (8)
Draw-MountainScene $g (8) (12) 1.0
Draw-SwapBadge $g 100 100 17 4.5 9 4.0
Save-Scaled $c[0] $g (Join-Path $outDir 'icon.png')

# ---- icon_file.png: replace from file (photo frame + up arrow into frame) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-PhotoFrame $g (24) (24) (80) (80) (8)
Draw-MountainScene $g (12) (18) 0.85
# Up-right arrow from bottom-left corner into the frame
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 9
$pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$g.DrawLine($pen, 26, 102, 66, 62)
$arrowB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$g.FillPolygon($arrowB, @(
    (New-Object System.Drawing.PointF 76, 52),
    (New-Object System.Drawing.PointF 60, 56),
    (New-Object System.Drawing.PointF 64, 72)
))
$pen.Dispose(); $arrowB.Dispose()
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_file.png')

# ---- icon_clipboard.png: replace from clipboard ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
# Clipboard body
$body = New-RoundedRectPath 40 30 48 68 8
$b2 = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$g.FillPath($b2, $body)
# Clipboard top tab
$tab = New-RoundedRectPath 46 18 36 24 6
$g.FillPath($b2, $tab)
$b2.Dispose()
# Clip line
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 84, 143, 255)), 5
$g.DrawLine($pen, 58, 26, 70, 26)
$pen.Dispose()
# Mini photo inside clipboard
Draw-MountainScene $g (26) (30) 0.62
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_clipboard.png')

# ---- icon_info.png: compatibility diagnostics ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
$b = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$g.FillEllipse($b, 34, 34, 60, 60)
$b.Dispose()
$blue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 102, 214))
$g.FillEllipse($blue, 59, 48, 10, 10)                       # dot
$stem = New-RoundedRectPath 60 62 8 22 4
$g.FillPath($blue, $stem)
$blue.Dispose(); $stem.Dispose()
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_info.png')

Write-Host "All icons generated."
