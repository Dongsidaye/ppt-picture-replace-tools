# generate_icons.ps1 - Draws 6 distinct 32x32 ribbon icons for Picture Replace Tools (WPS).
# Flat vector-style, drawn at 4x then downscaled for crispness.
Add-Type -AssemblyName System.Drawing

$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$S = 4
$W = 128

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

function Draw-RoundedFill($g, [float]$x, [float]$y, [float]$w, [float]$h, [float]$r, $color) {
    $path = New-RoundedRectPath $x $y $w $h $r
    $b = New-Object System.Drawing.SolidBrush $color
    $g.FillPath($b, $path)
    $b.Dispose(); $path.Dispose()
}

function Draw-MountainScene($g, [float]$ox, [float]$oy, [float]$scale) {
    $sunB = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 201, 60))
    $g.FillEllipse($sunB, $ox + 30*$scale, $oy + 2*$scale, 16*$scale, 16*$scale)
    $sunB.Dispose()
    $far = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 175, 201, 240))
    $pts = @(
        (New-Object System.Drawing.PointF ($ox + 8*$scale), ($oy + 76*$scale)),
        (New-Object System.Drawing.PointF ($ox + 34*$scale), ($oy + 34*$scale)),
        (New-Object System.Drawing.PointF ($ox + 62*$scale), ($oy + 76*$scale))
    )
    $g.FillPolygon($far, $pts)
    $far.Dispose()
    $near = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 61, 123, 234))
    $pts2 = @(
        (New-Object System.Drawing.PointF ($ox + 28*$scale), ($oy + 76*$scale)),
        (New-Object System.Drawing.PointF ($ox + 52*$scale), ($oy + 40*$scale)),
        (New-Object System.Drawing.PointF ($ox + 78*$scale), ($oy + 76*$scale))
    )
    $g.FillPolygon($near, $pts2)
    $near.Dispose()
}

function Draw-Arrow($g, [float]$x1, [float]$y1, [float]$x2, [float]$y2, [float]$width, $color, [float]$headLen, [float]$headHalf) {
    $pen = New-Object System.Drawing.Pen $color, $width
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($pen, $x1, $y1, $x2, $y2)
    $pen.Dispose()
    $dx = $x2 - $x1; $dy = $y2 - $y1
    $len = [Math]::Sqrt($dx*$dx + $dy*$dy)
    if ($len -le 0.001) { return }
    $ux = $dx / $len; $uy = $dy / $len
    $px = -$uy; $py = $ux
    $tipX = $x2 + $ux * $headLen; $tipY = $y2 + $uy * $headLen
    $baseX = $x2 - $ux * $headLen * 0.45; $baseY = $y2 - $uy * $headLen * 0.45
    $b = New-Object System.Drawing.SolidBrush $color
    $tri = @(
        (New-Object System.Drawing.PointF $tipX, $tipY),
        (New-Object System.Drawing.PointF ($baseX + $px*$headHalf), ($baseY + $py*$headHalf)),
        (New-Object System.Drawing.PointF ($baseX - $px*$headHalf), ($baseY - $py*$headHalf))
    )
    $g.FillPolygon($b, $tri)
    $b.Dispose()
}

function Draw-SwapBadge($g, [float]$cx, [float]$cy, [float]$R, [float]$penW, [float]$headLen, [float]$headHalf) {
    $badgeR = $R + 7
    $b = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $g.FillEllipse($b, $cx - $badgeR, $cy - $badgeR, (2*$badgeR), (2*$badgeR))
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
        $g.FillPolygon($arrowB, @(
            (New-Object System.Drawing.PointF $pEx, $pEy),
            (New-Object System.Drawing.PointF $b1x, $b1y),
            (New-Object System.Drawing.PointF $b2x, $b2y)
        ))
    }
    $pen.Dispose(); $arrowB.Dispose()
}

function Draw-Clipboard($g, [float]$x, [float]$y, [float]$w, [float]$h, $bodyColor, [float]$tabH) {
    # body
    Draw-RoundedFill $g $x ($y + $tabH - 6) $w ($h - $tabH + 6) 8 $bodyColor
    # tab
    Draw-RoundedFill $g ($x + 6) $y ($w - 12) ($tabH + 6) 6 $bodyColor
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

# ---- 1) icon.png: picture manager panel (photo frame + swap badge) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-RoundedFill $g 24 20 80 84 8 ([System.Drawing.Color]::White)
Draw-MountainScene $g 8 12 1.0
Draw-SwapBadge $g 100 100 17 4.5 9 4.0
Save-Scaled $c[0] $g (Join-Path $outDir 'icon.png')

# ---- 2) icon_file.png: replace from file (frame + up arrow into frame) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-RoundedFill $g 24 24 80 80 8 ([System.Drawing.Color]::White)
Draw-MountainScene $g 12 18 0.85
Draw-Arrow $g 26 102 66 62 9 ([System.Drawing.Color]::White) 14 7
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_file.png')

# ---- 3) icon_file_all.png: batch replace from file (stacked frames + swap badge) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-RoundedFill $g 34 14 70 70 8 ([System.Drawing.Color]::FromArgb(255, 199, 219, 255))
Draw-RoundedFill $g 24 34 70 70 8 ([System.Drawing.Color]::White)
Draw-MountainScene $g 18 44 0.70
Draw-SwapBadge $g 98 98 12 4 7 3.5
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_file_all.png')

# ---- 4) icon_clipboard.png: replace from clipboard (single clipboard + photo) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
Draw-Clipboard $g 40 24 48 70 ([System.Drawing.Color]::White) 18
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 84, 143, 255)), 5
$g.DrawLine($pen, 58, 26, 70, 26)
$pen.Dispose()
Draw-MountainScene $g 26 32 0.62
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_clipboard.png')

# ---- 5) icon_clipboard_all.png: batch replace from clipboard (stacked sheets) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
# two sheets peeking behind (multi-document glyph)
Draw-RoundedFill $g 52 34 48 64 8 ([System.Drawing.Color]::FromArgb(255, 199, 219, 255))
Draw-RoundedFill $g 60 26 48 64 8 ([System.Drawing.Color]::FromArgb(255, 224, 236, 255))
# front clipboard
Draw-Clipboard $g 24 42 56 72 ([System.Drawing.Color]::White) 18
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 84, 143, 255)), 4
$g.DrawLine($pen, 40, 44, 50, 44)
$pen.Dispose()
Draw-MountainScene $g 12 50 0.52
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_clipboard_all.png')

# ---- 6) icon_info.png: compatibility diagnostics (info circle) ----
$c = New-Canvas; $g = $c[1]
Draw-Background $g
$b = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$g.FillEllipse($b, 34, 34, 60, 60)
$b.Dispose()
$blue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 46, 102, 214))
$g.FillEllipse($blue, 59, 48, 10, 10)
Draw-RoundedFill $g 60 62 8 22 4 ([System.Drawing.Color]::FromArgb(255, 46, 102, 214))
$blue.Dispose()
Save-Scaled $c[0] $g (Join-Path $outDir 'icon_info.png')

Write-Host "All 6 icons generated."
