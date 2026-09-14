Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

function New-RoundedRectPath([float]$x,[float]$y,[float]$w,[float]$h,[float]$r) {
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = [Math]::Min($r * 2, [Math]::Min($w, $h))
  $p.AddArc($x, $y, $d, $d, 180, 90)
  $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $p.CloseFigure()
  return $p
}

function New-ArrowHead([float]$cx,[float]$cy,[float]$r,[float]$deg,[float]$len) {
  # triangle arrowhead at angle deg on circle (cx,cy,r), pointing along tangent (deg+90)
  $rad = $deg * [Math]::PI / 180.0
  $tipX = $cx + $r * [Math]::Cos($rad)
  $tipY = $cy + $r * [Math]::Sin($rad)
  $back1 = ($deg + 180 - 28) * [Math]::PI / 180.0
  $back2 = ($deg + 180 + 28) * [Math]::PI / 180.0
  $p1 = New-Object System.Drawing.PointF (($tipX + $len * [Math]::Cos($back1)), ($tipY + $len * [Math]::Sin($back1)))
  $p2 = New-Object System.Drawing.PointF (($tipX + $len * [Math]::Cos($back2)), ($tipY + $len * [Math]::Sin($back2)))
  $pts = [System.Drawing.PointF[]]@( (New-Object System.Drawing.PointF $tipX,$tipY), $p1, $p2 )
  return $pts
}

function Draw-SwapIcon([int]$px) {
  $scale = $px / 256.0
  $bmp = New-Object System.Drawing.Bitmap $px, $px
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.Clear([System.Drawing.Color]::Transparent)

  function S([float]$v) { return $v * $scale }

  # background rounded square with gradient
  $bgPath = New-RoundedRectPath (S 12) (S 12) (S 232) (S 232) (S 54)
  $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Point 0,0), (New-Object System.Drawing.Point $px,$px),
    [System.Drawing.Color]::FromArgb(255,58,135,255), [System.Drawing.Color]::FromArgb(255,16,76,198))
  $g.FillPath($bgBrush, $bgPath)

  # subtle top-light overlay
  $over = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Point 0,0), (New-Object System.Drawing.Point 0,$px),
    [System.Drawing.Color]::FromArgb(55,255,255,255), [System.Drawing.Color]::FromArgb(0,255,255,255))
  $g.FillPath($over, $bgPath)

  # white picture frame
  $framePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White, (S 10))
  $framePath = New-RoundedRectPath (S 56) (S 42) (S 144) (S 116) (S 20)
  $g.DrawPath($framePen, $framePath)

  # scene clipped inside frame
  $inner = New-RoundedRectPath (S 66) (S 52) (S 124) (S 96) (S 8)
  $g.SetClip($inner)
  $skyBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,222,241,255))
  $g.FillRectangle($skyBrush, (S 66), (S 52), (S 124), (S 54))
  $groundBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,148,190,244))
  $g.FillRectangle($groundBrush, (S 66), (S 104), (S 124), (S 44))
  # sun
  $sunBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,255,208,84))
  $g.FillEllipse($sunBrush, (S 96), (S 64), (S 28), (S 28))
  # mountains
  $m1 = [System.Drawing.PointF[]]@(
    (New-Object System.Drawing.PointF (S 62),(S 148)),
    (New-Object System.Drawing.PointF (S 106),(S 84)),
    (New-Object System.Drawing.PointF (S 150),(S 148)))
  $m1Brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,78,140,232))
  $g.FillPolygon($m1Brush, $m1)
  $m2 = [System.Drawing.PointF[]]@(
    (New-Object System.Drawing.PointF (S 104),(S 148)),
    (New-Object System.Drawing.PointF (S 144),(S 92)),
    (New-Object System.Drawing.PointF (S 192),(S 148)))
  $m2Brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,127,184,255))
  $g.FillPolygon($m2Brush, $m2)
  $g.ResetClip()

  # swap badge (green circle with white circular arrows)
  $cx = 196.0; $cy = 182.0; $br = 29.0
  $badgeBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255,34,197,94))
  $g.FillEllipse($badgeBrush, (S ($cx - $br)), (S ($cy - $br)), (S ($br * 2)), (S ($br * 2)))
  $ringPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White, (S 5))
  # arc 1
  $r1 = 15.5
  $a1Start = -115.0; $a1Sweep = 215.0
  $g.DrawArc($ringPen, (S ($cx - $r1)), (S ($cy - $r1)), (S ($r1 * 2)), (S ($r1 * 2)), $a1Start, $a1Sweep)
  $h1 = New-ArrowHead $cx $cy $r1 ($a1Start + $a1Sweep) (S 7)
  $g.FillPolygon($ringPen.Brush, $h1)
  # arc 2 (reverse, smaller)
  $r2 = 8.5
  $a2Start = 70.0; $a2Sweep = 200.0
  $g.DrawArc($ringPen, (S ($cx - $r2)), (S ($cy - $r2)), (S ($r2 * 2)), (S ($r2 * 2)), $a2Start, $a2Sweep)
  $h2 = New-ArrowHead $cx $cy $r2 ($a2Start + $a2Sweep) (S 6)
  $g.FillPolygon($ringPen.Brush, $h2)

  $g.Dispose()
  return $bmp
}

$outDir = 'C:\Users\Administrator\Documents\powerpoint\pic_replace_addin\wps_addin'
$sizes = @(256, 128, 64, 48, 32, 24, 16)
$pngs = @{}
foreach ($sz in $sizes) {
  $bmp = Draw-SwapIcon $sz
  $ms = New-Object System.IO.MemoryStream
  $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  $pngs[$sz] = $ms.ToArray()
  if ($sz -eq 256) {
    [System.IO.File]::WriteAllBytes((Join-Path $outDir 'installer_icon.png'), $pngs[$sz])
  }
  $bmp.Dispose(); $ms.Dispose()
}

# assemble ICO
$count = $sizes.Count
$ms = New-Object System.IO.MemoryStream
$bw = New-Object System.IO.BinaryWriter($ms)
$bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]$count)
$offset = 6 + 16 * $count
foreach ($sz in $sizes) {
  $b = if ($sz -ge 256) { 0 } else { $sz }
  $bw.Write([Byte]$b); $bw.Write([Byte]$b)
  $bw.Write([Byte]0); $bw.Write([Byte]0)
  $bw.Write([UInt16]1); $bw.Write([UInt16]32)
  $bw.Write([UInt32]$pngs[$sz].Length); $bw.Write([UInt32]$offset)
  $offset += $pngs[$sz].Length
}
foreach ($sz in $sizes) { $bw.Write($pngs[$sz]) }
$bw.Flush()
[System.IO.File]::WriteAllBytes((Join-Path $outDir 'installer_icon.ico'), $ms.ToArray())
$bw.Dispose(); $ms.Dispose()
Write-Output ("ico written: " + (Join-Path $outDir 'installer_icon.ico') + " bytes=" + (Get-Item (Join-Path $outDir 'installer_icon.ico')).Length)
