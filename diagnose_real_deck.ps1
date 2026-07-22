param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath
)

$ErrorActionPreference = 'Stop'
$outputDir = Join-Path $env:TEMP ("PicReplace_real_diag_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $outputDir | Out-Null

$powerPoint = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $powerPoint.Presentations.Open($PresentationPath, -1, 0, 0)
    for ($slideIndex = 1; $slideIndex -le [Math]::Min(4, $presentation.Slides.Count); $slideIndex++) {
        $source = $presentation.Slides.Item($slideIndex).Shapes.Item(1)
        $duplicateRange = $source.Duplicate()
        $probe = $duplicateRange.Item(1)
        $crop = $probe.PictureFormat.Crop

        $before = [pscustomobject]@{
            Slide = $slideIndex
            PictureWidth = $crop.PictureWidth
            PictureHeight = $crop.PictureHeight
            ShapeWidth = $crop.ShapeWidth
            ShapeHeight = $crop.ShapeHeight
            OffsetX = $crop.PictureOffsetX
            OffsetY = $crop.PictureOffsetY
        }

        $fullWidth = $crop.PictureWidth
        $fullHeight = $crop.PictureHeight
        $crop.ShapeWidth = $fullWidth
        $crop.ShapeHeight = $fullHeight
        $crop.PictureOffsetX = 0
        $crop.PictureOffsetY = 0
        $probe.Rotation = 0
        $probe.LockAspectRatio = 0
        $probe.Width = 128
        $probe.Height = 128

        $after = [pscustomobject]@{
            Slide = $slideIndex
            PictureWidth = $crop.PictureWidth
            PictureHeight = $crop.PictureHeight
            ShapeWidth = $crop.ShapeWidth
            ShapeHeight = $crop.ShapeHeight
            OffsetX = $crop.PictureOffsetX
            OffsetY = $crop.PictureOffsetY
        }

        $outputPath = Join-Path $outputDir ("slide{0}.png" -f $slideIndex)
        $scaleWidth = [Math]::Round($presentation.PageSetup.SlideWidth * 256 / $probe.Width)
        $scaleHeight = [Math]::Round($presentation.PageSetup.SlideHeight * 256 / $probe.Height)
        $probe.Export($outputPath, 2, $scaleWidth, $scaleHeight, 1)
        $probe.Delete()
        Write-Output ("BEFORE " + ($before | ConvertTo-Json -Compress))
        Write-Output ("AFTER " + ($after | ConvertTo-Json -Compress))
    }
    $presentation.Close()
}
finally {
    $powerPoint.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
}

Get-FileHash -Algorithm SHA256 -LiteralPath (Get-ChildItem -LiteralPath $outputDir -Filter '*.png').FullName |
    ForEach-Object { Write-Output ("PNG " + $_.Path + " " + $_.Hash) }

Add-Type -AssemblyName System.Drawing
foreach ($file in Get-ChildItem -LiteralPath $outputDir -Filter '*.png' | Sort-Object Name) {
    $bitmap = [Drawing.Bitmap]::new($file.FullName)
    try {
        $bytes = New-Object byte[] ($bitmap.Width * $bitmap.Height * 4)
        $offset = 0
        for ($y = 0; $y -lt $bitmap.Height; $y++) {
            for ($x = 0; $x -lt $bitmap.Width; $x++) {
                $color = $bitmap.GetPixel($x, $y)
                $bytes[$offset++] = $color.A
                $bytes[$offset++] = $color.R
                $bytes[$offset++] = $color.G
                $bytes[$offset++] = $color.B
            }
        }
        $sha = [Security.Cryptography.SHA256]::Create()
        $pixelHash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
        Write-Output ("PIXELS {0} {1}x{2} {3}" -f $file.Name, $bitmap.Width, $bitmap.Height, $pixelHash)
    }
    finally {
        $bitmap.Dispose()
    }
}

$files = @(Get-ChildItem -LiteralPath $outputDir -Filter '*.png' | Sort-Object Name)
for ($leftIndex = 0; $leftIndex -lt $files.Count; $leftIndex++) {
    for ($rightIndex = $leftIndex + 1; $rightIndex -lt $files.Count; $rightIndex++) {
        $left = [Drawing.Bitmap]::new($files[$leftIndex].FullName)
        $right = [Drawing.Bitmap]::new($files[$rightIndex].FullName)
        $leftSmall = [Drawing.Bitmap]::new(32, 32)
        $rightSmall = [Drawing.Bitmap]::new(32, 32)
        try {
            $leftGraphics = [Drawing.Graphics]::FromImage($leftSmall)
            $rightGraphics = [Drawing.Graphics]::FromImage($rightSmall)
            try {
                $leftGraphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $rightGraphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $leftGraphics.DrawImage($left, 0, 0, 32, 32)
                $rightGraphics.DrawImage($right, 0, 0, 32, 32)
            }
            finally {
                $leftGraphics.Dispose()
                $rightGraphics.Dispose()
            }
            [long]$sumDifference = 0
            [int]$maximumDifference = 0
            for ($y = 0; $y -lt 32; $y++) {
                for ($x = 0; $x -lt 32; $x++) {
                    $a = $leftSmall.GetPixel($x, $y)
                    $b = $rightSmall.GetPixel($x, $y)
                    $difference = [Math]::Abs($a.R - $b.R) + [Math]::Abs($a.G - $b.G) + [Math]::Abs($a.B - $b.B)
                    $sumDifference += $difference
                    if ($difference -gt $maximumDifference) { $maximumDifference = $difference }
                }
            }
            $meanDifference = $sumDifference / (32 * 32 * 3)
            Write-Output ("DISTANCE {0}-{1} mean={2:N4} maxRGBSum={3}" -f $files[$leftIndex].Name, $files[$rightIndex].Name, $meanDifference, $maximumDifference)
        }
        finally {
            $left.Dispose()
            $right.Dispose()
            $leftSmall.Dispose()
            $rightSmall.Dispose()
        }
    }
}

function Get-WiaCanonicalBytes([string]$path) {
    $image = New-Object -ComObject WIA.ImageFile
    $image.LoadFile($path)
    $process = New-Object -ComObject WIA.ImageProcess
    $process.Filters.Add($process.FilterInfos.Item('Scale').FilterID)
    $process.Filters.Item(1).Properties.Item('MaximumWidth').Value = 32
    $process.Filters.Item(1).Properties.Item('MaximumHeight').Value = 32
    $process.Filters.Item(1).Properties.Item('PreserveAspectRatio').Value = $false
    $process.Filters.Add($process.FilterInfos.Item('Convert').FilterID)
    $process.Filters.Item(2).Properties.Item('FormatID').Value = '{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}'
    $result = $process.Apply($image)
    return [byte[]]$result.FileData.BinaryData
}

$wiaBytes = @()
foreach ($file in $files) { $wiaBytes += ,(Get-WiaCanonicalBytes $file.FullName) }
for ($leftIndex = 0; $leftIndex -lt $files.Count; $leftIndex++) {
    for ($rightIndex = $leftIndex + 1; $rightIndex -lt $files.Count; $rightIndex++) {
        $leftBytes = $wiaBytes[$leftIndex]
        $rightBytes = $wiaBytes[$rightIndex]
        $leftOffset = [BitConverter]::ToInt32($leftBytes, 10)
        $rightOffset = [BitConverter]::ToInt32($rightBytes, 10)
        [long]$sumDifference = 0
        [int]$maximumDifference = 0
        for ($pixel = 0; $pixel -lt 1024; $pixel++) {
            $pixelDifference = 0
            for ($channel = 0; $channel -lt 3; $channel++) {
                $pixelDifference += [Math]::Abs($leftBytes[$leftOffset + $pixel * 4 + $channel] - $rightBytes[$rightOffset + $pixel * 4 + $channel])
            }
            $sumDifference += $pixelDifference
            if ($pixelDifference -gt $maximumDifference) { $maximumDifference = $pixelDifference }
        }
        $meanDifference = $sumDifference / (1024 * 3)
        Write-Output ("WIA_DISTANCE {0}-{1} mean={2:N4} maxRGBSum={3}" -f $files[$leftIndex].Name, $files[$rightIndex].Name, $meanDifference, $maximumDifference)
    }
}

Write-Output ("OUTPUT " + $outputDir)
