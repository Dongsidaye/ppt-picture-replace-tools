# Canonical end-to-end test launcher.
# Generates deterministic bitmap fixtures, then uses VBScript for PowerPoint COM
# because PowerShell cannot reliably bind Application.Run's optional arguments.

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$assets = Join-Path $root 'test_assets'
New-Item -ItemType Directory -Force $assets | Out-Null

Add-Type -AssemblyName System.Drawing
function New-SolidPng([string]$path, [int]$width, [int]$height, [System.Drawing.Color]$color) {
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try { $graphics.Clear($color) }
        finally { $graphics.Dispose() }
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
}

New-SolidPng (Join-Path $assets 'old_red.png') 400 300 ([System.Drawing.Color]::FromArgb(220, 30, 30))
New-SolidPng (Join-Path $assets 'new_blue.png') 400 300 ([System.Drawing.Color]::FromArgb(30, 30, 220))
New-SolidPng (Join-Path $assets 'new_green_wide.png') 800 300 ([System.Drawing.Color]::FromArgb(30, 180, 60))
New-SolidPng (Join-Path $assets 'bulk_old_yellow.png') 500 350 ([System.Drawing.Color]::FromArgb(235, 190, 20))
New-SolidPng (Join-Path $assets 'bulk_new_magenta.png') 500 350 ([System.Drawing.Color]::FromArgb(210, 30, 190))
New-SolidPng (Join-Path $assets 'bulk_decoy_cyan.png') 500 350 ([System.Drawing.Color]::FromArgb(20, 190, 210))

& cscript.exe //nologo (Join-Path $root 'test_functional.vbs')
if ($LASTEXITCODE -ne 0) { throw "Functional test failed with exit code $LASTEXITCODE" }

# Confirm exported slides contain the replacement colors inside the picture frame.
$expected = @(
    @{ Path = Join-Path $assets 'slide1_after.png'; R = 30; G = 30; B = 220; Label = 'blue replacement' },
    @{ Path = Join-Path $assets 'slide2_after.png'; R = 30; G = 180; B = 60; Label = 'green replacement' },
    @{ Path = Join-Path $assets 'slide5_after.png'; R = 210; G = 30; B = 190; Label = 'bulk replacement slide 5'; X = 213; Y = 153 },
    @{ Path = Join-Path $assets 'slide6_after.png'; R = 210; G = 30; B = 190; Label = 'bulk replacement slide 6'; X = 299; Y = 191 },
    @{ Path = Join-Path $assets 'slide7_after.png'; R = 210; G = 30; B = 190; Label = 'bulk replacement slide 7'; X = 162; Y = 222 }
)
foreach ($item in $expected) {
    $bitmap = New-Object System.Drawing.Bitmap($item.Path)
    $sampleX = if ($item.ContainsKey('X')) { $item.X } else { 288 }
    $sampleY = if ($item.ContainsKey('Y')) { $item.Y } else { 165 }
    try { $pixel = $bitmap.GetPixel($sampleX, $sampleY) }
    finally { $bitmap.Dispose() }
    if ([Math]::Abs($pixel.R - $item.R) -gt 5 -or
        [Math]::Abs($pixel.G - $item.G) -gt 5 -or
        [Math]::Abs($pixel.B - $item.B) -gt 5) {
        throw "$($item.Label) pixel mismatch: R=$($pixel.R) G=$($pixel.G) B=$($pixel.B)"
    }
    Write-Host "PASS $($item.Label): R=$($pixel.R) G=$($pixel.G) B=$($pixel.B)"
}

Write-Host 'END-TO-END TEST: PASS' -ForegroundColor Green
