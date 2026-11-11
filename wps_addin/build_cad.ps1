param(
    [string]$OutputDirectory = '',
    [switch]$SkipConverter
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$pptRoot = [IO.Path]::GetFullPath((Join-Path $root '..\ppt_addin'))
if (!$OutputDirectory) { $OutputDirectory = Join-Path $root 'build\cad' }
$output = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null

if (!$SkipConverter) {
    & dotnet publish (Join-Path $pptRoot 'cad_converter\DwgToSvg.csproj') -c Release -r win-x64 --self-contained true -o $output -p:RestoreLockedMode=true -p:PublishSingleFile=true -p:DebugType=none
    if ($LASTEXITCODE -ne 0) { throw 'DWG converter publication failed.' }
    $nativeRoot = Join-Path $pptRoot 'cad_converter\bin\Release\net8.0\win-x64'
    foreach ($nativeName in @('libSkiaSharp.dll', 'libHarfBuzzSharp.dll', 'onnxruntime.dll', 'onnxruntime_providers_shared.dll')) {
        $source = Join-Path $nativeRoot $nativeName
        if (!(Test-Path -LiteralPath $source -PathType Leaf)) { throw ('Missing CAD native runtime: ' + $nativeName) }
        Copy-Item -LiteralPath $source -Destination (Join-Path $output $nativeName) -Force
    }
    & (Join-Path $pptRoot 'cad_converter\Fetch-OcrModels.ps1') -DestinationFolder (Join-Path $output 'models\ocr')
    Copy-Item -LiteralPath (Join-Path $pptRoot 'cad_converter\THIRD_PARTY_NOTICES.md') -Destination $output -Force
    Copy-Item -LiteralPath (Join-Path $pptRoot 'cad_converter\licenses') -Destination $output -Recurse -Force
}
if (!(Test-Path -LiteralPath (Join-Path $output 'DwgToSvg.exe') -PathType Leaf)) { throw 'DwgToSvg.exe is missing.' }

# Reuse the maintained SVG validation and DTO without loading Office interop.
$generated = Join-Path $root 'build\cad-bridge-sources'
New-Item -ItemType Directory -Path $generated -Force | Out-Null
$svgSource = [IO.File]::ReadAllText((Join-Path $pptRoot 'DwgImport.cs'), [Text.Encoding]::UTF8)
$svgMarker = '    internal static class DwgSvg'
$svgStart = $svgSource.IndexOf($svgMarker, [StringComparison]::Ordinal)
if ($svgStart -lt 0 -or $svgSource.LastIndexOf($svgMarker, [StringComparison]::Ordinal) -ne $svgStart) { throw 'SVG validation source drift.' }
$svgHeader = "using System;`nusing System.Collections.Generic;`nusing System.Globalization;`nusing System.IO;`nusing System.Linq;`nusing System.Xml;`nusing System.Xml.Linq;`nnamespace Dongsidaye.PowerPointTools`n{`n"
$svgPath = Join-Path $generated 'DwgSvg.cs'
[IO.File]::WriteAllText($svgPath, $svgHeader + $svgSource.Substring($svgStart), [Text.UTF8Encoding]::new($false))

$textSource = [IO.File]::ReadAllText((Join-Path $pptRoot 'DwgTextInsertion.cs'), [Text.Encoding]::UTF8)
$dtoStart = $textSource.IndexOf('    internal sealed class DwgTextItem', [StringComparison]::Ordinal)
$dtoEnd = $textSource.IndexOf('    internal static class DwgTextInsertion', [StringComparison]::Ordinal)
if ($dtoStart -lt 0 -or $dtoEnd -le $dtoStart) { throw 'DWG text DTO source drift.' }
$dtoPath = Join-Path $generated 'DwgTextItem.cs'
[IO.File]::WriteAllText($dtoPath, "namespace Dongsidaye.PowerPointTools`n{`n" + $textSource.Substring($dtoStart, $dtoEnd - $dtoStart) + "`n}`n", [Text.UTF8Encoding]::new($false))

$compiler = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'cad_bridge') -Filter '*.cs' -File | ForEach-Object FullName)
$sources += @($svgPath, $dtoPath, (Join-Path $pptRoot 'DwgOcrLayout.cs'))
$options = @('/nologo', '/target:winexe', '/platform:anycpu', '/langversion:5', '/codepage:65001', '/optimize+', ('/out:' + (Join-Path $output 'WpsCadBridge.exe')))
foreach ($reference in @('System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Web.Extensions.dll', 'System.Xml.dll', 'System.Xml.Linq.dll')) {
    $options += '/reference:' + $reference
}
& $compiler @options @sources
if ($LASTEXITCODE -ne 0) { throw 'WPS CAD worker compilation failed.' }

foreach ($artifact in @(Get-ChildItem -LiteralPath $output -File | Where-Object { $_.Extension -in @('.pdb', '.lib') })) {
    Remove-Item -LiteralPath $artifact.FullName -Force
}
Write-Host ('Built WPS CAD components: ' + $output)
