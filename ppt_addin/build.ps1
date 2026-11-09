param([switch]$Package, [switch]$RefreshShared, [string]$OutputDirectory = '')
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$meta = Get-Content -LiteralPath (Join-Path $root 'package.json') -Raw | ConvertFrom-Json
if (!$OutputDirectory) { $OutputDirectory = Join-Path $root ('build\' + $meta.version) }
$out = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $out -Force | Out-Null
$sdk = Join-Path $env:LOCALAPPDATA ('PictureReplaceToolsBuild\WebView2-' + $meta.webView2Version)
New-Item -ItemType Directory -Path $sdk -Force | Out-Null
$archive = Join-Path $sdk 'sdk.zip'
if (!(Test-Path -LiteralPath $archive)) {
    $url = 'https://api.nuget.org/v3-flatcontainer/microsoft.web.webview2/{0}/microsoft.web.webview2.{0}.nupkg' -f $meta.webView2Version
    Invoke-WebRequest -Uri $url -OutFile $archive
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $meta.webView2Sha256) { throw 'WebView2 SDK SHA-256 mismatch.' }
Expand-Archive -LiteralPath $archive -DestinationPath $sdk -Force
if ($RefreshShared) {
    & python (Join-Path $root 'prepare_shared.py')
    if ($LASTEXITCODE -ne 0) { throw 'Shared engine preparation failed.' }
}
if (!(Test-Path -LiteralPath (Join-Path $root 'shared\source-manifest.json'))) { throw 'Missing shared snapshot. Run with -RefreshShared from the matching WPS source.' }
$ui = Join-Path $out 'ui'
New-Item -ItemType Directory -Path $ui -Force | Out-Null
Copy-Item -Path (Join-Path $root 'shared\*') -Destination $ui -Force
Copy-Item -Path (Join-Path $root 'web\*.js') -Destination $ui -Force
Copy-Item -LiteralPath (Join-Path $root 'tests\selftest.js') -Destination (Join-Path $ui 'selftest.js') -Force
$core = Join-Path $sdk 'lib\net462\Microsoft.Web.WebView2.Core.dll'
$forms = Join-Path $sdk 'lib\net462\Microsoft.Web.WebView2.WinForms.dll'
Copy-Item -LiteralPath $core,$forms -Destination $out -Force
foreach ($arch in @('x64', 'x86')) {
    $target = Join-Path $out $arch
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $sdk "runtimes\win-$arch\native\WebView2Loader.dll") -Destination $target -Force
}
$piaRoot = Join-Path $env:windir 'assembly\GAC_MSIL'
$office = Get-ChildItem -Path (Join-Path $piaRoot 'office\*\office.dll') | Select-Object -First 1 -ExpandProperty FullName
$ppt = Get-ChildItem -Path (Join-Path $piaRoot 'Microsoft.Office.Interop.PowerPoint\*\Microsoft.Office.Interop.PowerPoint.dll') | Select-Object -First 1 -ExpandProperty FullName
if (!$office -or !$ppt) { throw 'Office interop assemblies were not found. Install the Office .NET programmability support.' }
Copy-Item -LiteralPath $office,$ppt -Destination $out -Force
$csc = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$sources = @(Get-ChildItem -LiteralPath $root -Filter '*.cs' | ForEach-Object FullName)
$revision = [int]([DateTime]::UtcNow.TimeOfDay.TotalSeconds / 2)
$assemblyInfo = Join-Path $out 'AssemblyInfo.cs'
[IO.File]::WriteAllText($assemblyInfo, ('[assembly:System.Reflection.AssemblyVersion("' + $meta.version + '.' + $revision + '")]' + "`r`n" + '[assembly:System.Reflection.AssemblyFileVersion("' + $meta.version + '.0")]'))
$sources += $assemblyInfo
$references = @('System.dll','System.Core.dll','System.Windows.Forms.dll','System.Drawing.dll','System.Web.Extensions.dll','System.Xml.dll','System.Xml.Linq.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','Microsoft.CSharp.dll',$core,$forms,$office,$ppt)
$options = @('/nologo','/target:library','/platform:anycpu','/langversion:5','/codepage:65001','/optimize+',('/out:' + (Join-Path $out 'PowerPointTools.dll')))
foreach ($ref in $references) { $options += '/reference:' + $ref }
& $csc @options @sources
if ($LASTEXITCODE -ne 0) { throw 'PowerPoint native host compilation failed.' }
Copy-Item -LiteralPath (Join-Path $root 'package.json'),(Join-Path $root 'install.ps1'),(Join-Path $root 'uninstall.ps1'),(Join-Path $root 'install.cmd'),(Join-Path $root 'registration.json') -Destination $out -Force
foreach ($scriptName in @('install.ps1','uninstall.ps1')) {
    $scriptText = [IO.File]::ReadAllText((Join-Path $root $scriptName), [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText((Join-Path $out $scriptName), $scriptText, (New-Object Text.UTF8Encoding($true)))
}
foreach ($docName in @('README.md','PARITY.md','VALIDATION.md')) { Copy-Item -LiteralPath (Join-Path $root $docName) -Destination $out -Force }
Copy-Item -LiteralPath (Join-Path $root '..\LICENSE') -Destination $out -Force
$hashes = [ordered]@{}
Get-ChildItem -LiteralPath $out -Recurse -File | Where-Object Name -ne 'checksums.json' | ForEach-Object {
    $rel = $_.FullName.Substring($out.Length + 1).Replace('\','/')
    $hashes[$rel] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
}
$hashes | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $out 'checksums.json') -Encoding UTF8
Write-Host "Built: $out"
if ($Package) {
    $wps = Join-Path $root '..\wps_addin'
    $seven = Join-Path $wps 'node_modules\7zip-bin\win\x64\7za.exe'
    $sfx = Join-Path $wps 'node_modules\wpsjs\src\lib\res\7zsd.sfx'
    if (!(Test-Path -LiteralPath $seven) -or !(Test-Path -LiteralPath $sfx)) { throw 'Packaging tools are missing. Run npm ci in wps_addin, or build without -Package.' }
    $stage = Join-Path $env:TEMP ('ppt-tools-package-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    $sevenArchive = Join-Path $stage 'payload.7z'
    Push-Location $out
    try { & $seven a -t7z -mx=9 $sevenArchive '*' | Out-Null; if ($LASTEXITCODE) { throw '7zip failed.' } } finally { Pop-Location }
    $config = Join-Path $stage 'sfx-config.txt'
    $configText = ";!@Install@!UTF-8!`r`nTitle=`"PowerPoint - Dongsidaye Toolbox`"`r`nRunProgram=`"install.cmd`"`r`n;!@InstallEnd@!`r`n"
    [IO.File]::WriteAllText($config, $configText, (New-Object Text.UTF8Encoding($false)))
    $installer = Join-Path $root ("PictureReplaceTools-PPT-$($meta.version).exe")
    $stream = [IO.File]::Open($installer, [IO.FileMode]::Create)
    try {
        foreach ($part in @($sfx,$config,$sevenArchive)) {
            $inputStream = [IO.File]::OpenRead($part)
            try { $inputStream.CopyTo($stream) } finally { $inputStream.Dispose() }
        }
    } finally { $stream.Dispose() }
    $release = [ordered]@{ version=$meta.version; platform='PowerPoint'; downloadUrl="https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/download/ppt-v$($meta.version)/PictureReplaceTools-PPT-$($meta.version).exe"; sha256=(Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant() }
    $release | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'release.json') -Encoding UTF8
    Write-Host "Installer: $installer"
}
