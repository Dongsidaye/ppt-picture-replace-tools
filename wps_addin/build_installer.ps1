param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$package = Get-Content (Join-Path $projectRoot 'package.json') -Raw | ConvertFrom-Json
$pluginFolder = "$($package.name)_$($package.version)"
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Split-Path -Parent $projectRoot) "PictureReplaceTools-WPS-$($package.version).exe"
}
if (Test-Path -LiteralPath $OutputPath) { throw "Refusing to overwrite existing output: $OutputPath" }

$sevenZip = Join-Path $projectRoot 'node_modules\7zip-bin\win\x64\7za.exe'
$sfx = Join-Path $projectRoot 'node_modules\wpsjs\src\lib\res\7zsd.sfx'
foreach ($path in @($sevenZip, $sfx)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing packaging dependency: $path" }
}

$stage = Join-Path $env:TEMP ("picture-replace-wps-installer-" + [guid]::NewGuid().ToString('N'))
$payload = Join-Path $stage 'payload'
New-Item -ItemType Directory -Path $payload -Force | Out-Null
$pluginPath = Join-Path $payload $pluginFolder
New-Item -ItemType Directory -Path $pluginPath -Force | Out-Null

$pluginFiles = @('index.html', 'main.js', 'ribbon.xml', 'taskpane.html', 'README.md', 'icon.png', 'icon_file.png', 'icon_clipboard.png', 'icon_info.png')
foreach ($file in $pluginFiles) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $file) -Destination (Join-Path $pluginPath $file)
}
Copy-Item -LiteralPath (Join-Path $projectRoot 'install-wps.ps1') -Destination (Join-Path $payload 'install-wps.ps1')
Copy-Item -LiteralPath (Join-Path $projectRoot 'uninstall-wps.ps1') -Destination (Join-Path $payload 'uninstall-wps.ps1')
Copy-Item -LiteralPath (Join-Path $projectRoot 'install.cmd') -Destination (Join-Path $payload 'install.cmd')

$manifest = @"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="$($package.name)" type="wpp" version="$($package.version)" url="$pluginFolder" enable="enable_dev" install="null" customDomain="" />
</jsplugins>
"@
Set-Content -LiteralPath (Join-Path $payload 'publish.xml') -Value $manifest -Encoding UTF8

$archive = Join-Path $stage 'payload.7z'
Push-Location $payload
try {
    & $sevenZip a -t7z -mx=9 $archive 'install.cmd' 'install-wps.ps1' 'uninstall-wps.ps1' 'publish.xml' $pluginFolder | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "7za failed with exit code $LASTEXITCODE" }
} finally { Pop-Location }

$config = Join-Path $stage 'sfx-config.txt'
$sfxConfig = @"
;!@Install@!UTF-8!
Title="Picture Replace Tools - WPS"
BeginPrompt="Install Picture Replace Tools for WPS Presentation?"
RunProgram="install.cmd"
;!@InstallEnd@!
"@
Set-Content -LiteralPath $config -Value $sfxConfig -Encoding UTF8

$output = [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    foreach ($path in @($sfx, $config, $archive)) {
        $input = [IO.File]::OpenRead($path)
        try { $input.CopyTo($output) } finally { $input.Dispose() }
    }
} finally { $output.Dispose() }

Write-Host "Created one-click installer: $OutputPath"
