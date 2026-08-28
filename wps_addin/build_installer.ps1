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
$rcedit = Join-Path $projectRoot 'node_modules\rcedit\bin\rcedit-x64.exe'
$installerIcon = Join-Path $projectRoot 'installer_icon.ico'
foreach ($path in @($sevenZip, $sfx)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Missing packaging dependency: $path" }
}

$stage = Join-Path $env:TEMP ("picture-replace-wps-installer-" + [guid]::NewGuid().ToString('N'))
$payload = Join-Path $stage 'payload'
New-Item -ItemType Directory -Path $payload -Force | Out-Null
$pluginPath = Join-Path $payload $pluginFolder
New-Item -ItemType Directory -Path $pluginPath -Force | Out-Null

$pluginFiles = @('index.html', 'main.js', 'ribbon.xml', 'taskpane.html', 'dialog_progress.html', 'README.md', 'icon.png', 'icon_file.png', 'icon_file_all.png', 'icon_clipboard.png', 'icon_clipboard_all.png', 'icon_info.png', 'icon_update.png', 'icon_filter.png', 'icon_smart_zoom.png', 'icon_github.png', 'icon_design_style.png', 'icon_design_text.png', 'icon_design_layout.png', 'icon_design_cleanup.png', 'icon_design_export.png', 'icon_design_color.png', 'icon_design_photoshop.png')
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

# Embed a custom icon into the SFX stub (optional: falls back to the
# default 7-Zip icon when rcedit or the icon file is unavailable).
$iconSfx = Join-Path $stage '7zsd-icon.sfx'
Copy-Item -LiteralPath $sfx -Destination $iconSfx -Force
if ((Test-Path -LiteralPath $rcedit) -and (Test-Path -LiteralPath $installerIcon)) {
    & $rcedit $iconSfx --set-icon $installerIcon --set-version-string "FileDescription" "Picture Replace Tools WPS Installer" --set-version-string "ProductName" "Picture Replace Tools WPS" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "rcedit failed with exit code $LASTEXITCODE; using default SFX icon."
        Copy-Item -LiteralPath $sfx -Destination $iconSfx -Force
    }
} else {
    Write-Warning "rcedit or installer_icon.ico not found; using default SFX icon."
}

$output = [IO.File]::Open($OutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
try {
    foreach ($path in @($iconSfx, $config, $archive)) {
        $input = [IO.File]::OpenRead($path)
        try { $input.CopyTo($output) } finally { $input.Dispose() }
    }
} finally { $output.Dispose() }

Write-Host "Created one-click installer: $OutputPath"
