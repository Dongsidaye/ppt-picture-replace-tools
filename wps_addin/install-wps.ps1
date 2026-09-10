param(
    [switch]$RestartWps
)

$ErrorActionPreference = 'Stop'

$manifestPath = Join-Path $PSScriptRoot 'publish.xml'
if (-not (Test-Path -LiteralPath $manifestPath)) {
    throw "Missing installer manifest: $manifestPath"
}

[xml]$manifest = Get-Content -LiteralPath $manifestPath -Raw
$plugin = $manifest.SelectSingleNode('/jsplugins/jsplugin')
if ($null -eq $plugin) { throw 'Installer manifest does not contain a jsplugin entry.' }

$pluginName = [string]$plugin.GetAttribute('name')
$relativeFolder = [string]$plugin.GetAttribute('url')
if ([string]::IsNullOrWhiteSpace($pluginName) -or
    [string]::IsNullOrWhiteSpace($relativeFolder) -or
    [IO.Path]::GetFileName($relativeFolder) -ne $relativeFolder -or
    $relativeFolder -notmatch '^[A-Za-z0-9._-]+$') {
    throw 'Installer manifest contains an unsafe plugin name or folder.'
}

$sourceFolder = Join-Path $PSScriptRoot $relativeFolder
if (-not (Test-Path -LiteralPath (Join-Path $sourceFolder 'ribbon.xml'))) {
    throw "Plugin payload is missing ribbon.xml: $sourceFolder"
}

$jsAddinsRoot = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$targetFolder = Join-Path $jsAddinsRoot $relativeFolder
$publishPath = Join-Path $jsAddinsRoot 'publish.xml'

New-Item -ItemType Directory -Path $jsAddinsRoot -Force | Out-Null
# Move previous versions aside before installing the new payload. This is the
# upgrade path used by the one-click installer; it avoids leaving stale JS
# files that WPS could load after a restart.
$backupRoot = Join-Path $jsAddinsRoot 'picture-replace-tools-backups'
$upgradeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
foreach ($oldFolder in @(Get-ChildItem -LiteralPath $jsAddinsRoot -Directory -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like ($pluginName + '_*')
})) {
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    $backupFolder = Join-Path $backupRoot ($oldFolder.Name + '.' + $upgradeStamp)
    Move-Item -LiteralPath $oldFolder.FullName -Destination $backupFolder -Force
}
New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
Copy-Item -Path (Join-Path $sourceFolder '*') -Destination $targetFolder -Recurse -Force

if (Test-Path -LiteralPath $publishPath) {
    $backupPath = Join-Path $jsAddinsRoot ("publish.xml.picture-replace-tools-backup-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Copy-Item -LiteralPath $publishPath -Destination $backupPath -Force
    try { [xml]$current = Get-Content -LiteralPath $publishPath -Raw }
    catch { throw "Existing WPS publish.xml is not valid XML. Backup saved at $backupPath" }
} else {
    $current = New-Object System.Xml.XmlDocument
    $declaration = $current.CreateXmlDeclaration('1.0', 'UTF-8', $null)
    [void]$current.AppendChild($declaration)
    [void]$current.AppendChild($current.CreateElement('jsplugins'))
}

$root = $current.SelectSingleNode('/jsplugins')
if ($null -eq $root) { throw 'Existing WPS publish.xml has no jsplugins root.' }

foreach ($old in @($current.SelectNodes('/jsplugins/jsplugin[@name="' + $pluginName + '"]'))) {
    [void]$root.RemoveChild($old)
}
# Also drop stale debug/online entries for the same add-in so WPS never
# tries a dead local debug server before the offline install.
foreach ($oldOnline in @($current.SelectNodes('/jsplugins/jspluginonline[@name="' + $pluginName + '"]'))) {
    [void]$root.RemoveChild($oldOnline)
}

$newPlugin = $current.CreateElement('jsplugin')
foreach ($attribute in @('name', 'type', 'version', 'url', 'enable', 'install', 'customDomain')) {
    if ($plugin.HasAttribute($attribute)) {
        $newPlugin.SetAttribute($attribute, $plugin.GetAttribute($attribute))
    }
}
[void]$root.AppendChild($current.ImportNode($newPlugin, $true))
$current.Save($publishPath)

# ---------------------------------------------------------------------------
# Keep authaddin.json in sync so WPS loads the newly installed version.
# The md5 field is WPS's add-in auth token; we generate a stable per-version
# token matching the format WPS itself writes (base64(hex(token))).
# ---------------------------------------------------------------------------
function New-AddinAuthToken([string]$name, [string]$version) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($name + '|' + $version))
    } finally {
        $sha.Dispose()
    }
    $prefix = [Convert]::ToBase64String($bytes).Substring(0, 24).Replace('+', '-').Replace('/', '_')
    $token = $prefix + 'WTX75ITas65Un42B_GS-30='
    $hex = -join ($token.ToCharArray() | ForEach-Object { '{0:x2}' -f [int][char]$_ })
    return [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($hex))
}

$authPath = Join-Path $jsAddinsRoot 'authaddin.json'
if (Test-Path -LiteralPath $authPath) {
    try {
        $authBackup = Join-Path $jsAddinsRoot ("authaddin.json.bak-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Copy-Item -LiteralPath $authPath -Destination $authBackup -Force
        $auth = Get-Content -LiteralPath $authPath -Raw | ConvertFrom-Json
        $authChanged = $false
        if ($auth -and $auth.wpp) {
            foreach ($authKey in @($auth.wpp.PSObject.Properties.Name)) {
                $entry = $auth.wpp.$authKey
                if ($entry -is [PSCustomObject] -and $entry.name -eq $pluginName) {
                    $entry.path = $targetFolder.Replace('\', '/')
                    $entry.md5 = New-AddinAuthToken $pluginName ([string]$plugin.GetAttribute('version'))
                    $authChanged = $true
                }
            }
        }
        if ($authChanged) {
            $auth | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $authPath -Encoding UTF8
        }
    } catch {
        Write-Warning "Failed to update authaddin.json: $_"
    }
}

# ---------------------------------------------------------------------------
# One-click update marker: when the add-in launches the installer for an
# auto-update it writes picture_replace_auto_restart.flag in %TEMP% and/or the
# jsaddons root. Honor it by restarting WPS after install.
# ---------------------------------------------------------------------------
$markerTmp = Join-Path $env:TEMP 'picture_replace_auto_restart.flag'
$markerRoot = Join-Path $jsAddinsRoot 'picture_replace_auto_restart.flag'
$autoMarker = (Test-Path -LiteralPath $markerTmp -ErrorAction SilentlyContinue) -or
              (Test-Path -LiteralPath $markerRoot -ErrorAction SilentlyContinue)
if ($autoMarker) {
    Remove-Item -LiteralPath $markerTmp -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $markerRoot -Force -ErrorAction SilentlyContinue
}

Write-Host "Picture Replace Tools WPS installed for $pluginName."
Write-Host "Installed files: $targetFolder"
Write-Host "Installed version: $($plugin.GetAttribute('version'))"

if ($RestartWps -or $autoMarker) {
    Write-Host "Auto-restarting WPS Office..."
    Start-Sleep -Seconds 2
    # Graceful close first (lets WPS auto-save), then force-close stragglers.
    Get-Process -Name wps,wpp -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } |
        ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch {} }
    Start-Sleep -Seconds 8
    Get-Process -Name wps,wpp -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    $wpsExe = 'D:\WPS Office\12.1.0.28043\office6\wps.exe'
    if (-not (Test-Path -LiteralPath $wpsExe)) {
        $candidate = Get-ItemProperty 'HKLM:\SOFTWARE\Kingsoft\Office\12.0\common' -Name 'InstallRoot' -ErrorAction SilentlyContinue
        if ($candidate) { $wpsExe = Join-Path $candidate.InstallRoot 'office6\wps.exe' }
    }
    if (Test-Path -LiteralPath $wpsExe) {
        Start-Process -FilePath $wpsExe -ArgumentList '/prometheus','/wpp'
    }
    Write-Host "WPS Office restarted."
}
