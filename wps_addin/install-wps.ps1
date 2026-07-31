param()

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

$newPlugin = $current.CreateElement('jsplugin')
foreach ($attribute in @('name', 'type', 'version', 'url', 'enable', 'install', 'customDomain')) {
    if ($plugin.HasAttribute($attribute)) {
        $newPlugin.SetAttribute($attribute, $plugin.GetAttribute($attribute))
    }
}
[void]$root.AppendChild($current.ImportNode($newPlugin, $true))
$current.Save($publishPath)

Write-Host "Picture Replace Tools WPS installed for $pluginName."
Write-Host "Installed files: $targetFolder"
Write-Host "Restart WPS Office before using the ribbon tab."
