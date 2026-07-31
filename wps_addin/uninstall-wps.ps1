param()

$ErrorActionPreference = 'Stop'
$pluginName = 'picture-replace-tools-wps'
$relativeFolder = 'picture-replace-tools-wps_1.1.6'
$jsAddinsRoot = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$publishPath = Join-Path $jsAddinsRoot 'publish.xml'
$targetFolder = Join-Path $jsAddinsRoot $relativeFolder

if (Test-Path -LiteralPath $publishPath) {
    [xml]$current = Get-Content -LiteralPath $publishPath -Raw
    $root = $current.SelectSingleNode('/jsplugins')
    if ($null -ne $root) {
        foreach ($old in @($current.SelectNodes('/jsplugins/jsplugin[@name="' + $pluginName + '"]'))) {
            [void]$root.RemoveChild($old)
        }
        $current.Save($publishPath)
    }
}

if (Test-Path -LiteralPath $targetFolder) {
    Remove-Item -LiteralPath $targetFolder -Recurse -Force
}
foreach ($oldFolder in @(Get-ChildItem -LiteralPath $jsAddinsRoot -Directory -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like ($pluginName + '_*')
})) {
    Remove-Item -LiteralPath $oldFolder.FullName -Recurse -Force
}
Write-Host "Picture Replace Tools WPS uninstalled. Restart WPS Office to apply."

