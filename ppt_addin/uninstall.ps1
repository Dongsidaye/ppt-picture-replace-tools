$ErrorActionPreference = 'Stop'
$reg = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'registration.json') -Raw | ConvertFrom-Json
try {
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $ppt.COMAddIns.Item($reg.addinProgId).Connect = $false
    [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
} catch { }
$views = @([Microsoft.Win32.RegistryView]::Registry32)
if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
$scopePath = Join-Path $env:LOCALAPPDATA 'Dongsidaye\PowerPointTools\registration-scope.txt'
$hives = @([Microsoft.Win32.RegistryHive]::CurrentUser)
if ((Test-Path -LiteralPath $scopePath) -and (Get-Content -LiteralPath $scopePath -Raw).Trim() -eq 'LocalMachine') { $hives += [Microsoft.Win32.RegistryHive]::LocalMachine }
foreach ($hive in $hives) {
foreach ($view in $views) {
    $hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
    try {
        foreach ($key in @("Software\Microsoft\Office\PowerPoint\Addins\$($reg.addinProgId)","Software\Classes\$($reg.addinProgId)","Software\Classes\$($reg.paneProgId)","Software\Classes\CLSID\$($reg.addinClsid)","Software\Classes\CLSID\$($reg.paneClsid)")) {
            $hkcu.DeleteSubKeyTree($key, $false)
        }
        if ($hive -eq [Microsoft.Win32.RegistryHive]::LocalMachine) {
            foreach ($virtualSoftware in @('Software','Software\Wow6432Node')) {
                $hkcu.DeleteSubKeyTree("Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\$virtualSoftware\Microsoft\Office\PowerPoint\Addins\$($reg.addinProgId)", $false)
            }
        }
    } finally { $hkcu.Dispose() }
}
}
Write-Host 'PowerPoint toolbox unregistered. Documents, the WPS add-in, and installed recovery files were retained.'
