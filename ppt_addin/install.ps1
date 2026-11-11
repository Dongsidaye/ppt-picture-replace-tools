param([switch]$RegisterOnly)
$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot
$meta = Get-Content -LiteralPath (Join-Path $source 'package.json') -Raw | ConvertFrom-Json
$reg = Get-Content -LiteralPath (Join-Path $source 'registration.json') -Raw | ConvertFrom-Json
$checksums = Get-Content -LiteralPath (Join-Path $source 'checksums.json') -Raw | ConvertFrom-Json
$prefix = [IO.Path]::GetFullPath($source).TrimEnd('\') + '\'
foreach ($entry in $checksums.PSObject.Properties) {
    $path = [IO.Path]::GetFullPath((Join-Path $source $entry.Name))
    if (!$path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Invalid package path.' }
    if (!(Test-Path -LiteralPath $path -PathType Leaf) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $entry.Value) { throw "Package verification failed: $($entry.Name)" }
}
$dllHash = (Get-FileHash -LiteralPath (Join-Path $source 'PowerPointTools.dll') -Algorithm SHA256).Hash.Substring(0,12)
$admin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$uacDisabled = (Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -ErrorAction SilentlyContinue).EnableLUA -eq 0
$machine = $admin -and $uacDisabled
$hive = [Microsoft.Win32.RegistryHive]::CurrentUser
$base = Join-Path $env:LOCALAPPDATA 'Dongsidaye\PowerPointTools'
if ($machine) {
    # Windows ignores per-user COM registration in this pre-existing configuration.
    # Use an administrator-protected code directory; never change UAC settings.
    $hive = [Microsoft.Win32.RegistryHive]::LocalMachine
    $base = Join-Path $env:ProgramFiles 'Dongsidaye\PowerPointTools'
}
$target = Join-Path $base ($meta.version + '-' + $dllHash)
if (!(Test-Path -LiteralPath $target)) {
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Copy-Item -Path (Join-Path $source '*') -Destination $target -Recurse -Force
}
foreach ($entry in $checksums.PSObject.Properties) {
    $path = Join-Path $target $entry.Name
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $entry.Value) { throw "Installed file verification failed: $($entry.Name)" }
}
$dll = Join-Path $target 'PowerPointTools.dll'
$assembly = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$codebase = [Uri]::UnescapeDataString(([Uri]$dll).AbsoluteUri)
$views = @([Microsoft.Win32.RegistryView]::Registry32)
if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
foreach ($view in $views) {
    $hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
    try {
        foreach ($kind in @('addin','pane')) {
            $progId = $reg.($kind + 'ProgId'); $clsid = $reg.($kind + 'Clsid'); $class = $reg.($kind + 'Class')
            $prog = $hkcu.CreateSubKey("Software\Classes\$progId\CLSID")
            $prog.SetValue('', $clsid); $prog.Dispose()
            $cls = $hkcu.CreateSubKey("Software\Classes\CLSID\$clsid")
            $cls.SetValue('', $class)
            $p = $cls.CreateSubKey('ProgId'); $p.SetValue('', $progId); $p.Dispose()
            $server = $cls.CreateSubKey('InprocServer32')
            $server.SetValue('', 'mscoree.dll'); $server.SetValue('ThreadingModel','Both')
            $server.SetValue('Class',$class); $server.SetValue('Assembly',$assembly)
            $server.SetValue('RuntimeVersion','v4.0.30319'); $server.SetValue('CodeBase',$codebase)
            $versionKey = $server.CreateSubKey([Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString())
            $versionKey.SetValue('Class',$class); $versionKey.SetValue('Assembly',$assembly)
            $versionKey.SetValue('RuntimeVersion','v4.0.30319'); $versionKey.SetValue('CodeBase',$codebase)
            $versionKey.Dispose()
            $server.Dispose()
            $managedCategory = $cls.CreateSubKey('Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'); $managedCategory.Dispose()
            if ($kind -eq 'pane') {
                $control = $cls.CreateSubKey('Control'); $control.Dispose()
                $category = $cls.CreateSubKey('Implemented Categories\{40FC6ED3-2438-11CF-A3DB-080036F12502}'); $category.Dispose()
                $misc = $cls.CreateSubKey('MiscStatus'); $misc.SetValue('', '131473'); $misc.Dispose()
            }
            $cls.Dispose()
        }
        $addin = $hkcu.CreateSubKey("Software\Microsoft\Office\PowerPoint\Addins\$($reg.addinProgId)")
        $addin.SetValue('FriendlyName','东四大爷的工具箱 (PowerPoint)')
        $addin.SetValue('Description','图片清单、智能缩放、对象管理与设计工具')
        $addin.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
        $addin.SetValue('CommandLineSafe',1,[Microsoft.Win32.RegistryValueKind]::DWord)
        $addin.Dispose()
    } finally { $hkcu.Dispose() }
    if ($machine) {
        # Office persists a user override when an older instance is disconnected.
        $userHive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $userAddin = $userHive.CreateSubKey("Software\Microsoft\Office\PowerPoint\Addins\$($reg.addinProgId)")
            $userAddin.SetValue('FriendlyName','东四大爷的工具箱 (PowerPoint)')
            $userAddin.SetValue('Description','图片清单、智能缩放、对象管理与设计工具')
            $userAddin.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
            $userAddin.SetValue('CommandLineSafe',1,[Microsoft.Win32.RegistryValueKind]::DWord)
            $userAddin.Dispose()
        } finally { $userHive.Dispose() }
        # Click-to-Run can retain its own per-add-in LoadBehavior override after
        # a previous instance was disconnected. Repair only existing keys for
        # this product; never clear Office resiliency or other add-in settings.
        $machineHive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        try {
            foreach ($virtualSoftware in @('Software','Software\Wow6432Node')) {
                $shadow = $machineHive.OpenSubKey("Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\$virtualSoftware\Microsoft\Office\PowerPoint\Addins\$($reg.addinProgId)", $true)
                if ($null -ne $shadow) {
                    try {
                        $shadow.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
                        $shadow.SetValue('FriendlyName','东四大爷的工具箱 (PowerPoint)')
                        $shadow.SetValue('Description','图片清单、智能缩放、对象管理与设计工具')
                        $shadow.SetValue('CommandLineSafe',1,[Microsoft.Win32.RegistryValueKind]::DWord)
                    } finally { $shadow.Dispose() }
                }
            }
        } finally { $machineHive.Dispose() }
    }
}
Set-Content -LiteralPath (Join-Path $base 'installed-path.txt') -Value $target -Encoding UTF8
$userState = Join-Path $env:LOCALAPPDATA 'Dongsidaye\PowerPointTools'
New-Item -ItemType Directory -Path $userState -Force | Out-Null
Set-Content -LiteralPath (Join-Path $userState 'installed-path.txt') -Value $target -Encoding UTF8
Set-Content -LiteralPath (Join-Path $userState 'registration-scope.txt') -Value $hive -Encoding UTF8
Write-Host "Installed PowerPoint toolbox $($meta.version): $target"
if (!$RegisterOnly) {
    try {
        $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        $ppt.COMAddIns.Update()
        $item = $ppt.COMAddIns.Item($reg.addinProgId)
        if (!$item.Connect) { $item.Connect = $true; Write-Host 'Connected to the running PowerPoint session.' }
        else { Write-Host 'An instance is already loaded. The registered update will be used after PowerPoint next starts.' }
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    } catch { Write-Host ('Registration is complete. PowerPoint will load it on next startup. Live connection: ' + $_.Exception.Message) }
}
# Never close PowerPoint or mark presentations saved during installation.
