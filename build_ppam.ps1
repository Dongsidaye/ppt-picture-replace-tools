# build_ppam.ps1
# 构建 PictureReplaceTools.ppam：
#   1. 用 PowerPoint COM 新建演示文稿，注入 VBA 模块（AddFromString 走 BSTR，中文不乱码）
#   2. SaveAs 为 PPAM（ppSaveAsOpenXMLAddin = 30）
#   3. 向包内注入 customUI（Ribbon 选项卡）
# 产物：pic_replace_addin\dist\PictureReplaceTools.ppam

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$basPath = Join-Path $root 'modPictureReplace.bas'
$uiPath  = Join-Path $root 'customUI.xml'
$distDir = Join-Path $root 'dist'
$ppamPath = Join-Path $distDir 'PictureReplaceTools.ppam'

New-Item -ItemType Directory -Force $distDir | Out-Null
if (Test-Path $ppamPath) { Remove-Item $ppamPath -Force }

# ---- 读取 VBA 源码（UTF-8），去掉 Attribute 行（AddFromString 不接受）----
$vba = [System.IO.File]::ReadAllText($basPath, [System.Text.Encoding]::UTF8)
$vba = ($vba -split "`r?`n" | Where-Object { $_ -notmatch '^Attribute\s' }) -join "`r`n"

# ---- 1+2. 注入模块并另存为 PPAM ----
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Add()
    $vbproj = $pres.VBProject
    $comp = $vbproj.VBComponents.Add(1)   # vbext_ct_StdModule
    $comp.Name = 'modPictureReplace'
    $comp.CodeModule.AddFromString($vba)

    $cm = $comp.CodeModule
    if ($cm.CountOfLines -lt 50) { throw 'VBA 模块行数异常，注入可能失败' }

    $pres.SaveAs($ppamPath, 30)   # ppSaveAsOpenXMLAddin
    $pres.Close()
    Write-Host "PPAM saved: $ppamPath"
}
finally {
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}

# ---- 3. 注入 customUI ----
# .NET ZipArchive Update corrupts the Office-generated empty presentation.xml
# entry (method STORE with stale DEFLATE bytes). Rebuild the archive instead.
$python = Join-Path (Split-Path -Parent $root) 'ppt-master\.venv\Scripts\python.exe'
if (-not (Test-Path $python)) { throw "Python runtime not found: $python" }
& $python (Join-Path $root 'inject_customui.py') $ppamPath $uiPath
if ($LASTEXITCODE -ne 0) { throw 'customUI injection failed' }

Write-Host 'customUI injected.'
Write-Host 'package validation passed.'
Write-Host "Build complete: $ppamPath"
