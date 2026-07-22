# inject_customui.ps1 - convert PPTM to PPAM and inject customUI (Ribbon). ASCII only.
# Steps: 1) pptm -> ppam (rewrite main part content type to addin.macroEnabled.main+xml)
#        2) inject customUI part + relationship + content type override
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pptmPath = Join-Path $root 'dist\PictureReplaceTools.pptm'
$ppamPath = Join-Path $root 'dist\PictureReplaceTools.ppam'
$uiPath = Join-Path $root 'customUI.xml'

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipTextFromPath([string]$zipPath, [string]$name) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { $zip.Dispose(); throw "missing part: $name" }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Close(); $zip.Dispose()
    return $t
}
function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { throw "missing part: $name" }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Close()
    return $t
}
function Write-ZipText($zip, [string]$name, [string]$text) {
    $old = $zip.GetEntry($name)
    if ($null -ne $old) { $old.Delete() }
    $e = $zip.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text); $sw.Close()
}

$uiXml = [System.IO.File]::ReadAllText($uiPath, [System.Text.Encoding]::UTF8)

# ---- 1. pptm -> ppam: rewrite main part content type ----
if (Test-Path $ppamPath) { Remove-Item $ppamPath -Force }
Copy-Item $pptmPath $ppamPath
$ct = Read-ZipTextFromPath $ppamPath '[Content_Types].xml'
$ctNew = $ct -replace 'application/vnd\.ms-powerpoint\.presentation\.macroEnabled\.main\+xml', 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml'
if ($ctNew -eq $ct) { throw 'main part content type not found for conversion' }
$zip = [System.IO.Compression.ZipFile]::Open($ppamPath, 'Update')
try {
    Write-ZipText $zip '[Content_Types].xml' $ctNew

    # ---- 2. inject customUI ----
    $ct = Read-ZipText $zip '[Content_Types].xml'
    if ($ct -notlike '*customUI*') {
        $ct = $ct -replace '</Types>', '<Override PartName="/customUI/customUI.xml" ContentType="application/vnd.ms-office.customUI+xml"/></Types>'
        Write-ZipText $zip '[Content_Types].xml' $ct
    }
    $rels = Read-ZipText $zip '_rels/.rels'
    if ($rels -notlike '*ui/extensibility*') {
        $rels = $rels -replace '</Relationships>', '<Relationship Id="rIdPicReplaceCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>'
        Write-ZipText $zip '_rels/.rels' $rels
    }
    Write-ZipText $zip 'customUI/customUI.xml' $uiXml
}
finally {
    $zip.Dispose()
}
Write-Host 'customUI injected.'
