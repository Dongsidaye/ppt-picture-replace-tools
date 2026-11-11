param(
    [Parameter(Mandatory=$true)][string]$DestinationFolder,
    [string]$CacheDirectory = (Join-Path $env:LOCALAPPDATA 'PictureReplaceToolsBuild\OcrModelsV5')
)
$ErrorActionPreference='Stop'
$models=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'ocr-assets\models.json') -Raw|ConvertFrom-Json
$destination=[IO.Path]::GetFullPath($DestinationFolder)
$cache=[IO.Path]::GetFullPath($CacheDirectory)
New-Item -ItemType Directory -Path $destination,$cache -Force|Out-Null
foreach($model in $models) {
    if([IO.Path]::GetFileName($model.file) -ne $model.file -or $model.sha256 -notmatch '^[a-f0-9]{64}$' -or !([uri]$model.url).Scheme.Equals('https')) {
        throw 'Invalid pinned OCR model manifest'
    }
    $cached=Join-Path $cache $model.file
    $valid=(Test-Path -LiteralPath $cached -PathType Leaf) -and ((Get-FileHash -LiteralPath $cached -Algorithm SHA256).Hash -eq $model.sha256)
    if(!$valid) {
        $download=Join-Path $cache ([guid]::NewGuid().ToString('N')+'.download')
        try {
            Invoke-WebRequest -Uri $model.url -OutFile $download -UseBasicParsing
            if((Get-FileHash -LiteralPath $download -Algorithm SHA256).Hash -ne $model.sha256) {
                throw ('OCR model SHA-256 mismatch: '+$model.file)
            }
            Move-Item -LiteralPath $download -Destination $cached -Force
        }finally {
            if(Test-Path -LiteralPath $download){Remove-Item -LiteralPath $download -Force}
        }
    }
    Copy-Item -LiteralPath $cached -Destination (Join-Path $destination $model.file) -Force
    Write-Output ('Verified OCR model '+$model.file)
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ocr-assets\NOTICE.md') -Destination $destination -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ocr-assets\LICENSE-PaddleOCR.txt') -Destination $destination -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'ocr-assets\code-licenses') -Destination $destination -Recurse -Force
