param([string]$BuildDirectory = '')
$ErrorActionPreference='Stop'
if(!$BuildDirectory){$version=(Get-Content (Join-Path $PSScriptRoot '..\package.json') -Raw|ConvertFrom-Json).version;$BuildDirectory=Join-Path $PSScriptRoot ('..\build\'+$version)}
$assembly=[Reflection.Assembly]::LoadFrom((Join-Path $BuildDirectory 'PowerPointTools.dll'))
$method=$assembly.GetType('Dongsidaye.PowerPointTools.DwgSvg').GetMethod('Validate',[Reflection.BindingFlags]'Static,NonPublic')
$validate=[Delegate]::CreateDelegate([Func[string,double[]]],$method)
$scratch=Join-Path $env:TEMP ('ppt-dwg-svg-tests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch|Out-Null
$tests=@(
    @{Name='vector geometry';Valid=$true;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="-10 -10 200 100"><path d="M0 0 L100 50"/></svg>'},
    @{Name='local hatch reference';Valid=$true;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><defs><pattern id="h"/></defs><path fill="url(#h)" d="M0 0 L100 50"/></svg>'},
    @{Name='script';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><script>alert(1)</script></svg>'},
    @{Name='external image';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><image href="https://example.com/file.png"/></svg>'},
    @{Name='event attribute';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100" onload="alert(1)"/>'},
    @{Name='nonfinite origin';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="NaN Infinity 200 100"><path d="M0 0 L100 50"/></svg>'},
    @{Name='zero dimensions';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 0 100"><path d="M0 0 L100 50"/></svg>'},
    @{Name='external DTD';Valid=$false;Xml='<!DOCTYPE svg SYSTEM "file:///C:/private.txt"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"/>'},
    @{Name='animation URL mutation';Valid=$false;Xml='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100"><set attributeName="href" to="https://example.com"/></svg>'}
)
$failures=0
try {
    foreach($test in $tests) {
        $path=Join-Path $scratch 'drawing.svg'
        [IO.File]::WriteAllText($path,$test.Xml,[Text.UTF8Encoding]::new($false))
        $accepted=$true
        try {$null=$validate.Invoke([string]$path)}catch{$accepted=$false;if($test.Valid){Write-Output $_.Exception.GetBaseException().Message}}
        if($accepted -ne $test.Valid){$failures++;Write-Output ('FAIL '+$test.Name)}else{Write-Output ('PASS '+$test.Name)}
    }
}finally {
    $target=[IO.Path]::GetFullPath($scratch)
    $root=[IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')+'\'
    if($target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($target).StartsWith('ppt-dwg-svg-tests-')){Remove-Item -LiteralPath $target -Recurse -Force}
}
if($failures){throw ('SVG validation failures: '+$failures)}
