param([string]$Package=(Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\CFL-SubmitTools'))
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$testRoot=Join-Path $root '.tools\startup-check'
$detached=Join-Path $testRoot ([Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $detached -Force|Out-Null
Copy-Item -Path (Join-Path $Package '*') -Destination $detached -Recurse -Force
$exe=Join-Path $detached 'ZTools.Standalone.exe'
$response=Join-Path $detached 'response.json'
$oldUrls=$env:ASPNETCORE_URLS
$children=@()
function Read-Local([int]$Port,[string]$Path='/api/health'){
    & curl.exe --noproxy '*' --silent --max-time 3 --fail "http://127.0.0.1:$Port$Path" --output $response
    if($LASTEXITCODE -ne 0){return $null}
    return Get-Content -LiteralPath $response -Raw
}
function Wait-Ready([int]$Port){
    for($i=0;$i -lt 40;$i++){
        $body=Read-Local $Port
        if($body){return $body|ConvertFrom-Json}
        Start-Sleep -Milliseconds 250
    }
    throw "Application did not start on port $Port."
}
function Start-Check([string[]]$Arguments){
    $id=[Guid]::NewGuid().ToString('N')
    return Start-Process -FilePath $exe -ArgumentList $Arguments -WorkingDirectory $testRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $detached ($id+'.log')) -RedirectStandardError (Join-Path $detached ($id+'.err'))
}
try{
    $env:ASPNETCORE_URLS=$null
    foreach($relative in 'wwwroot\index.html','wwwroot\vendor\three.module.js','ZTools.Standalone.dll'){
        if(-not(Test-Path -LiteralPath (Join-Path $detached $relative))){throw "Build output is missing $relative"}
    }
    $first=Start-Check @('--no-browser','--port','55191');$children+=$first
    $health=Wait-Ready 55191
    if($health.application -ne 'CFL.SubmitTools'){throw 'Incorrect application identity'}
    if(-not (Read-Local 55191 '/') -or -not (Read-Local 55191 '/app.js')){throw 'Packaged browser files are not served'}
    $before=(Read-Local 55191 '/api/state')|ConvertFrom-Json
    foreach($template in $before.data.templates){if($template.source -ne 'embedded' -or -not $template.exists){throw 'Application did not use its embedded DWG templates'}}
    $second=Start-Check @('--no-browser','--port','55191');$children+=$second
    if(-not $second.WaitForExit(15000) -or $second.ExitCode -ne 0){throw 'Repeat launch crashed instead of reusing the running application'}
    $after=(Read-Local 55191 '/api/state')|ConvertFrom-Json
    if($before.token -ne $after.token){throw 'Repeat launch replaced the original application session'}
    Write-Host 'PASS executable launches from a detached folder and repeat launch preserves its running session'
    & (Join-Path $root 'tools\prepare-build.ps1') -OutputDirectory $Package -Ports 55191
    if(-not (Read-Local 55191)){throw 'Build preparation closed a different application copy'}
    # Visual Studio runs this helper through Windows PowerShell 5, whose default
    # file encoding differs from PowerShell 7. State includes Chinese template labels.
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'tools\prepare-build.ps1') -OutputDirectory $detached -Ports 55191
    if($LASTEXITCODE -ne 0){throw 'Windows PowerShell rebuild preparation failed'}
    if(-not $first.WaitForExit(15000) -or $first.ExitCode -ne 0){throw 'Application did not close cleanly before rebuilding'}
    Write-Host 'PASS rebuild preparation closes only the matching idle output copy'
    $env:ASPNETCORE_URLS='http://127.0.0.1:55192'
    $configured=Start-Check @('--no-browser');$children+=$configured
    $null=Wait-Ready 55192
    & (Join-Path $root 'tools\prepare-build.ps1') -OutputDirectory $detached -Ports 55192
    if(-not $configured.WaitForExit(15000) -or $configured.ExitCode -ne 0){throw 'Configured-address application did not close cleanly'}
    Write-Host 'PASS Visual Studio launch address is honored'
    $invalid=Start-Check @('--no-browser','--port','0');$children+=$invalid
    if(-not $invalid.WaitForExit(15000) -or $invalid.ExitCode -ne 1 -or -not(Test-Path (Join-Path $detached 'application-error.log'))){throw 'Invalid startup settings were not reported with a readable error log'}
    Write-Host 'PASS startup failure is logged with a controlled exit code'
    Write-Host 'STANDALONE_STARTUP_OK'
}finally{
    $env:ASPNETCORE_URLS=$oldUrls
    # These are private test copies with no workbook opened; never stop user Excel.
    foreach($child in $children){try{if(-not $child.HasExited){Stop-Process -Id $child.Id -Force}}catch{}}
}
