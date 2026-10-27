param([string]$Setup=(Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\CFL-Setup\CFL-SubmitTools-Setup-1.0.0-win-x64.exe'))
$ErrorActionPreference='Stop'
$registration='HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{1A9B83B8-65B6-49F5-BA87-C79DD195710F}_is1'
if(Test-Path $registration){throw 'An installation is already registered. Run this check with a clean test user.'}
$root=Split-Path -Parent $PSScriptRoot
$testRoot=Join-Path $root ('.tools\setup-test-'+[Guid]::NewGuid().ToString('N'))
$install=Join-Path $testRoot 'installed app'
$group='CFL Setup Test '+[Guid]::NewGuid().ToString('N')
New-Item -ItemType Directory -Path $testRoot -Force|Out-Null
$exe=Join-Path $install 'ZTools.Standalone.exe'
$uninstall=Join-Path $install 'unins000.exe'
$child=$null
function Invoke-Setup([string]$File,[string[]]$Arguments){
    $line=($Arguments|ForEach-Object {'"'+$_+'"'}) -join ' '
    $process=Start-Process -FilePath $File -ArgumentList $line -WindowStyle Hidden -Wait -PassThru
    if($process.ExitCode -ne 0){throw "Setup operation failed: $($process.ExitCode)"}
}
function Read-Api([string]$Path){
    $response=Join-Path $testRoot 'response.json'
    & curl.exe --noproxy '*' --silent --max-time 3 --fail "http://127.0.0.1:55195$Path" --output $response
    if($LASTEXITCODE -ne 0){return $null}
    Get-Content -LiteralPath $response -Raw|ConvertFrom-Json
}
try{
    $arguments=@('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/SP-','/CURRENTUSER',"/DIR=$install","/GROUP=$group",'/TASKS=',"/LOG=$(Join-Path $testRoot 'install.log')")
    Invoke-Setup $Setup $arguments
    if(-not(Test-Path $exe) -or -not(Test-Path $uninstall)){throw 'Installed executable or uninstaller is missing.'}
    $runtime=Get-Content (Join-Path $install 'ZTools.Standalone.runtimeconfig.json') -Raw|ConvertFrom-Json
    if(-not $runtime.runtimeOptions.includedFrameworks -or $runtime.runtimeOptions.frameworks){throw 'Installed app is not self-contained.'}
    foreach($asset in 'coreclr.dll','hostfxr.dll','cfl.ico','wwwroot\cfl-logo.jpg','licenses\Microsoft-NET-LICENSE.txt'){
        if(-not(Test-Path (Join-Path $install $asset))){throw "Installed asset is missing: $asset"}
    }
    $inventory=& $exe --verify-templates|ConvertFrom-Json
    if($LASTEXITCODE -ne 0 -or @($inventory).Count -ne 5){throw 'Installed embedded templates could not be read.'}
    foreach($asset in $inventory){
        $relative=if($asset.name -in 'ZTools.xls','Input.xlsx'){'Resources\Workbooks\'+$asset.name}else{'Referrings\Ztools\TEMPLATE\'+$asset.name}
        if((Get-FileHash (Join-Path $root $relative)).Hash.ToLowerInvariant() -ne $asset.sha256){throw 'Installed template differs from the original.'}
    }
    $child=Start-Process -FilePath $exe -ArgumentList '--no-browser','--port','55195' -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $testRoot 'app.log') -RedirectStandardError (Join-Path $testRoot 'app.err')
    $null=$child.Handle
    $health=$null
    for($i=0;$i -lt 40;$i++){$health=Read-Api '/api/health';if($health){break};Start-Sleep -Milliseconds 250}
    if(-not $health -or $health.application -ne 'CFL.SubmitTools' -or [IO.Path]::GetFullPath($health.baseDirectory).TrimEnd('\') -ne $install){throw 'Installed app did not start from its own folder.'}
    $state=Read-Api '/api/state'
    if(-not $state.token -or $state.data.editing){throw 'Installed app did not open an idle session.'}
    & curl.exe --noproxy '*' --silent --fail --max-time 8 -H "X-CFL-Token: $($state.token)" -H 'Content-Type: application/json' --data '{}' 'http://127.0.0.1:55195/api/action/close'
    if($LASTEXITCODE -ne 0 -or -not $child.WaitForExit(15000)){throw 'Installed app did not close cleanly.'}
    Write-Host 'PASS installed app starts and reads three original DWGs plus its model and input templates'
    $userFile=Join-Path $install 'user-project-preservation.txt'
    'User-created project data'|Set-Content $userFile
    $before=(Get-FileHash $userFile).Hash
    Invoke-Setup $Setup $arguments
    if((Get-FileHash $userFile).Hash -ne $before){throw 'Upgrade changed user-created data.'}
    Write-Host 'PASS upgrade preserves user-created files'
    Invoke-Setup $uninstall @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART',"/LOG=$(Join-Path $testRoot 'uninstall.log')")
    if((Test-Path $exe) -or (Test-Path $registration)){throw 'Uninstall left application files or registration.'}
    if((Get-FileHash $userFile).Hash -ne $before){throw 'Uninstall removed user-created data.'}
    Write-Host 'PASS uninstall removes installed app and preserves user-created files'
    Write-Host 'WINDOWS_SETUP_OK'
}finally{
    if($child){try{if(-not $child.HasExited){Stop-Process -Id $child.Id -Force}}catch{}}
    if((Test-Path $uninstall) -and (Test-Path $registration)){
        Invoke-Setup $uninstall @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART')
    }
}
$global:LASTEXITCODE=0
