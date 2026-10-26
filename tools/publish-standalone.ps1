param([ValidateSet('Debug','Release')][string]$Configuration='Release')
$ErrorActionPreference='Stop'
$projectRoot=Split-Path -Parent $PSScriptRoot
$destination=Join-Path $projectRoot 'dist\CFL-SubmitTools'
& (Join-Path $PSScriptRoot 'prepare-build.ps1') -OutputDirectory $destination
$env:MSBUILDDISABLENODEREUSE='1'
$env:DOTNET_CLI_WORKLOAD_UPDATE_NOTIFY_DISABLE='true'
$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
New-Item -ItemType Directory -Force -Path (Join-Path $projectRoot '.tools')|Out-Null
function Invoke-Dotnet([string[]]$Arguments,[string]$Phase){
    $log=Join-Path $projectRoot ('.tools/build-'+$Phase+'-'+[Guid]::NewGuid().ToString('N'))
    $argumentLine=($Arguments|ForEach-Object {'"'+$_+'"'}) -join ' '
    $process=Start-Process dotnet -ArgumentList $argumentLine -WorkingDirectory $projectRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput ($log+'.log') -RedirectStandardError ($log+'.err')
    if(-not $process.WaitForExit(1200000)){throw "dotnet $Phase did not finish within 20 minutes."}
    Get-Content ($log+'.log')
    Get-Content ($log+'.err')
    if($process.ExitCode -ne 0){throw "dotnet $Phase failed with exit code $($process.ExitCode)."}
}
Invoke-Dotnet @('restore',(Join-Path $projectRoot 'ZTools.Standalone\ZTools.Standalone.csproj'),'--packages',(Join-Path $projectRoot '.tools\nuget'),'--disable-build-servers','/nr:false','/m:1') 'restore'
Invoke-Dotnet @('publish',(Join-Path $projectRoot 'ZTools.Standalone\ZTools.Standalone.csproj'),'--no-restore','--configuration',$Configuration,'--output',$destination,'--disable-build-servers','/p:UseSharedCompilation=false','/nr:false','/m:1','-v:minimal') 'publish'
Copy-Item -LiteralPath (Join-Path $projectRoot 'tools\Start.ps1') -Destination (Join-Path $destination 'Start.ps1') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'THIRD-PARTY-NOTICES.md') -Destination $destination -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination $destination -Force
function Get-SharedFileHash([string]$Path){
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $hasher=[Security.Cryptography.SHA256]::Create()
    try{return [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-','')}
    finally{$stream.Dispose();$hasher.Dispose()}
}
$inventoryLog=Join-Path $projectRoot ('.tools/embedded-inventory-'+[Guid]::NewGuid().ToString('N'))
$inspect=Start-Process -FilePath (Join-Path $destination 'ZTools.Standalone.exe') -ArgumentList '--verify-templates' -WindowStyle Hidden -PassThru -RedirectStandardOutput ($inventoryLog+'.json') -RedirectStandardError ($inventoryLog+'.err')
if(-not $inspect.WaitForExit(30000) -or $inspect.ExitCode -ne 0){throw 'Embedded template verification failed.'}
$inventory=@(Get-Content ($inventoryLog+'.json') -Raw|ConvertFrom-Json)
if($inventory.Count -ne 4){throw 'Expected three embedded DWGs and the embedded Excel workbook.'}
foreach($asset in $inventory){
    $relative=if($asset.name -eq 'ZTools.xls'){'Referrings\Ztools\ZTools.xls'}else{'Referrings\Ztools\TEMPLATE\'+$asset.name}
    if((Get-SharedFileHash (Join-Path $projectRoot $relative)) -ne $asset.sha256.ToUpperInvariant()){throw "Embedded template mismatch: $($asset.name)"}
}
if(@(Get-ChildItem $destination -Recurse -File | Where-Object {$_.Extension -in '.dwg','.xls','.dvb','.exb'}).Count){throw 'External business templates remain in the published application.'}
$licenses=Join-Path $destination 'licenses';New-Item -ItemType Directory -Force -Path $licenses|Out-Null
foreach($package in 'acadsharp/3.8.0','pdfsharp-gdi/6.2.4'){
    $packageRoot=Join-Path (Join-Path $projectRoot '.tools/nuget') $package
    if(Test-Path -LiteralPath $packageRoot){foreach($license in Get-ChildItem $packageRoot -Recurse -File|Where-Object {$_.Name -match '^(LICENSE|LICENSE-MIT|License\.txt|LICENSE\.md)$'}){Copy-Item -LiteralPath $license.FullName -Destination (Join-Path $licenses (($package.Replace('/','-'))+'-'+$license.Name)) -Force}}
}
Copy-Item -LiteralPath (Join-Path $projectRoot 'ZTools.Standalone/wwwroot/vendor/THREE-LICENSE.txt') -Destination (Join-Path $licenses 'three-LICENSE.txt') -Force
if(Test-Path -LiteralPath (Join-Path $projectRoot 'licenses')){Copy-Item -Path (Join-Path $projectRoot 'licenses/*') -Destination $licenses -Recurse -Force}
Write-Host "Built standalone application: $destination\Start.ps1"
# Start-Process checks its own exit code above; it does not reset LASTEXITCODE.
# PowerShell CI wrappers use this value as the script's final process status.
$global:LASTEXITCODE=0
