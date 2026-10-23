param([string]$AndroidSdkDirectory='C:\Program Files (x86)\Android\android-sdk',[string]$JavaSdkDirectory='C:\Program Files (x86)\Android\openjdk\jdk-17.0.14')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$destination=Join-Path $root 'dist\CFL-Mobile\Android'
$env:DOTNET_CLI_WORKLOAD_UPDATE_NOTIFY_DISABLE='true';$env:DOTNET_CLI_TELEMETRY_OPTOUT='1'
New-Item -ItemType Directory -Path (Join-Path $root '.tools'),$destination -Force|Out-Null
$log=Join-Path $root '.tools\mobile-android-publish'
$arguments=@('publish',(Join-Path $root 'ZTools.Mobile.Android\ZTools.Mobile.Android.csproj'),'-c','Release','-r','android-arm64','--disable-build-servers',('/p:RestorePackagesPath='+(Join-Path $root '.tools\nuget')),('/p:AndroidSdkDirectory='+$AndroidSdkDirectory),('/p:JavaSdkDirectory='+$JavaSdkDirectory),'/p:UseSharedCompilation=false','/nr:false','/m:1','-v:minimal')
$process=Start-Process dotnet -ArgumentList (($arguments|ForEach-Object {'"'+$_+'"'}) -join ' ') -WorkingDirectory (Join-Path $root 'ZTools.Mobile.Android') -WindowStyle Hidden -PassThru -RedirectStandardOutput ($log+'.log') -RedirectStandardError ($log+'.err')
if(-not $process.WaitForExit(600000)){throw 'Android publication timed out.'}
Get-Content ($log+'.log'),($log+'.err')
if($process.ExitCode -ne 0){throw 'Android publication failed.'}
$apk=Join-Path $root 'ZTools.Mobile.Android\bin\Release\net9.0-android35.0\android-arm64\publish\com.paciguard.sdoctools-Signed.apk'
if(-not(Test-Path -LiteralPath $apk)){throw 'The signed APK was not created.'}
Copy-Item -LiteralPath $apk -Destination (Join-Path $destination 'CFL-SubmitTools-Android-arm64.apk') -Force
Write-Host "Android installer: $destination\CFL-SubmitTools-Android-arm64.apk"
$global:LASTEXITCODE=0
