$ErrorActionPreference='Stop'
$application=Join-Path $PSScriptRoot 'ZTools.Standalone.exe'
if(-not(Test-Path -LiteralPath $application)){throw 'Build the application with tools/build.ps1 first, then run dist/CFL-SubmitTools/Start.ps1.'}
$running=Get-CimInstance Win32_Process -Filter "Name='ZTools.Standalone.exe'"|Where-Object {$_.ExecutablePath -eq $application}
if($running){Start-Process 'http://127.0.0.1:5099';return}
Start-Process -FilePath $application -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot 'application.log') -RedirectStandardError (Join-Path $PSScriptRoot 'application-error.log')
