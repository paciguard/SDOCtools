$ErrorActionPreference='Stop'
$application=Join-Path $PSScriptRoot 'ZTools.Standalone.exe'
if(-not(Test-Path -LiteralPath $application)){throw 'Build the application with tools/build.ps1 first, then run dist/CFL-SubmitTools/Start.ps1.'}
# The application identifies and reopens an existing instance itself.
# Unique logs avoid sharing violations when the launcher is clicked twice.
$log='application-'+[DateTime]::Now.ToString('yyyyMMdd-HHmmssfff')
Start-Process -FilePath $application -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput (Join-Path $PSScriptRoot ($log+'.log')) -RedirectStandardError (Join-Path $PSScriptRoot ($log+'-error.log'))
