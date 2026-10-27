$ErrorActionPreference='Stop'
& python (Join-Path $PSScriptRoot 'prepare-input-xlsx.py')
if($LASTEXITCODE -ne 0){throw 'XLSX template preparation failed.'}
