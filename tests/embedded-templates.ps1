param([string]$Package=(Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\CFL-SubmitTools'))
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$manifest=Get-Content (Join-Path $root 'migration\source-manifest-20261023-embedded.json') -Raw|ConvertFrom-Json
$detached=Join-Path ([IO.Path]::GetTempPath()) ('CFL-Embedded-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $detached|Out-Null
Copy-Item -Path (Join-Path $Package '*') -Destination $detached -Recurse -Force
if(@(Get-ChildItem $detached -Recurse -File|Where-Object {$_.Extension -in '.dwg','.xls','.dvb','.exb'}).Count){throw 'The application still contains external business templates.'}
$exe=Join-Path $detached 'ZTools.Standalone.exe'
$inventoryFile=Join-Path $detached 'inventory.json'
$inspect=Start-Process $exe -ArgumentList '--verify-templates' -WorkingDirectory $detached -WindowStyle Hidden -PassThru -RedirectStandardOutput $inventoryFile -RedirectStandardError (Join-Path $detached 'inventory.err')
if(-not $inspect.WaitForExit(30000) -or $inspect.ExitCode -ne 0){throw 'Could not inspect embedded assets.'}
$assets=@(Get-Content $inventoryFile -Raw|ConvertFrom-Json)
if($assets.Count -ne 5){throw 'Expected three embedded DWGs and two workbook templates.'}
foreach($asset in $assets){
    if($asset.name -in 'ZTools.xls','Input.xls'){
        if($asset.sha256 -ne (Get-FileHash (Join-Path $root ('Resources\Workbooks\'+$asset.name))).Hash.ToLowerInvariant()){throw 'Embedded workbook differs from its generated source.'}
        continue
    }
    $original=$manifest|Where-Object {$_.path -eq ('Referrings/Ztools/'+$asset.name) -or $_.path -eq ('Referrings/Ztools/TEMPLATE/'+$asset.name)}
    if(@($original).Count -ne 1 -or $asset.sha256 -ne $original.sha256 -or $asset.bytes -ne $original.bytes){throw "Embedded bytes differ from the source: $($asset.name)"}
}
Write-Host 'PASS all five embedded hashes match the three original DWGs and generated workbook templates'
$probeLog=Join-Path $detached 'probe.log'
$probe=Start-Process $exe -ArgumentList '--probe' -WorkingDirectory $detached -WindowStyle Hidden -PassThru -RedirectStandardOutput $probeLog -RedirectStandardError (Join-Path $detached 'probe.err')
if(-not $probe.WaitForExit(120000) -or $probe.ExitCode -ne 0){Get-Content (Join-Path $detached 'probe.err');throw 'Embedded template probe failed.'}
foreach($kind in 'GA','PID','SEC'){
    $drawing=Join-Path $detached ('tests\standalone-probe\'+$kind+'.dwg')
    if(-not(Test-Path -LiteralPath $drawing)){throw "Embedded $kind template could not be read and written."}
}
$workbook=@(Get-ChildItem (Join-Path $detached 'tests\standalone-probe') -Recurse -Filter ZTools.xls)
if($workbook.Count -ne 1){throw 'The probe did not create a private Excel project copy.'}
& $exe --verify-input $workbook[0].FullName
if($LASTEXITCODE -ne 0){throw 'Generated public project is not an input-only workbook backed by the embedded model.'}
Write-Host 'PASS detached application reads all three embedded DWGs and creates an input-only Excel project without external databases'
Write-Host 'EMBEDDED_TEMPLATES_OK'
$global:LASTEXITCODE=0
