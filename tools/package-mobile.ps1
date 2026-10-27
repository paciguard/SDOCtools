param([switch]$SkipAndroidBuild)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not $SkipAndroidBuild){& (Join-Path $PSScriptRoot 'build-mobile-android.ps1')}
$destination=Join-Path $root 'dist\CFL-Mobile\iOS'
$stage=Join-Path $root ('.tools\mobile-kit-staging\'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $destination,$stage -Force|Out-Null
$files=@()
foreach($folder in 'ZTools.Mobile.iOS','ZTools.Offline','ZTools.Mobile.Shared','ZTools.Standalone/Core','licenses'){
    $files+=Get-ChildItem (Join-Path $root $folder) -Recurse -File|Where-Object {$_.FullName -notmatch '[\\/](bin|obj)[\\/]'}
}
foreach($relative in 'ZTools.Standalone/CadAdapter.cs','ZTools.Standalone/DrawingOutput.cs','ZTools.Standalone/wwwroot/cfl-logo.jpg','ZTools.Standalone/wwwroot/vendor/three.module.js','ZTools.Standalone/wwwroot/vendor/three.core.js','ZTools.Standalone/wwwroot/vendor/THREE-LICENSE.txt','Referrings/Ztools/ZTools.xls','Resources/Workbooks/ZTools.xls','Resources/Workbooks/Input.xls','Referrings/Ztools/TEMPLATE/GA.dwg','Referrings/Ztools/TEMPLATE/PID.dwg','Referrings/Ztools/TEMPLATE/SEC.dwg','tools/build-mobile-ios.sh','MOBILE-INSTALL.md','THIRD-PARTY-NOTICES.md'){$files+=Get-Item (Join-Path $root $relative)}
foreach($file in $files){
    $relative=$file.FullName.Substring($root.Length).TrimStart('\','/')
    $copy=Join-Path $stage $relative;New-Item -ItemType Directory -Path (Split-Path -Parent $copy) -Force|Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $copy
    if((Get-FileHash $copy).Hash -ne (Get-FileHash $file.FullName).Hash){throw "Build kit copy differs: $relative"}
}
$zip=Join-Path $destination 'CFL-SubmitTools-iPhone-Mac-Build-Kit.zip'
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal -Force
Copy-Item -LiteralPath (Join-Path $root 'MOBILE-INSTALL.md') -Destination (Join-Path $root 'dist\CFL-Mobile\INSTALL.md') -Force
$manifest=@{android='Android/CFL-SubmitTools-Android-arm64.apk';iphoneBuildKit='iOS/CFL-SubmitTools-iPhone-Mac-Build-Kit.zip';iphoneInstallerStatus='Requires running the included Mac signing script; no signed IPA is supplied by the Windows build.';files=@()}
foreach($path in @((Join-Path $root 'dist\CFL-Mobile\Android\CFL-SubmitTools-Android-arm64.apk'),$zip)){
    if(Test-Path -LiteralPath $path){$manifest.files+=@{name=(Split-Path -Leaf $path);bytes=(Get-Item $path).Length;sha256=(Get-FileHash $path).Hash.ToLowerInvariant()}}
}
$manifest|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'dist\CFL-Mobile\package-manifest.json') -Encoding UTF8
Write-Host "Mobile packages: $root\dist\CFL-Mobile"
$global:LASTEXITCODE=0
