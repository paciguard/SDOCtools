param([ValidateSet('Debug','Release')][string]$Configuration='Release',
      [ValidatePattern('^\d+\.\d+\.\d+(\.\d+)?$')][string]$Version='1.0.0',
      [string]$InnoSetupCompiler='')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not $InnoSetupCompiler){
    $compiler=Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if($compiler){$InnoSetupCompiler=$compiler.Source}
    else{$InnoSetupCompiler=Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'}
}
if(-not(Test-Path -LiteralPath $InnoSetupCompiler)){throw 'Install Inno Setup 6.5 or later, or supply -InnoSetupCompiler.'}
& (Join-Path $PSScriptRoot 'publish-standalone.ps1') -Configuration $Configuration -SelfContained
$package=Join-Path $root 'dist\CFL-SubmitTools'
foreach($file in 'coreclr.dll','hostfxr.dll','Microsoft.AspNetCore.dll','ZTools.Standalone.exe'){
    if(-not(Test-Path -LiteralPath (Join-Path $package $file))){throw "Self-contained package is missing $file"}
}

# Derive the installer and shortcut branding from the supplied CFL logo.
Add-Type -AssemblyName System.Drawing
$logo=[Drawing.Image]::FromFile((Join-Path $root 'ZTools.Standalone\wwwroot\cfl-logo.jpg'))
$bitmap=New-Object Drawing.Bitmap 64,64
$graphics=[Drawing.Graphics]::FromImage($bitmap)
$png=New-Object IO.MemoryStream
try{
    $graphics.Clear([Drawing.Color]::White)
    $graphics.InterpolationMode=[Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $scale=[Math]::Min(60.0/$logo.Width,60.0/$logo.Height)
    $width=[int]($logo.Width*$scale);$height=[int]($logo.Height*$scale)
    $graphics.DrawImage($logo,[int]((64-$width)/2),[int]((64-$height)/2),$width,$height)
    $bitmap.Save((Join-Path $package 'cfl-setup.bmp'),[Drawing.Imaging.ImageFormat]::Bmp)
    $bitmap.Save($png,[Drawing.Imaging.ImageFormat]::Png)
    $icon=[IO.BinaryWriter]::new([IO.File]::Create((Join-Path $package 'cfl.ico')))
    try{
        $icon.Write([uint16]0);$icon.Write([uint16]1);$icon.Write([uint16]1)
        $icon.Write([byte]64);$icon.Write([byte]64);$icon.Write([byte]0);$icon.Write([byte]0)
        $icon.Write([uint16]1);$icon.Write([uint16]32);$icon.Write([uint32]$png.Length);$icon.Write([uint32]22)
        $icon.Write($png.ToArray())
    }finally{$icon.Dispose()}
}finally{$png.Dispose();$graphics.Dispose();$bitmap.Dispose();$logo.Dispose()}

$output=Join-Path $root 'dist\CFL-Setup'
New-Item -ItemType Directory -Path $output -Force|Out-Null
& $InnoSetupCompiler '/Q' ("/DAppVersion=$Version") ("/DPackageDir=$package") ("/DSetupDir=$output") (Join-Path $PSScriptRoot 'installer\CFL-SubmitTools.iss')
if($LASTEXITCODE -ne 0){throw 'Installer compilation failed.'}
$setup=Join-Path $output "CFL-SubmitTools-Setup-$Version-win-x64.exe"
if(-not(Test-Path -LiteralPath $setup)){throw 'The setup executable was not created.'}
$commit=& git -C $root rev-parse HEAD
if($LASTEXITCODE -ne 0){throw 'Could not record the source commit.'}
$setupStream=[IO.File]::OpenRead($setup)
$sha256=[Security.Cryptography.SHA256]::Create()
try{$hash=[BitConverter]::ToString($sha256.ComputeHash($setupStream)).Replace('-','').ToLowerInvariant()}
finally{$sha256.Dispose();$setupStream.Dispose()}
@{file=(Split-Path -Leaf $setup);version=$Version;configuration=$Configuration;sourceCommit=$commit;sha256=$hash;bytes=(Get-Item $setup).Length;runtime='win-x64 self-contained';signing='Unsigned; no publisher signing certificate configured'} |
    ConvertTo-Json | Set-Content (Join-Path $output 'package-manifest.json') -Encoding UTF8
"$hash  $(Split-Path -Leaf $setup)" | Set-Content ($setup+'.sha256') -Encoding ASCII
Write-Host "Windows setup: $setup"
$global:LASTEXITCODE=0
