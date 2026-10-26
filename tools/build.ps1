param([ValidateSet('Debug','Release')][string]$Configuration='Release',
      [ValidatePattern('^\d+\.\d+\.\d+(\.\d+)?$')][string]$Version='1.0.0',
      [string]$InnoSetupCompiler='',
      [switch]$SkipSetup)
$ErrorActionPreference='Stop'
if($SkipSetup){
    & (Join-Path $PSScriptRoot 'publish-standalone.ps1') -Configuration $Configuration
}else{
    & (Join-Path $PSScriptRoot 'build-setup.ps1') -Configuration $Configuration -Version $Version -InnoSetupCompiler $InnoSetupCompiler
}
