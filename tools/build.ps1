param([ValidateSet('Debug','Release')][string]$Configuration='Release')
& (Join-Path $PSScriptRoot 'publish-standalone.ps1') -Configuration $Configuration
