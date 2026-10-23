param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot)).TrimEnd('\','/')
$output=[IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\','/')
$original=Join-Path $root 'Referrings'
$legacy=Join-Path $output 'Referrings'
if(-not(Test-Path -LiteralPath $legacy)){return}
# Move only obsolete runtime copies inside this workspace. Never move original
# references or user-selected directories outside the project.
if(-not $output.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){return}
if($legacy.Equals($original,[StringComparison]::OrdinalIgnoreCase) -or $legacy.StartsWith($original+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to move the original reference files.'}
$backup=Join-Path $root ('.tools\retired-runtime-references\'+[Guid]::NewGuid().ToString('N')+'\Referrings')
if(-not $legacy.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase) -or -not $backup.StartsWith($root+'\.tools\',[StringComparison]::OrdinalIgnoreCase)){throw 'Runtime reference archive paths are outside the workspace.'}
New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force|Out-Null
Move-Item -LiteralPath $legacy -Destination $backup
Write-Host 'Archived obsolete runtime references; the original source files remain unchanged.'
