param([Parameter(Mandatory=$true)][string]$OutputDirectory,[int[]]$Ports=@(5099,5100))
$ErrorActionPreference='Stop'
$outputPath=[IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\','/')
$curl=(Get-Command curl.exe -ErrorAction Stop).Source
$temp=Join-Path ([IO.Path]::GetTempPath()) ('cfl-build-'+[Guid]::NewGuid().ToString('N'))
function Request-Cfl([string]$Url,[string]$Token=''){
    $arguments=@('--noproxy','*','--silent','--show-error','--max-time','8','--output',$temp,'--write-out','%{http_code}')
    if($Token){$arguments+=@('-H',"X-CFL-Token: $Token",'-H','Content-Type: application/json','--data','{}')}
    $arguments+=$Url
    # An absent server is expected. Do not leak curl's exit code to the caller.
    $previousExit=$global:LASTEXITCODE
    try{$status=& $curl @arguments 2>$null;$requestExit=$LASTEXITCODE}catch{return $null}
    finally{$global:LASTEXITCODE=$previousExit}
    if($requestExit -ne 0){return $null}
    $body=[IO.File]::ReadAllText($temp,[Text.Encoding]::UTF8)
    return @{status=[string]$status;body=$body}
}
try{
    foreach($port in $Ports){
        $info=$null
        $url="http://127.0.0.1:$port"
        $health=Request-Cfl "$url/api/health"
        if(-not $health){continue}
        $ours=$false
        if($health.status -eq '200'){
            try{$info=$health.body|ConvertFrom-Json}catch{continue}
            if($info.application -eq 'CFL.SubmitTools'){$ours=[string]::Equals(([IO.Path]::GetFullPath($info.baseDirectory)).TrimEnd('\','/'),$outputPath,[StringComparison]::OrdinalIgnoreCase)}
        }
        if(-not $ours -and $health.status -eq '404'){
            # Compatibility with the first release, before /api/health existed.
            $legacy=Request-Cfl "$url/api/state"
            if($legacy -and $legacy.status -eq '200'){
                $state=$legacy.body|ConvertFrom-Json
                $templateRoot=Join-Path $outputPath 'Referrings\Ztools\TEMPLATE'
                $templates=@($state.data.templates)
                $ours=$state.token -match '^[0-9A-F]{48}$' -and $templates.Count -eq 3 -and @($templates|Where-Object {-not $_.path.StartsWith($templateRoot+'\',[StringComparison]::OrdinalIgnoreCase)}).Count -eq 0
            }
        }
        if(-not $ours){continue}
        $current=Request-Cfl "$url/api/state"
        if(-not $current -or $current.status -ne '200'){throw 'The application is busy. Finish its current operation before rebuilding.'}
        $state=$current.body|ConvertFrom-Json
        if($state.data.editing){throw 'Save and return from Excel before rebuilding; the project workbook is still open for editing.'}
        $closed=Request-Cfl "$url/api/action/close" $state.token
        if(-not $closed -or $closed.status -ne '200'){throw 'The running application could not close safely. Use its Exit button, then rebuild.'}
        if($info -and $info.processId){
            try{$server=[Diagnostics.Process]::GetProcessById([int]$info.processId)}catch{$server=$null}
            if($server -and -not $server.WaitForExit(10000)){throw 'The application is still shutting down. Wait before rebuilding.'}
        }
        $executable=Join-Path $outputPath 'ZTools.Standalone.exe'
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        do{
            try{$file=[IO.File]::Open($executable,'Open','ReadWrite','None');$file.Dispose();$unlocked=$true}
            catch [IO.IOException]{$unlocked=$false;Start-Sleep -Milliseconds 200}
        }while(-not $unlocked -and [DateTime]::UtcNow -lt $deadline)
        if(-not $unlocked){throw 'The application files are still in use. Stop debugging before rebuilding.'}
        Write-Host "Closed the idle application safely before updating $outputPath"
    }
}finally{if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp -Force}}
& (Join-Path $PSScriptRoot 'retire-runtime-references.ps1') -OutputDirectory $outputPath
