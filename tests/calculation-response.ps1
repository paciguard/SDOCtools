param([string]$Workbook,[int]$Port=5193)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$folder=Join-Path $PSScriptRoot ('calculation-output/'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder -Force|Out-Null
$source=if($Workbook){[IO.Path]::GetFullPath($Workbook)}else{Join-Path $root 'Referrings/Ztools/ZTools.xls'}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$copy=Join-Path $folder 'ZTools.xls'
Copy-Item -LiteralPath $source -Destination $copy
$url="http://127.0.0.1:$Port"
$process=Start-Process (Join-Path $root 'dist/CFL-SubmitTools/ZTools.Standalone.exe') -ArgumentList "--no-browser --port $Port" -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $folder 'server.log') -RedirectStandardError (Join-Path $folder 'server.err')
$client=[Net.Http.HttpClient]::new([Net.Http.HttpClientHandler]@{UseProxy=$false})
$client.Timeout=[TimeSpan]::FromSeconds(60)
function Get-Json([string]$Endpoint){$client.GetStringAsync($url+$Endpoint).GetAwaiter().GetResult()|ConvertFrom-Json}
function Begin-Action([string]$Command,$Body){
    $content=[Net.Http.StringContent]::new(($Body|ConvertTo-Json -Depth 10),[Text.Encoding]::UTF8,'application/json')
    $client.PostAsync($url+'/api/action/'+$Command,$content)
}
function Action([string]$Command,$Body){
    $response=(Begin-Action $Command $Body).GetAwaiter().GetResult()
    $text=$response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    if(-not $response.IsSuccessStatusCode){throw $text}
    $text|ConvertFrom-Json
}
function Read-Data([string]$Path){
    $excel=New-Object -ComObject Excel.Application
    $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false
    try{
        $book=$excel.Workbooks.Open($Path,0,$true,5,'123','123')
        $snapshot=@()
        foreach($area in @(@('input','D21:AD50'),@('Sheet1','B40:CN69'),@('Sheet1','C15'),@('Sheet1','DB72:DE123'),@('Sec-List','F10:F39'))){
            $snapshot+=@{sheet=$area[0];range=$area[1];data=@($book.Worksheets.Item($area[0]).Range($area[1]).Formula)}
        }
        ConvertTo-Json -InputObject $snapshot -Depth 10 -Compress
    }finally{
        if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
        $excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
try {
    for($attempt=0;$attempt -lt 40;$attempt++){
        try{$initial=Get-Json '/api/state';break}catch{Start-Sleep -Milliseconds 250}
    }
    if(-not $initial){throw 'Isolated test server failed to start.'}
    $health=Get-Json '/api/health'
    if($health.processId -ne $process.Id){throw 'Test port belongs to another application.'}
    $client.DefaultRequestHeaders.Add('X-CFL-Token',$initial.token)
    $null=Action 'open' @{path=$copy}
    if(-not $Workbook){$null=Action 'save-parameters' @{group='equipment';row=0;changes=@(@{sheet='input';address='E21';value='7版'})}}
    # Excel updates OLE directory timestamps even without Save. Compare saved
    # inputs/output formulas instead of mistaking that metadata for a data edit.
    $before=Read-Data $copy
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $pending=Begin-Action 'calculate' @{}
    $progressChecks=0
    while(-not $pending.IsCompleted){
        $check=[Diagnostics.Stopwatch]::StartNew()
        $progress=Get-Json '/api/progress'
        if($check.Elapsed.TotalSeconds -gt 3){throw 'Progress was blocked behind Excel.'}
        $progressChecks++
        Start-Sleep -Milliseconds 200
    }
    $response=$pending.GetAwaiter().GetResult()
    $failure=$response.Content.ReadAsStringAsync().GetAwaiter().GetResult()|ConvertFrom-Json
    if([int]$response.StatusCode -ne 400 -or $failure.error -notmatch 'ZE7.*ZE80-2200.*160L.*21'){throw ('Missing lookup context: '+$failure.error)}
    $elapsed=$timer.Elapsed.TotalSeconds
    if((Read-Data $copy) -cne $before){throw 'Failed calculation changed saved inputs or output formulas.'}
    $state=Get-Json '/api/state'
    if(-not $state.data.needsCalculation){throw 'Failed calculation enabled drawing generation.'}
    Write-Output ("PASS missing ZE7 record returns HTTP 400 in {0:N1}s; model/baseplate/row reported; {1} progress requests remain responsive; saved data unchanged" -f $elapsed,$progressChecks)
    Write-Output $failure.error
    # Verify the same server can calculate a valid project after the failure.
    $validFolder=Join-Path $folder 'valid';New-Item -ItemType Directory -Path $validFolder|Out-Null
    $valid=Join-Path $validFolder 'ZTools.xls'
    Copy-Item -LiteralPath (Join-Path $root 'Referrings/Ztools/ZTools.xls') -Destination $valid
    $null=Action 'open' @{path=$valid}
    $result=Action 'calculate' @{}
    if($result.needsCalculation -or $result.rows.Count -ne 6){throw 'Valid project calculation did not recover.'}
    Write-Output 'PASS same STA queue calculates six valid rows after the failed project'
    $null=Action 'close' @{}
} finally {
    $client.Dispose()
    if(-not $process.WaitForExit(5000)){Stop-Process -Id $process.Id}
    if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Original workbook was changed.'}
}
