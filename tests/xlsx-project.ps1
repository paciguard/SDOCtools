param([int]$Port=55197,[switch]$SkipApi)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$folder=Join-Path $root ('.tools/xlsx-project-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder|Out-Null
$source=Join-Path $root 'Referrings/CFL资料工具输入模板.xlsx'
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$exe=Join-Path $root 'dist/CFL-SubmitTools/ZTools.Standalone.exe'
$child=Start-Process $exe -ArgumentList '--no-browser','--port',"$Port" -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $folder 'app.log') -RedirectStandardError (Join-Path $folder 'app.err')
$null=$child.Handle
$base="http://127.0.0.1:$Port"
$token=$null
function Call([string]$command,$body){Invoke-RestMethod "$base/api/action/$command" -Method Post -NoProxy -Headers @{'X-CFL-Token'=$token} -ContentType 'application/json' -Body ([Text.Encoding]::UTF8.GetBytes(($body|ConvertTo-Json -Depth 10)))}
try{
    for($i=0;$i -lt 40;$i++){
        try{$state=Invoke-RestMethod "$base/api/state" -NoProxy -TimeoutSec 2;$token=$state.token;if($token){break}}catch{}
        Start-Sleep -Milliseconds 250
    }
    if(-not $token){throw 'Private XLSX app did not start'}
    $state=Call 'new' @{path=$folder}
    $workbook=$state.job
    if([IO.Path]::GetExtension($workbook) -ne '.xlsx'){throw 'New project does not use XLSX'}
    if(-not $SkipApi){& (Join-Path $root 'tests/standalone-api.ps1') -Workbook $workbook -Port $Port}
    $null=Call 'save-parameters' @{group='materials';row=0;changes=@(@{sheet='Sec-List';address='O10';value='XLSX-BH-330'})}
    $null=Call 'calculate' @{}
    $generated=Call 'generate' @{kinds=@('SEC');rows=@(0)}
    $generated.output|Set-Content (Join-Path $folder 'bearing-output.txt')
    $files=Invoke-RestMethod "$base/api/files" -NoProxy
    $scene=Invoke-RestMethod "$base/api/file/$($files[0].id)/scene" -NoProxy
    if(($scene|ConvertTo-Json -Depth 12 -Compress) -notmatch 'XLSX-BH-330'){throw 'Saved SEC scene lost the XLSX bearing material'}
    foreach($part in '360.01','360.02'){
        $matched=$false
        foreach($number in @($scene.items|Where-Object text -eq $part)){
            foreach($material in @($scene.items|Where-Object text -eq 'XLSX-BH-330')){
                if([Math]::Abs($number.points[0][1]-$material.points[0][1]) -lt 0.01){$matched=$true}
            }
        }
        if(-not $matched){throw "Saved Windows SEC material row did not inherit the bearing-house material: $part"}
    }
    $excel=New-Object -ComObject Excel.Application
    $excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
    try{
        $original=$excel.Workbooks.Open((Join-Path $root 'Resources/Workbooks/Input.xlsx'),0,$true,5,'','')
        $book=$excel.Workbooks.Open($workbook,0,$false,5,'','')
        if($book.FileFormat -ne 51 -or $book.Worksheets.Count -ne 1 -or $book.Worksheets.Item(1).Name -ne 'input'){throw 'Excel did not recognize a true input-only XLSX'}
        $input=$book.Worksheets.Item('input');$reference=$original.Worksheets.Item('input')
        foreach($address in 'F10','D21','N21','O21','I53'){
            $actual=$input.Range($address);$expected=$reference.Range($address)
            if($actual.Interior.Color -ne $expected.Interior.Color -or $actual.Font.Name -ne $expected.Font.Name -or $actual.RowHeight -ne $expected.RowHeight -or [Math]::Abs($actual.ColumnWidth-$expected.ColumnWidth) -gt 0.1 -or $actual.MergeArea.Address() -ne $expected.MergeArea.Address()){throw "Supplied template formatting changed at $address"}
        }
        if($input.Range('AA21').Validation.Type -ne 3 -or $input.Range('L15').Formula -ne '=TODAY()'){throw 'XLSX dropdown or date formula was lost'}
        $input.Range('D21').Value2='XLSX-EXCEL-ROUNDTRIP'
        $book.Save();$book.Close($false);$original.Close($false)
    }finally{$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}
    $state=Call 'open' @{path=$workbook}
    if($state.rows[0].tag -ne 'XLSX-EXCEL-ROUNDTRIP'){throw 'Native Excel edit did not survive reopen'}
    $form=Call 'parameters' @{group='materials';row=0}
    if(($form.fields|Where-Object address -eq O10).value -ne 'XLSX-BH-330'){throw 'Excel save lost internal material state'}
    if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Supplied reference was modified'}
    Write-Host 'PASS XLSX native Excel format/layout/dropdowns, app calculation and SEC output, internal state after Excel save/reopen, unchanged source'
    Write-Host "XLSX_PROJECT_OK $workbook"
}finally{
    if(-not $child.HasExited){
        try{if($token){$null=Call 'close' @{};if(-not $child.WaitForExit(15000)){throw 'Private app did not close'}}}
        finally{if(-not $child.HasExited){Stop-Process -Id $child.Id -Force}}
    }
}
