param([string]$Workbook,[int]$Port=5099)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$temp = Join-Path $repoRoot '.tools'
function Read-Api([string]$Endpoint) {
    $response = Join-Path $temp 'api-test-response.json'
    & curl.exe --noproxy '*' --silent --show-error --fail-with-body --max-time 180 "http://127.0.0.1:$Port$Endpoint" --output $response
    if($LASTEXITCODE -ne 0){throw (Get-Content $response -Raw)}
    return Get-Content $response -Raw | ConvertFrom-Json
}
$initial = Read-Api '/api/state'
$token = $initial.token
function Call-Api([string]$Command,$Body) {
    $request = Join-Path $temp 'api-test-request.json'
    $response = Join-Path $temp 'api-test-response.json'
    [IO.File]::WriteAllText($request,($Body|ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
    & curl.exe --noproxy '*' --silent --show-error --fail-with-body --max-time 180 -H "X-CFL-Token: $token" -H 'Content-Type: application/json' --data-binary "@$request" "http://127.0.0.1:$Port/api/action/$Command" --output $response
    if($LASTEXITCODE -ne 0){throw (Get-Content $response -Raw)}
    return Get-Content $response -Raw | ConvertFrom-Json
}
if($Workbook){$state=Call-Api 'open' @{path=$Workbook}}else{$state=$initial.data}
if($state.rows.Count -ne 6){throw 'Expected six supplied equipment rows'}
foreach($group in 'equipment','motor','pid','materials','project','utility','results'){
    $form=Call-Api 'parameters' @{group=$group;row=0}
    if($form.fields.Count -lt 5){throw "Incomplete parameter form: $group"}
    if($group -eq 'results' -and @($form.fields|Where-Object {-not $_.readOnly}).Count){throw 'Calculated result fields must be read-only'}
    if($group -eq 'materials'){
        foreach($address in 'F10','N10','O10'){
            $entry=$form.fields|Where-Object {$_.address -eq $address}
            if(-not $entry -or $entry.readOnly){throw "Missing editable Sec-List field: $address"}
            if($address -ne 'O10' -and $entry.choices.Count -lt 2){throw "Missing Sec-List choices: $address"}
        }
    }
    $form|ConvertTo-Json -Depth 12|Set-Content (Join-Path $temp "form-$group.json") -Encoding UTF8
    $choices=($form.fields|Where-Object {$_.choices.Count -gt 0}).Count
    Write-Output "PASS $group fields=$($form.fields.Count) selectionFields=$choices"
}
$equipment=Get-Content (Join-Path $temp 'form-equipment.json') -Raw|ConvertFrom-Json
$field=$equipment.fields|Where-Object {$_.address -eq 'D21'}
$original=$field.value
try {
    $change=@{sheet=$field.sheet;address=$field.address;value='CFL-API-ROUNDTRIP'}
    $saved=Call-Api 'save-parameters' @{group='equipment';row=0;changes=@($change)}
    if($saved.rows[0].tag -ne 'CFL-API-ROUNDTRIP'){throw 'Parameter write did not persist to Excel'}
    Write-Output 'PASS equipment parameter writes to the project Excel workbook'
} finally {
    $change=@{sheet=$field.sheet;address=$field.address;value=$original}
    $restored=Call-Api 'save-parameters' @{group='equipment';row=0;changes=@($change)}
    if($restored.rows[0].tag -ne $original){throw 'Parameter restoration failed'}
}
$originalOptions=$state.options
try {
    $saved=Call-Api 'print' @{options=@{minX=0;minY=0;maxX=420;maxY=297;margin=7;copies=2}}
    if($saved.options.margin -ne 7 -or $saved.options.copies -ne 2){throw 'Print settings did not persist'}
    Write-Output 'PASS print settings persist in the project workbook'
} finally {$null=Call-Api 'print' @{options=$originalOptions}}
$null=Call-Api 'calculate' @{}
Write-Output 'PASS current Excel inputs recalculated before drawing generation'
$generated=Call-Api 'generate' @{kinds=@('GA','PID','SEC');rows=@(0,5)}
$files=Read-Api '/api/files'
if($files.Count -ne 6){throw 'Expected six drawings for two equipment rows'}
Write-Output "PASS six DWG/PDF drawing pairs generated without AutoCAD: $($generated.output)"
Write-Output 'STANDALONE_API_OK'
