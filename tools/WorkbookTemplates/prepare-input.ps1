$ErrorActionPreference='Stop'
$root=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$source=Join-Path $root 'Referrings\Updates\ZTools-20261026.xls'
$output=Join-Path $root 'Resources\Workbooks\Input.xls'
$intermediate=Join-Path $root '.tools\input-template.xlsx'
$excel=New-Object -ComObject Excel.Application
$excel.AutomationSecurity=3;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false
try{
    $original=$excel.Workbooks.Open($source,0,$true,5,'123','123')
    $original.Unprotect('123')
    $copy=$excel.Workbooks.Add(-4167)
    $blank=$copy.Worksheets.Item(1)
    $blank.Name='input'
    $sourceSheet=$original.Worksheets.Item('input')
    $range=$sourceSheet.Range('A1:AR143')
    $range.Copy($blank.Range($range.Address()))
    foreach($property in 'Orientation','PaperSize','Zoom','FitToPagesWide','FitToPagesTall','LeftMargin','RightMargin','TopMargin','BottomMargin','PrintArea','PrintTitleRows','PrintTitleColumns'){$blank.PageSetup.$property=$sourceSheet.PageSetup.$property}
    if($blank.Range('N21').Text -ne $sourceSheet.Range('N21').Text){throw 'Native input range copy failed.'}
    if($copy.Worksheets.Count -ne 1 -or $copy.Worksheets.Item(1).Name -ne 'input'){throw 'Excel did not create an input-only worksheet copy.'}
    for($i=$blank.Shapes.Count;$i -ge 1;$i--){$shape=$blank.Shapes.Item($i);if($shape.Type -in 8,12 -or [string]$shape.OnAction){$shape.Delete()}}
    foreach($link in @($copy.LinkSources(1))){if($link){$copy.BreakLink($link,1)}}
    # XLSX removes the legacy VBA module attached to the input sheet. Exporting
    # back to BIFF8 preserves the sheet's native styling, validation and images.
    $copy.SaveAs($intermediate,51,'','');$copy.Close($false)
    $copy=$excel.Workbooks.Open($intermediate,0,$false,5,'','')
    $copy.SaveAs($output,56,'','');$copy.Close($false)
    $original.Close($false)
    Write-Host "Prepared native input template: $output"
}finally{$excel.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)|Out-Null}
