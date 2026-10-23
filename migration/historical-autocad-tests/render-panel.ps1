param([int]$Width = 370, [int]$Height = 780, [string]$OutputName = 'tools-panel.png')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$acadDirectory = 'C:\Program Files\Autodesk\AutoCAD 2021'
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $acadDirectory 'AcCoreMgd.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $acadDirectory 'AcDbMgd.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $acadDirectory 'AcMgd.dll'))
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$assembly = [System.Reflection.Assembly]::LoadFrom((Join-Path $projectRoot 'ZTools.AutoCAD\bin\Debug\ZTools.AutoCAD.dll'))
$panelType = $assembly.GetType('ZTools.AutoCAD.ToolsPanel',$true)
$panel = [Activator]::CreateInstance($panelType,$true)
$panel.Size = New-Object System.Drawing.Size($Width,$Height)
$panel.CreateControl()
$panel.PerformLayout()
$bitmap = New-Object System.Drawing.Bitmap($Width,$Height)
$panel.DrawToBitmap($bitmap,(New-Object System.Drawing.Rectangle(0,0,$Width,$Height)))
$bitmap.Save((Join-Path $PSScriptRoot $OutputName),[System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()
$panel.Dispose()
$optionsType = $assembly.GetType('ZTools.AutoCAD.PlotParameters',$true)
$options = [Activator]::CreateInstance($optionsType,$true)
$formType = $assembly.GetType('ZTools.AutoCAD.PlotSettingsForm',$true)
$constructor = $formType.GetConstructors([System.Reflection.BindingFlags]'Instance,NonPublic')[0]
$form = $constructor.Invoke(@($options))
$previewHost = New-Object System.Windows.Forms.UserControl
$previewHost.Size = $form.ClientSize
$previewHost.Font = $form.Font
$settingsLayout = $form.Controls[0]
$form.Controls.Remove($settingsLayout)
$previewHost.Controls.Add($settingsLayout)
$previewHost.CreateControl()
$previewHost.PerformLayout()
$settingsLayout.PerformLayout()
foreach ($child in $settingsLayout.Controls) { $child.CreateControl(); $child.PerformLayout() }
$preview = New-Object System.Drawing.Bitmap($previewHost.Width,$previewHost.Height)
$previewHost.DrawToBitmap($preview,(New-Object System.Drawing.Rectangle(0,0,$previewHost.Width,$previewHost.Height)))
$preview.Save((Join-Path $PSScriptRoot 'plot-settings.png'),[System.Drawing.Imaging.ImageFormat]::Png)
$preview.Dispose()
$previewHost.Dispose()
$form.Dispose()
Write-Host 'Tools panel rendered.'
