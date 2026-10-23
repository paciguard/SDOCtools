# Methods, parameters and AutoCAD integration

## SEC name calculation correction (2026-10-23)

`SectionSelection.Resolve(model, revision, bearing, czBearingSize)` returns the
SEC base block name and a material-list search key independently. Revision
7/8/10, G assemblies, LK6 bearing frames and the five CZ balancing/shaft cases
have explicit branches. Missing combinations are not replaced with another list.

`CalculateSectionNames(pumpIndex)` searches the actual `List-all` row 2 from
column B to its last populated header, using a trimmed, case-insensitive exact
match. The matched cell's original text is returned to `Sheet1!CC40:CC69`.
Missing and ambiguous headers fail the calculation with a diagnostic. The parts
are read from that matched column, including columns added after BA.

`PrepareSection` resolves material/shaft overrides and finalizes `CA40:CA69`
with the actual main SEC block name, including `_K6` and `WO` where applicable.
`CE` retains the material class; `CF` retains the sleeve flag. The calculation
validates every pump's material list before saving, with rollback of names,
material flags and scratch part-list output on failure.

`BuildSEC` delegates main-block insertion to `InsertSection(point, row)`, which
uses `Sheet1!CA` directly. CZ inserts exactly one main block; ZE G assemblies
use the corresponding `ZE7G` or `ZEG` supplement. Original workbook calculation
references to ordinal worksheet 1 are replaced with the named `Sheet1` so
reordering worksheets cannot redirect output into a different sheet.

The original workbook and extracted VBA remain unchanged. The converter includes
these replacements so regenerating C# preserves the correction. Existing job
copies should be recalculated after updating the add-in.

The calculation and drawing rules come from the supplied VBA. AutoCAD-specific lifecycle, UI, command-context, transaction and export code is added in C#. The application uses managed ObjectARX (`AcCoreMgd`, `AcDbMgd`, `AcMgd`) and the AutoCAD ribbon API (`AdWindows`). Excel COM is used only to access the retained workbook.

## Workflow entry points

| VBA method | C# method | Parameters and result |
| --- | --- | --- |
| `tranfer_data()` | `SubmissionEngine.Calculate()` → `TransferData()` | Validates the job workbook, reads the original database, updates calculation cells, returns pump count. `TransferData` contains the migrated rules. |
| `GAcreat()` | `DrawingBatch.Generate("GA")` | Drawing kind `GA`; returns a new output-run directory. |
| `PIDCreate_For_Blocks()` | `PreparePid(int count)`, `Generate("PID")` | Number of pump rows; writes CG:CM plan components before drawing generation. |
| `PIDCreate_For_files()` | `PreparePid(int count)` | Original plan-name cleanup is shared with block generation. |
| `SecCreate()` | `PrepareSection(int k)`, `Generate("SEC")` | Zero-based pump index `k`; selects part list/materials and prepares DB:DE. |
| `Utility()` | `CalculateUtility()` | Uses current input rows and the original UtilityData lookup table. |
| AutoCAD `GA()`, `PID()`, `SEC()` | `BuildGA()`, `BuildPID()`, `BuildSEC()` | Operate on the engine's current workbook row and native drawing transaction. |
| `changeFld(fldName,newVal)` | `ChangeField(object name, object value)` → `CadDrawing.SetCustomProperty` | Property name and value from the original worksheet pairs; ignores blank names. |
| `UpdateProgress(Pct)` and progress forms | Command-line messages and palette log | Reports pump/type progress without VBA UserForms or application visibility changes. |

`SubmissionEngine(WorkbookSession workbook, Action<string> report)` binds one explicit workbook to the calculation/drawing rules. `DrawingBatch(WorkbookSession workbook, Action<string> report, PlotParameters plotParameters = null)` coordinates a drawing run and copies plot settings. `WorkbookSession(string path, bool visible)` opens a job copy in a private Excel instance with macros disabled. Neither workbook selection nor drawing generation relies on another application's active workbook.

## Plot parameters

The VBA signature in all three DVB files is:

```text
PlotFunction(PrinterName, Styles, MediaName, Copies,
             AutoMedia, AutoRotate, AutoClose, AutoFrame, AutoSave)
```

`PdfPlotter.Plot(Document document, string output, PlotParameters parameters = null)` uses an explicit native AutoCAD document, output PDF path and validated settings. The **Plot settings** button / `ZTOOLSPLOTSET` command edits these settings for the current session.

| C# property | Type | Default | VBA relationship / behavior |
| --- | --- | --- | --- |
| `PrinterName` | `string` | `DWG To PDF.pc3` | Original `PrinterName`; passed to `SetPlotConfigurationName`. |
| `Styles` | `string` | `acad.ctb` | Original `Styles`; blank falls back to `acad.ctb`. |
| `MediaName` | `string` | empty | Original `MediaName`; explicit canonical device media name when AutoMedia is false. |
| `Copies` | `int` | `1` | Original `Copies`; validated from 1 to 99, passed to `BeginDocument`. |
| `AutoMedia` | `bool` | `true` | Original A3 selection, resolved to an actual canonical ISO A3 medium offered by the device. |
| `AutoRotate` | `bool` | `true` | Fits plot-window orientation to the selected medium. Avoids rotating an already-landscape A3 medium twice. |
| `AutoFrame` | `bool` | `true` | Uses the explicit template window; false selects drawing extents. |
| `MinX`, `MinY` | `double` | `0`, `0` | Plot-window minimum, based on the original p point. |
| `MaxX`, `MaxY` | `double` | `420`, `297` | Plot-window maximum, based on the original q point. |

The old `AutoClose` and `AutoSave` flags are present in the VBA signature but do not control its final save/close calls. The migrated batch always writes the requested DWG/PDF documents and closes its own temporary drawing documents. It leaves the user's original AutoCAD documents and session open. These ownership rules are intentionally not exposed as misleading plot settings.

The original paper unit is millimeters, scale is fit-to-page, plotting is centered, lineweights and plot styles are enabled. The managed implementation retains those defaults. `BACKGROUNDPLOT` is restored after plotting. PID DWG extraction continues to use the original 0,0–420,297 A3 sheet window independently of an advanced PDF window override.

## Workbook cell contracts

Excel cell addresses are retained so the original formulas and database remain usable.

| Input range | Meaning |
| --- | --- |
| `input!D21:D50` | Equipment tags |
| `input!E21:E50` | Pump design/revision selection used in ZE7/ZE8/ZE10 rules |
| `input!F21:F50` | Service/equipment description |
| `input!N21:N50` | Pump model and consecutive-row count |
| `input!O21:O50` | Material class |
| `input!R21:R50` | Speed |
| `input!S21:W50` | Motor model, manufacturer, baseplate selection, voltage, power |
| `input!Y21:Z50` | Seal flush plan and cooling plan |
| `input!AA21:AA50` | Bearing/cooling form |
| `input!AB21:AB50` | Custom pump display prefix |
| `input!AC21:AC50` | Optional discharge/drain interface override |
| `input!I53` | Flange standard |
| `input!I56` | Nozzle-load multiplier used by the VBA |

| Calculation / lookup range | Meaning |
| --- | --- |
| `Sheet1!C3:C10` | Drawing revision, designer, checker, formatted date, customer, project and order fields |
| `Sheet1!C12:C14` | GB flange-pressure selections |
| `Sheet1!C15` | Calculated pump count |
| `Sheet1!C16:C18` | Mating flange, oil lubrication and drain options |
| `Sheet1!C19` | Current calculation row, restored after a drawing batch |
| `Sheet1!B40:CN69` | Pump calculations, dimensions, mass, nozzle loads and drawing parameters |
| `Sheet1!P40:R69` | GA, PID and section output names |
| `Sheet1!CE40:CF69` | Section material class and spacer/sleeve rule flag |
| `Sheet1!CG40:CM69` | Full PID plan and up to six component identifiers, stored as text |
| `Sheet1!DB72:DE111` | Up to 40 BOM rows: part number, description, material, quantity |
| `Sheet1!G4:H18`, `I4:J18`, `K4:L18` | GA custom-property name/value pairs |
| `CZ`, `ZA`, `ZE`, `ZE7`, `ZF` | Original pump dimension and weight databases |
| `Data1`, `CZ-BearingSize`, `nozzle-load` | Drain, bearing, interface-position and nozzle-load lookups |
| `List-all`, `List-detail`, `Sec-List` | Part-list selection, part/material details and per-order overrides |
| `UtilityData` | Original utility-flow lookup table |

Pump index `k` is zero-based: input row is `21+k`, calculation row is `40+k`, section/PID override row is `10+k`, and utility row is `13+k`. Lookup offsets remain one-based Excel offsets where used by VLOOKUP/HLOOKUP. Database matching explicitly checks pump size and baseplate, including the original `(6KV)` suffix rule.

## Drawing-operation parameters

| Ported operation | Parameters | Managed ObjectARX implementation |
| --- | --- | --- |
| `InsertBlock(point,name,sx,sy,sz,angle)` | `double[3]` point, block name, scale factors, radians | `BlockReference`, `Scale3d`, block-definition lookup, attribute-reference instantiation, transaction registration |
| `AddMText(point,width,content)` | Drawing point, width, text | `MText`, inherited template text style, native append/transaction; stale text at the same title anchor is replaced |
| `ScaleEntity(point,factor)` | Base point and scalar | `Entity.TransformBy(Matrix3d.Scaling(...))` |
| Attribute `TagString` / `TextString` | Tag name / formatted value | Native `AttributeReference.Tag` and `TextString` |
| Layer `LayerOn` | Boolean visibility | Native `LayerTableRecord.IsOff` inverse |
| Custom property name/value | Worksheet-provided strings | `DatabaseSummaryInfoBuilder.CustomPropertyTable` |
| PID sheet export | A3 window, target DWG path | `Editor.SelectWindow`, `Database.Wblock`, `SaveAs` |

The small `CadDrawing` surface lets the source-derived rules retain recognizable method names while executing native ObjectARX operations. It is not an AutoCAD COM object. Required missing blocks/layers produce named errors; optional empty components and explicitly absent no-interface/label-only components are handled separately.

## Additional AutoCAD lifecycle and UI methods

| Method / API | Parameters / purpose |
| --- | --- |
| `Commands.Initialize()` / `Terminate()` | `IExtensionApplication` lifecycle; register deferred UI creation and dispose owned UI. Headless tests skip UI. |
| `[CommandMethod(name, CommandFlags.Session)]` | Registers each named workflow; allows batch document operations in application context. |
| `UiHost.Initialize()` / `OnIdle(sender,args)` | Waits for ribbon availability and restores the tab after workspace changes without duplication. |
| `AddPanel(RibbonTab tab,string title)` | Creates a `RibbonPanelSource` and `RibbonPanel` in the ZTools tab. |
| `AddButton(source,label,command,glyph,hint)` | Creates a labeled `RibbonButton` with stable ID, 16/32-pixel vector icons, tooltip and command handler. |
| `CommandHandler.Execute(object parameter)` | `ICommand` callback; reads the button's registered command. |
| `UiHost.Enqueue(string command)` | Uses `Document.SendStringToExecute` so modeless palette/ribbon callbacks execute native work in command context. |
| `UiHost.ShowPalette()` | Creates/shows a `PaletteSet` with a stable GUID, left/right docking, close and auto-hide controls, and a minimum usable size. |
| `ToolsPanel.AddGroup(title,labels,commands)` | Builds responsive button groups; job path and operation log remain visible. |
| `Commands.SelectJob(bool force=false)` | Reuses the selected job unless an explicit Open job action requests another workbook. |
| `PlotSettingsForm(PlotParameters initial)` | Edits a copy of plotting parameters; validates before returning an applied result. |
| `Document.LockDocument()` | Locks only the owned drawing being modified. |
| `Database.TransactionManager.StartTransaction()` | Owns native DBObject opens, entity appends, attribute updates and commit/rollback. |
| `DrawingBatch.Generate(params string[] kinds)` | Accepts GA/PID/SEC; creates unique output folders, uses per-pump template copies, reports progress and writes a run log. |

The ribbon has four panels and 14 buttons. The dockable tools panel exposes 13 workflow/settings buttons; its counterpart ribbon button opens the panel itself. All production commands are also registered in `deploy/PackageContents.xml` for AutoCAD bundle loading.
