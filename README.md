# CFL SubmitTools — standalone DWG application

The active application is C#/.NET 8 with a Chinese interface, the supplied CFL logo, and an offline Three.js/WebGL drawing viewer. It reads the original DWG templates directly and generates editable DWG files and vector PDF files without starting AutoCAD. No DXF conversion is involved.

The [Chinese user guide (PDF)](output/pdf/凯乐提交文件工具_用户操作指南_中文版.pdf) covers installation, the XLSX project workflow, parameter dialogs, SEC material rules, drawing output and troubleshooting. It includes the CFL logo, an interface map, quick steps and a mobile appendix. The Windows package installs it as `CFL-User-Guide-zh-CN.pdf`. Its source is `tools/docs/build_user_guide.py`; regeneration requires Python with `reportlab` and `pypdf` and uses the bundled licensed Chinese font. PDF previews under `tmp/pdfs/` are temporary and excluded from Git.

Native Android and iPhone projects also share a portable offline engine that uses the embedded XLS model and reads/writes XLSX input projects without Excel. Android APK and the iPhone Mac build kit are prepared under `dist/CFL-Mobile`. See [MOBILE-INSTALL.md](MOBILE-INSTALL.md) for installation, Mac signing and validation limits. The phone projects are separate from the Windows application.

Mobile version 1.1 (build 2) enables **生成 DWG＋PDF** after a successful completed calculation and adds live progress and cooperative cancellation on both phone platforms. The Android installer and iPhone Mac build kit are refreshed together; the build kit is not a signed IPA.

## Build and run

Requires Windows x64 and a WebGL-capable browser. Editing in Excel requires Microsoft Excel desktop. The setup package includes the .NET runtime; the ordinary framework-dependent folder build requires the .NET 8 ASP.NET Core and Windows Desktop runtimes. Building requires the .NET 8 or 9 SDK.

### Windows setup package

Run `powershell -NoProfile -ExecutionPolicy Bypass -File tools/build.ps1` with Inno Setup 6.5 or later installed. This builds and publishes the self-contained `win-x64` application, verifies all five embedded templates, and regenerates `dist/CFL-Setup/CFL-SubmitTools-Setup-1.0.0-win-x64.exe`, its SHA-256 file and a source-commit/configuration manifest in the same command. A custom compiler path can be supplied with `-InnoSetupCompiler`; use `-Version` for a subsequent release number. `tools/build-setup.ps1` remains a direct entry point to the same packaging pipeline. `tools/build.ps1 -SkipSetup` explicitly requests only the framework-dependent application folder. Packaging errors fail the build.

The Chinese installer uses the supplied CFL logo, installs for the current user under `%LOCALAPPDATA%/Programs/CFL/SubmitTools`, and creates Start menu and optional desktop shortcuts plus an uninstaller. Windows 10 1809 or later is required. Save Excel edits and exit the app before upgrading. Project workbooks should be stored in Documents or another project folder. Uninstall removes installed files, leaving user-created project files. No publisher signing certificate is configured, so this setup executable is unsigned. `tests/setup.ps1` verifies installation, startup, upgrade, and uninstall in a private folder; it refuses to replace an existing registered installation.

### Visual Studio and GitHub

Open `ZTools.sln` directly from this repository's local folder in Visual Studio 2022 (17.12 or later when using .NET 9). The solution references the local C# projects by relative path. **Build > Rebuild Solution** (or rebuilding `ZTools.Standalone`) now regenerates both the application folder and setup executable after successful compilation, using the selected Debug/Release configuration and project version. Inno Setup 6.5 or later must be installed. The packaging publish disables this rebuild hook internally to prevent recursion. Ordinary incremental Build/F5 continues to build the development output. The existing launch profile uses port 5100. The `WindowsSetup` publish profile produces the self-contained application folder. Packaging files are also visible in the solution's **Windows Setup** folder.

The equivalent command-line rebuild is `dotnet build ZTools.sln -t:Rebuild -c Release`. For an explicit application-only rebuild, pass `-p:BuildWindowsSetupOnRebuild=false`. A custom installer compiler path may be passed as `-p:InnoSetupCompiler=...`. GitHub Actions uses the combined build command, tests installation/upgrade/uninstall and uploads **CFL-SubmitTools-Windows-Setup** alongside the application-folder artifact. This generates an installer; installation on a user's computer is a separate action.

Visual Studio's Git Changes / Git Repository windows use this folder's `.git` repository and `origin` at `https://github.com/paciguard/SDOCtools.git`. Pull before editing, then commit and push checked changes on `main`. No separate source copy or legacy solution source-control binding is needed. Build products stay under ignored `dist/` and are not committed to GitHub.

Run `powershell -ExecutionPolicy Bypass -File tools/build.ps1`, then launch `dist/CFL-SubmitTools/Start.ps1` or `ZTools.Standalone.exe` inside that published folder. Keep the complete application folder, including its libraries and `wwwroot`. GA.dwg, PID.dwg, SEC.dwg and the split-workbook model are embedded in the application assembly; no external business-template files or `Referrings` folder are needed at runtime. The interface opens at http://127.0.0.1:5099. Third-party JavaScript is bundled locally; drawing generation does not require an internet connection.

Repeated launches reopen the current application and preserve its project session. The build script closes an idle running copy of the same output folder before replacing its files. If Excel is open for input, save and return to the workbench first; rebuild will refuse to interrupt editing.

In Visual Studio 2022, open `ZTools.sln`, set `ZTools.Standalone` as the startup project, and run its **ZTools.Standalone** profile. That profile uses http://127.0.0.1:5100 so it can run alongside the published application. Regular build output includes the embedded templates and `wwwroot`; keep the whole output folder together. If startup fails, the app shows a Chinese error and writes `application-error.log` beside the executable (or under `%LOCALAPPDATA%/CFL/SubmitTools` if the folder is not writable). Custom addresses use `--port <number>`; close those instances with **退出程序** before rebuilding.

## Workflow

1. Choose **新建项目** to create an editable `.xlsx` workbook containing only **input** in your selected project folder, or **打开项目** to select an existing saved project workbook. Starting the application never recreates or overwrites an existing project, and Excel opens only when requested.
2. Use **在 Excel 中输入数据** for the original input template. All database and calculation sheets remain inside the application. Return with **保存并返回工作台** before running other operations.
   You can also save and close Excel normally. The workbench detects the closed editor, reloads the last saved project file, and enables its controls again. A Chinese notice explains that unsaved changes cannot be recovered. If Excel reports a genuine save error, its editor remains open so you can fix the problem and retry.
3. Select the equipment row and use **设备参数**, **电机与底座**, **管仪参数**, **剖面材料**, **公用工程参数**, or **项目信息**. Internal PID/Sec-List/Utility values remain editable through these dialogs and travel with the input-only project. **查看内部计算结果** shows calculated values without exposing a worksheet. Changed parameters require recalculation before drawing creation.
4. Click **计算项目数据** after opening a project or changing equipment inputs. Generation is blocked until calculation succeeds. Select devices and drawing types, then click **生成 DWG ＋ PDF**.
5. Preview the result, zoom with the mouse wheel, pan by dragging, and download either format. A3 landscape PDF window, margins, and copy count are stored as a custom property of the project Excel workbook.

After a drawing batch is successfully saved, **图纸生成完成** displays its output folder. Choose **确定** to return to the workbench. Android/iPhone show the same completion dialog with the saved DWG/PDF count. Failed creation does not show a success prompt.

Results are stored in `<project>/Output-Standalone/<timestamp>/GA|PID|SEC`. Each drawing has a DWG, PDF, and a temporary scene JSON used by the viewer. Import notices are recorded in `compatibility.log`. Use **退出程序** when finished to close the local drawing service. Closing only the browser tab leaves the service running. The source Excel and DWG templates are never used as writable job files. Excel VBA macros are disabled in the app's private Excel instance. Close separately opened project copies before opening them in the app.

## Drawing and calculation components

- **ACadSharp 3.8.0** reads DWG, inserts blocks, edits attributes/text/layers, and writes DWG in AC1032 format (AutoCAD 2018 generation). This is an independent .NET library, not Autodesk ObjectARX or RealDWG.
- **C# calculation engine** retains the migrated workbook and GA/PID/SEC drawing methods.
- **Three.js 0.180.0** displays the generated geometry and attribute text through WebGL.
- **PDFsharp-GDI 6.2.4** writes monochrome vector A3 PDFs. Chinese text uses a suitable locally installed TrueType font; no Windows font files are redistributed.
- **Managed calculation** reconstructs the embedded XLS database model and saves XLSX input projects without starting Excel. Excel automation remains available for Windows editing, parameter dialogs and drawing generation; there is no replacement business database.

The build embeds the original `Referrings/Ztools/TEMPLATE/GA.dwg`, `PID.dwg`, `SEC.dwg` byte for byte, plus the derived `Resources/Workbooks/ZTools.xls` model and clean `Resources/Workbooks/Input.xlsx` input template. The newer supplied workbook is preserved unchanged at `Referrings/Updates/ZTools-20261026.xls`; the earlier source is also unchanged. The derived model removes `map` and contains the internal lookup tables, calculation sheets and form defaults. Public project workbooks contain only `input`, retain its formatting/dropdowns, and carry project parameters/results in compressed document properties that survive Excel save/reopen. They contain no VBA. Windows reconstructs a private internal workbook for Excel-backed operations; the portable engine reconstructs it in memory. Macros remain disabled. Opening an old full workbook creates a separate input-only project copy and preserves the old file. Updating a built-in template requires rebuilding the application. Preserve original block names, attributes, and layers. The app is a template-driven drawing tool; it does not provide a general AutoCAD editor. It preserves CAD content supported by the DWG library. Non-graphical legacy object warnings are logged; unsupported visible PDF entity types stop export instead of silently disappearing. Visual fidelity for new templates or untested block combinations must be checked before issuing drawings.

## GA database selection

For pump family **ZE** with standard **7版** (or the legacy **7th** value), GA dimensions and weights come from worksheet **ZE7**, matched to the pump model and baseplate, including the voltage variant. Other ZE editions use **ZE**. Each equipment row selects its own sheet, so mixed editions can share one project. Missing ZE7 records stop calculation rather than falling back to ZE. The GA drawing inserts the existing **ZE7** main block from **GA.dwg**; all auxiliary block selections follow the same path as ZE8. The original Excel template and database are preserved.

## Input-only projects and internal parameters

`tools/WorkbookTemplates/prepare-input-xlsx.ps1` prepares the embedded `Resources/Workbooks/Input.xlsx` from the owner’s supplied `Referrings/CFL资料工具输入模板.xlsx`. The original is preserved unchanged. Unused legacy defined names that prevented desktop Excel from opening the supplied file are removed from the embedded copy; cell/validation/print dependencies are retained. Sheet XML, styles and all other package parts remain byte-identical. Both hashes and the removed names are recorded in `migration/input-xlsx-20261027.json`. New projects use true Office Open XML (`.xlsx`), including the supplied layout and dropdowns. The internal model remains the original XLS database. Legacy full or input-only `.xls` files import to separate `.xlsx` copies, retaining internal project parameters and preserving their source files. Existing mobile XLS projects migrate beside their original file when opened. `tools/WorkbookTemplates` retains the database model generator and creates an XLSX example; `prepare-input.ps1` and `Input.xls` remain historical migration references. Both `Resources/Workbooks/Input.xlsx` and `Resources/Workbooks/ZTools.xls` are embedded. Normal builds use these checked-in templates and do not regenerate them. `migration/workbook-split-20261026.json` retains the earlier XLS source/model/template hashes; the newer record covers the supplied XLSX template. The application never reads database sheets from the selected external project; it uses the embedded model. Existing full workbooks import their input values and internal project overrides, rather than replacing the application database.

`Sec-List!F10:O39` controls section/material output. Column F selects the material grade; G:M are per-part overrides; N controls bearing type/SEC list selection and defaults to `input!AA`; O supplies bearing-house material for **330.01, 330.02 and 330.03**. When specified, that same material also applies to **360.01 and 360.02**. Blank part-material overrides retain the selected database grade. Invalid grades, missing lists and missing parts are reported rather than guessed. Material changes affect the actual saved SEC DWG and vector-PDF scene. The original source workbook is retained for traceability.

## SEC calculation behavior

The calculations address the worksheet **Sheet1 by name**, including when sheets are reordered. `Sheet1!CA40:CA69` contains the actual SEC drawing-block name, including required `_K6` and `WO` variants. `CC40:CC69` independently contains the exact matching material-list header from **List-all row 2**, searched through the last populated header. Missing and duplicate matches produce an error. Material quantities/descriptions come from the original workbook, not invented defaults.

Known source-data gap: the `ZAG(LK6)_QWL` list references part `411.20`, absent from the supplied `List-detail` sheet. That variant stops with an explicit error until the database is completed by its owner.

Calculation uses the portable XLS engine without starting Excel. Progress is displayed while it reads the workbook, queries each pump, and saves the results. The same engine serves Android and iPhone. Database ranges are read once per calculation and formulas are evaluated once per read. A six-pump Windows check returned in 1.7 seconds, versus about 20 seconds previously; these are local PC measurements, not phone timing guarantees. After the completed result returns to the workbench, **生成 DWG ＋ PDF** becomes enabled. **取消计算** stops a running calculation without changing the project workbook. A calculation with no progress for 60 seconds, or over three minutes in total, stops automatically and returns to the workbench. Calculation uses a private workbook copy and a separate hidden process; it updates the project only after successful save/close and refuses to overwrite a project changed elsewhere during calculation. Excel editing and Windows drawing generation remain available as before.

A missing pump/baseplate record reports the selected database, model, baseplate and input row before Excel's generic VLOOKUP error can hide the cause. For example, `ZE80-2200 / 160L` with standard `7版` requires that exact record in `ZE7`; an existing record in `ZE` does not satisfy it. Confirm the intended edition and supply the correct database record rather than changing editions merely to bypass an error. See [the calculation investigation](migration/calculation-response-20261026.md).

## Project layout and checks

- `ZTools.Standalone/`: active application, independent of Autodesk assemblies.
- `ZTools.Standalone/Core/`: reusable calculation and drawing rules migrated from VBA.
- `Referrings/`: original workbook, DWG/DVB, EXB, and image references.
- `migration/legacy/`: extracted VBA for traceability; never executed by the new app.
- `migration/archived-autocad-addin-20261023.zip`: recovery source for the superseded add-in.
- `tools/verify_sources.py`: verifies original-reference hashes against the selected snapshot. Current source snapshot: `migration/source-manifest-20261023-embedded.json`, adopting the owner's current Excel template for embedding; earlier snapshots remain historical records.
- `tools/verify_project.py`: checks standalone wiring, original DWGs and the published internal model.
- `tests/xlsx-project.ps1`: verifies native Excel XLSX format, supplied formatting/dropdowns, SEC output and internal settings after Excel save/reopen.
- `tests/standalone-api.ps1`: exercises Excel-backed forms, parameter and print-setting persistence, calculation, and two complete GA/PID/SEC output sets through the app.
- `tests/calculation-response.ps1`: requires desktop Excel; checks a missing ZE7 record on a private workbook copy, responsive progress, real calculation cancellation with byte-identical original data, and successful six-row calculation afterward. Pass `-Workbook <path>` to reproduce a supplied project without editing it.
- `tests/CalculationProcessSmoke`: checks deliberately stalled workers, idle/total timeouts, cancellation, unexpected exit, concurrent project edits and atomic result publication without requiring Excel.
- `tests/startup.ps1`: checks detached executable launch, repeat launches, safe rebuild preparation, the Visual Studio address, and startup error reporting without opening Excel.
- `tests/embedded-templates.ps1`: runs a package outside the checkout with no external templates, verifies all five embedded hashes, reads/writes the three DWGs, and checks input-only project creation backed by the embedded model without requiring Office.
- `tests/EditorLifecycleSmoke.csproj`: requires desktop Excel; checks save-and-return, saved/closed workbook recovery, disconnected Excel, save-error retry, and preservation of another open workbook, using private workbook copies.
- `tests/DrawingSmoke.csproj`: regenerates vector PDFs from six saved DWG files for visual regression review.
- `tests/WorkbookSmoke.csproj`: detailed workbook/SEC regression checks; requires the .NET Framework 4.8 targeting pack.

The application additionally reopens every saved DWG to check entity counts, equipment tags, and the first SEC material-list part. Sample PDF pages are rendered for visual review. These checks do not establish fidelity for every possible input/model combination.

## GitHub synchronization

The canonical repository is https://github.com/paciguard/SDOCtools. The owner's instruction to update both the local project and GitHub after future completed changes is recorded in `AGENTS.md`. Source, required reference files, licenses, and migration records are versioned. Build products, tool caches, credentials, and generated project drawings are excluded. Preserve remote history and never force-push unrelated work.
