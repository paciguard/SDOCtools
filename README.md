# CFL SubmitTools — standalone DWG application

The active application is C#/.NET 8 with a Chinese interface, the supplied CFL logo, and an offline Three.js/WebGL drawing viewer. It reads the original DWG templates directly and generates editable DWG files and vector PDF files without starting AutoCAD. No DXF conversion is involved.

Native Android and iPhone projects also share a portable offline engine that reads/writes the original XLS workbook without Excel. Android APK and the iPhone Mac build kit are prepared under `dist/CFL-Mobile`. See [MOBILE-INSTALL.md](MOBILE-INSTALL.md) for installation, Mac signing and validation limits. The phone projects are separate from the Windows application.

## Build and run

Requires Windows x64, Microsoft Excel desktop, a WebGL-capable browser, and the .NET 8 ASP.NET Core and Windows Desktop runtimes. Building requires the .NET 8 SDK; the SDK supplies those runtimes on the development computer.

Run `powershell -ExecutionPolicy Bypass -File tools/build.ps1`, then launch `dist/CFL-SubmitTools/Start.ps1` or `ZTools.Standalone.exe` inside that published folder. Keep the complete application folder, including its libraries and `wwwroot`. GA.dwg, PID.dwg, SEC.dwg and the complete ZTools.xls input/database template are embedded in the application assembly; no external business-template files or `Referrings` folder are needed at runtime. The interface opens at http://127.0.0.1:5099. Third-party JavaScript is bundled locally; drawing generation does not require an internet connection.

Repeated launches reopen the current application and preserve its project session. The build script closes an idle running copy of the same output folder before replacing its files. If Excel is open for input, save and return to the workbench first; rebuild will refuse to interrupt editing.

In Visual Studio 2022, open `ZTools.sln`, set `ZTools.Standalone` as the startup project, and run its **ZTools.Standalone** profile. That profile uses http://127.0.0.1:5100 so it can run alongside the published application. Regular build output includes the embedded templates and `wwwroot`; keep the whole output folder together. If startup fails, the app shows a Chinese error and writes `application-error.log` beside the executable (or under `%LOCALAPPDATA%/CFL/SubmitTools` if the folder is not writable). Custom addresses use `--port <number>`; close those instances with **退出程序** before rebuilding.

## Workflow

1. Choose **新建项目** to extract a complete editable copy of the built-in Excel workbook into your selected project folder, or **打开项目** to select an existing saved project workbook. Starting the application never recreates or overwrites an existing project, and Excel opens only when requested.
2. Use **在 Excel 中输入数据** for the original input template. All database sheets remain inside the workbook. Return with **保存并返回工作台** before running other operations.
   You can also save and close Excel normally. The workbench detects the closed editor, reloads the last saved project file, and enables its controls again. A Chinese notice explains that unsaved changes cannot be recovered. If Excel reports a genuine save error, its editor remains open so you can fix the problem and retry.
3. Use the equipment, motor/baseplate, PID, and section-material dialogs to review or edit the selected equipment row. Selection lists come from the original Excel data validation. Changes are written into the same project workbook.
4. Click **计算项目数据** after opening a project or changing equipment inputs. Generation is blocked until calculation succeeds. Select devices and drawing types, then click **生成 DWG ＋ PDF**.
5. Preview the result, zoom with the mouse wheel, pan by dragging, and download either format. A3 landscape PDF window, margins, and copy count are stored as a custom property of the project Excel workbook.

Results are stored in `<project>/Output-Standalone/<timestamp>/GA|PID|SEC`. Each drawing has a DWG, PDF, and a temporary scene JSON used by the viewer. Import notices are recorded in `compatibility.log`. Use **退出程序** when finished to close the local drawing service. Closing only the browser tab leaves the service running. The source Excel and DWG templates are never used as writable job files. Excel VBA macros are disabled in the app's private Excel instance. Close separately opened project copies before opening them in the app.

## Drawing and calculation components

- **ACadSharp 3.8.0** reads DWG, inserts blocks, edits attributes/text/layers, and writes DWG in AC1032 format (AutoCAD 2018 generation). This is an independent .NET library, not Autodesk ObjectARX or RealDWG.
- **C# calculation engine** retains the migrated workbook and GA/PID/SEC drawing methods.
- **Three.js 0.180.0** displays the generated geometry and attribute text through WebGL.
- **PDFsharp-GDI 6.2.4** writes monochrome vector A3 PDFs. Chinese text uses a suitable locally installed TrueType font; no Windows font files are redistributed.
- **Excel automation** reads and updates the existing workbook and its embedded database; there is no replacement business database.

The build embeds the original `Referrings/Ztools/TEMPLATE/GA.dwg`, `PID.dwg`, `SEC.dwg`, and `Referrings/Ztools/ZTools.xls` byte for byte. DWGs are read directly from assembly resource streams, without extracting them to disk. Excel is extracted only as an editable project workbook, preserving all sheets, VBA content and the embedded database; macros remain disabled during app automation. Existing project workbooks are never replaced by the built-in template. Updating a built-in template requires rebuilding the application. Preserve original block names, attributes, and layers. The app is a template-driven drawing tool; it does not provide a general AutoCAD editor. It preserves CAD content supported by the DWG library. Non-graphical legacy object warnings are logged; unsupported visible PDF entity types stop export instead of silently disappearing. Visual fidelity for new templates or untested block combinations must be checked before issuing drawings.

## GA database selection

For pump family **ZE** with standard **7版** (or the legacy **7th** value), GA dimensions and weights come from worksheet **ZE7**, matched to the pump model and baseplate, including the voltage variant. Other ZE editions use **ZE**. Each equipment row selects its own sheet, so mixed editions can share one project. Missing ZE7 records stop calculation rather than falling back to ZE. The GA drawing inserts the existing **ZE7** main block from **GA.dwg**; all auxiliary block selections follow the same path as ZE8. The original Excel template and database are preserved.

## SEC calculation behavior

The calculations address the worksheet **Sheet1 by name**, including when sheets are reordered. `Sheet1!CA40:CA69` contains the actual SEC drawing-block name, including required `_K6` and `WO` variants. `CC40:CC69` independently contains the exact matching material-list header from **List-all row 2**, searched through the last populated header. Missing and duplicate matches produce an error. Material quantities/descriptions come from the original workbook, not invented defaults.

Known source-data gap: the `ZAG(LK6)_QWL` list references part `411.20`, absent from the supplied `List-detail` sheet. That variant stops with an explicit error until the database is completed by its owner.

Calculation progress is displayed while Excel opens, queries each pump, and saves the results. A missing pump/baseplate record reports the selected database, model, baseplate and input row before Excel's generic VLOOKUP error can hide the cause. For example, `ZE80-2200 / 160L` with standard `7版` requires that exact record in `ZE7`; an existing record in `ZE` does not satisfy it. Confirm the intended edition and supply the correct database record rather than changing editions merely to bypass an error. See [the calculation investigation](migration/calculation-response-20261026.md).

## Project layout and checks

- `ZTools.Standalone/`: active application, independent of Autodesk assemblies.
- `ZTools.Standalone/Core/`: reusable calculation and drawing rules migrated from VBA.
- `Referrings/`: original workbook, DWG/DVB, EXB, and image references.
- `migration/legacy/`: extracted VBA for traceability; never executed by the new app.
- `migration/archived-autocad-addin-20261023.zip`: recovery source for the superseded add-in.
- `tools/verify_sources.py`: verifies original-reference hashes against the selected snapshot. Current source snapshot: `migration/source-manifest-20261023-embedded.json`, adopting the owner's current Excel template for embedding; earlier snapshots remain historical records.
- `tools/verify_project.py`: checks standalone wiring and the published reference files.
- `tests/standalone-api.ps1`: exercises Excel-backed forms, parameter and print-setting persistence, calculation, and two complete GA/PID/SEC output sets through the app.
- `tests/calculation-response.ps1`: requires desktop Excel; checks a missing ZE7 record on a private workbook copy, responsive progress, unchanged saved inputs/output formulas, and successful six-row calculation afterward. Pass `-Workbook <path>` to reproduce a supplied project without editing it.
- `tests/startup.ps1`: checks detached executable launch, repeat launches, safe rebuild preparation, the Visual Studio address, and startup error reporting without opening Excel.
- `tests/embedded-templates.ps1`: runs a package outside the checkout with no external templates, verifies all four embedded hashes, reads/writes the three DWGs, and checks byte-identical Excel project extraction without requiring Office.
- `tests/EditorLifecycleSmoke.csproj`: requires desktop Excel; checks save-and-return, saved/closed workbook recovery, disconnected Excel, save-error retry, and preservation of another open workbook, using private workbook copies.
- `tests/DrawingSmoke.csproj`: regenerates vector PDFs from six saved DWG files for visual regression review.
- `tests/WorkbookSmoke.csproj`: detailed workbook/SEC regression checks; requires the .NET Framework 4.8 targeting pack.

The application additionally reopens every saved DWG to check entity counts, equipment tags, and the first SEC material-list part. Sample PDF pages are rendered for visual review. These checks do not establish fidelity for every possible input/model combination.

## GitHub synchronization

The canonical repository is https://github.com/paciguard/SDOCtools. The owner's instruction to update both the local project and GitHub after future completed changes is recorded in `AGENTS.md`. Source, required reference files, licenses, and migration records are versioned. Build products, tool caches, credentials, and generated project drawings are excluded. Preserve remote history and never force-push unrelated work.
