# CFL SubmitTools — standalone DWG application

The active application is C#/.NET 8 with a Chinese interface, the supplied CFL logo, and an offline Three.js/WebGL drawing viewer. It reads the original DWG templates directly and generates editable DWG files and vector PDF files without starting AutoCAD. No DXF conversion is involved.

## Build and run

Requires Windows x64, Microsoft Excel desktop, a WebGL-capable browser, and the .NET 8 ASP.NET Core and Windows Desktop runtimes. Building requires the .NET 8 SDK; the SDK supplies those runtimes on the development computer.

Run `powershell -ExecutionPolicy Bypass -File tools/build.ps1`, then launch `dist/CFL-SubmitTools/Start.ps1` or `ZTools.Standalone.exe` inside that published folder. Keep its `Referrings` and `wwwroot` folders alongside the executable. The interface opens at http://127.0.0.1:5099. Third-party JavaScript is bundled locally; drawing generation does not require an internet connection.

## Workflow

1. Choose **新建项目** to make a complete project copy of the Excel workbook, or **打开项目** to select an existing project copy.
2. Use **在 Excel 中输入数据** for the original input template. All database sheets remain inside the workbook. Return with **保存并返回工作台** before running other operations.
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

Templates are `Referrings/Ztools/TEMPLATE/GA.dwg`, `PID.dwg`, and `SEC.dwg`. Preserve their original block names, attributes, and layers. The app is a template-driven drawing tool; it does not provide a general AutoCAD editor. It preserves CAD content supported by the DWG library. Non-graphical legacy object warnings are logged; unsupported visible PDF entity types stop export instead of silently disappearing. Visual fidelity for new templates or untested block combinations must be checked before issuing drawings.

## SEC calculation behavior

The calculations address the worksheet **Sheet1 by name**, including when sheets are reordered. `Sheet1!CA40:CA69` contains the actual SEC drawing-block name, including required `_K6` and `WO` variants. `CC40:CC69` independently contains the exact matching material-list header from **List-all row 2**, searched through the last populated header. Missing and duplicate matches produce an error. Material quantities/descriptions come from the original workbook, not invented defaults.

Known source-data gap: the `ZAG(LK6)_QWL` list references part `411.20`, absent from the supplied `List-detail` sheet. That variant stops with an explicit error until the database is completed by its owner.

## Project layout and checks

- `ZTools.Standalone/`: active application, independent of Autodesk assemblies.
- `ZTools.Standalone/Core/`: reusable calculation and drawing rules migrated from VBA.
- `Referrings/`: original workbook, DWG/DVB, EXB, and image references.
- `migration/legacy/`: extracted VBA for traceability; never executed by the new app.
- `migration/archived-autocad-addin-20261023.zip`: recovery source for the superseded add-in.
- `tools/verify_sources.py`: verifies original-reference hashes against the selected snapshot. Current source snapshot: `migration/source-manifest-20261023.json`.
- `tools/verify_project.py`: checks standalone wiring and the published reference files.
- `tests/standalone-api.ps1`: exercises Excel-backed forms, parameter and print-setting persistence, calculation, and two complete GA/PID/SEC output sets through the app.
- `tests/DrawingSmoke.csproj`: regenerates vector PDFs from six saved DWG files for visual regression review.
- `tests/WorkbookSmoke.csproj`: detailed workbook/SEC regression checks; requires the .NET Framework 4.8 targeting pack.

The application additionally reopens every saved DWG to check entity counts, equipment tags, and the first SEC material-list part. Sample PDF pages are rendered for visual review. These checks do not establish fidelity for every possible input/model combination.

## GitHub synchronization

The canonical repository is https://github.com/paciguard/SDOCtools. The owner's instruction to update both the local project and GitHub after future completed changes is recorded in `AGENTS.md`. Source, required reference files, licenses, and migration records are versioned. Build products, tool caches, credentials, and generated project drawings are excluded. Preserve remote history and never force-push unrelated work.
