# Migration record

The active project is now the standalone `ZTools.Standalone` application. The earlier VSTO and AutoCAD/ObjectARX host directions were superseded by the owner's standalone C# / Three.js / WebGL request. The standalone engine reads original DWG templates directly. Excel remains the primary input and database, and all parameter dialogs save to project workbook copies.

The superseded AutoCAD add-in source is preserved in `archived-autocad-addin-20261023.zip`. A complete byte-verified local recovery archive, including the former deployment bundle, is also retained in the ignored tools cache. Old host-specific test scripts are historical reference only.

Read-only extraction recovered 38 Excel VBA modules and two modules from each of GA.dvb, PID.dvb and SEC.dvb. Original files remain under `../Referrings`; extracted VBA under `legacy` is traceability material, not runtime code.

| Original function | Active replacement |
|---|---|
| Workbook transfer, dimensions, weights and utilities | Core/Generated workbook methods coordinated by Core/SubmissionEngine.cs |
| GA/PID/SEC VBA drawing methods | Core/Generated drawing methods through CadAdapter.cs |
| SEC template and material selection | Core/SectionSelection.cs; exact List-all row 2 header matching |
| AutoCAD drawing API | ACadSharp DWG document, block, attribute, layer and entity APIs |
| Drawing display | Offline Three.js/WebGL scene |
| Plotting | DrawingOutput.cs vector PDF renderer |
| Main input and database | Original Excel workbook, copied per project |
| Intermediate parameter/selection forms | Chinese browser dialogs backed by the original Excel cells and validation lists |

Reference snapshots are deliberately separate: `source-manifest.json` records the earlier source state; `source-manifest-20261023.json` records the current owner-supplied workbook and 16 reference files. Do not restore the older workbook over the newer source.

The source SEC drawing was restored byte for byte from the owner's duplicate after an earlier AutoCAD console validation rewrote its input. Its current verified SHA-256 is `B09243A89B9FB6953DF0E9923F7C30BA1BE49DBBF431A7044946084281F5573A`. The new engine reads templates and writes only separate generated files.

Known database gap: `ZAG(LK6)_QWL` references part 411.20, missing from List-detail. The engine reports the missing data rather than inventing a material.

Historical validation logs refer to the retired AutoCAD host. Current standalone verification is documented in `standalone-verification.md`.
