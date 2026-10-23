# Standalone verification — 2026-10-23

The application was built and published as a Windows x64 C#/.NET 8 program without Autodesk assembly references. AutoCAD was not started for these standalone checks. Original DWG files were read directly; no DXF conversion was used.

Passed checks:

- All 16 original reference files match the SHA-256 hashes and sizes in source-manifest-20261023.json. The same files in the published package match those hashes.
- All four Excel-backed dialogs load the original parameter values and selections: equipment 22 fields / 4 selection lists; motor and baseplate 5 / 2; PID 6 / 1; section materials 8 / 3.
- An equipment tag was changed through the app, read back from the saved project workbook, and restored to its original value.
- A3 PDF margin and copy count were saved to the project workbook, read back, and restored.
- The six supplied equipment rows calculated successfully. Current SEC block and material-list results for the first ZE row are ZE8_QWLWO and ZE8_QWL; the ZA row uses ZA_QWL. Material headers are matched against List-all row 2, using Sheet1 by name.
- Two selected equipment rows generated six DWG/PDF pairs: GA, PID, and SEC for each. The published executable passed the same open/calculate/generate flow.
- Every generated DWG was reopened and checked for entity count and equipment-tag value. The SEC first-part attribute was checked against Sheet1's calculated material data. PDF scenes include the equipment tags.
- The drawing regression utility read all six saved DWGs and regenerated vector PDFs and viewer scenes. Its final result was DRAWING_REVIEW_OK.
- GA, PID, and SEC PDF pages were rendered and visually inspected for geometry, Chinese text, attribute values, material tables, paragraph wrapping, title blocks, and sheet borders. Drawing text width factors and horizontal anchors are preserved while retaining stored text baselines.
- The browser displayed a generated drawing through WebGL. The Chinese motor/baseplate parameter dialog showed Excel-backed values and selection fields. Screenshots are under ../docs/images.
- The retired AutoCAD project, deployment files, and old package were preserved in a complete local archive verified byte for byte against all 76 archived files before removal. A source-only recovery ZIP is versioned in migration.

Coverage limits: these standalone output checks cover the two selected supplied equipment configurations, not every model and template combination. The viewer/PDF output is monochrome and uses an installed Chinese font. Import notices for legacy non-graphical objects are recorded per output batch; unsupported visible PDF entities stop export. The original database still lacks part 411.20 referenced by ZAG(LK6)_QWL; that case reports a data error.

Earlier workbook and native AutoCAD regression records remain historical evidence for the calculation migration; they are not presented as standalone drawing-engine validation.

Startup and rebuild follow-up — 2026-10-23:

- Reproduced the executable crash by launching a second instance while the original service occupied port 5099. The old executable exited with code -532462766 and an unhandled address-in-use exception.
- The corrected executable identifies and reopens the existing application. The second process exits with code 0; the original process and session token remain unchanged.
- Both normal Release build output and the published package launched from private copied folders with a different working directory. Their web assets loaded and their DWG paths resolved to the copied Referrings folder.
- A real Visual Studio 2022 MSBuild Rebuild passed with warnings treated as errors while the Release executable was running on port 5100. Build preparation closed that idle copy before cleaning or replacing its files.
- The publication script also closed its own idle running published executable before replacing the package. It identifies the output directory and refuses to interrupt an open Excel editor. Startup tests verified that preparation leaves a different application copy running.
- The Visual Studio address was honored; invalid startup settings produced a controlled exit code 1 and a readable error log. Startup tests run in Windows CI without opening Excel.
- Original reference hashes remain unchanged. Calculation and drawing-engine source were not changed in this follow-up.

Excel editor recovery follow-up — 2026-10-23:

- Reproduced the reported RPC_E_DISCONNECTED (0x80010108) at WorkbookSession.Save after saving and closing Excel, followed by Workspace.Finish. Old disposal also threw, leaving the editing flag set and other controls disabled.
- Desktop-Excel regression tests now pass for manual save followed by return, save/close followed by return, automatic detection of a closed workbook, and a disconnected test-owned Excel process. Recovery reloads the last saved project file and explicitly warns that unsaved edits cannot be recovered.
- A simulated disk-full save failure preserves the live editor for retry. Returning from the project also leaves another workbook opened in that Excel instance intact. Repeated disposal succeeds without rethrowing stale COM errors.
- Browser verification used the Visual Studio Debug application on port 5100. After the private test workbook was saved and closed, the Chinese notice appeared automatically, the saved equipment tag was refreshed, and input/calculation/parameter controls were enabled. The screenshot is docs/images/excel-editor-recovery.png.
- The same recovered Debug session passed all four parameter dialogs, input/print-setting save-and-restore checks, calculation of the six equipment rows, and generation of six GA/PID/SEC DWG/PDF pairs for two selected rows.
- Both Release and Debug application builds succeeded, the .NET Framework 4.8 calculation harness still compiles, and startup regression checks pass. Windows CI compiles the new editor harness; executing its Excel scenarios requires a machine with desktop Excel.
- The source ZTools.xls already had a local owner modification when this follow-up began. That file is preserved as found, and is excluded from the code-fix commit; it is not reset to the earlier repository snapshot. All 16 current reference files are checked against a separate before/after snapshot for this fix.

Embedded-template follow-up — 2026-10-23:

- GA.dwg, PID.dwg, SEC.dwg and the complete current ZTools.xls are now compiled as assembly resources. All four embedded byte lengths and SHA-256 hashes match their original files. The workbook retains its complete input sheets, VBA and database content; extraction is byte-identical before Excel first opens the copy.
- The owner requested keeping New Project / Open Project and opening Excel only on request. Startup leaves the project unset. New Project extracts a unique editable copy into the selected folder; Open Project uses the existing workbook without replacing it. DWGs are read from resource streams without external template files.
- A package copied outside the checkout, with no Referrings folder or external DWG/XLS/DVB/EXB files, passed template inventory, three DWG read/write probes and complete Excel extraction without requiring Office for those checks. Windows CI runs these detached-package checks.
- A second detached package passed the real desktop-Excel New Project flow, all four parameter forms, equipment-tag and print-setting save/restore, six-row calculation and generation of six GA/PID/SEC DWG/PDF pairs for two equipment rows. No AutoCAD instance or source-template path was required.
- Desktop-Excel editor regression tests passed on a copy extracted from the embedded workbook: manual save, save/close recovery, closed-workbook detection, disconnected process recovery, genuine save failure/retry and preservation of another open workbook.
- Release publishing, Visual Studio Debug Rebuild with warnings treated as errors, and compilation of the .NET Framework calculation harness passed. Startup tests also exercise the build helper through Windows PowerShell 5; its API response reader now explicitly uses UTF-8 so Chinese messages cannot corrupt JSON during rebuild preparation.
- Obsolete Referrings folders in local build outputs were moved into private .tools backups, after checking their paths. The original Referrings source directory was left untouched. The published package contains no loose business templates.
- This update adopts and versions the owner's already modified local ZTools.xls as the requested built-in workbook. Its bytes were not edited during this work. All 16 reference files match migration/source-manifest-20261023-embedded.json; the earlier manifests remain historical records.

Offline Android/iPhone follow-up — 2026-10-23:

- Added separate native C# Android and iOS hosts, a shared Chinese phone interface with the supplied CFL logo, touch pan/pinch zoom, original Excel-backed parameter groups, project settings, equipment add/remove, printing and workbook/DWG/PDF import/export. The Windows application remains available.
- The shared portable engine links the migrated calculation/drawing rules and reads/writes the original BIFF8 XLS using Apache-licensed NPOI 2.7.5, without Excel COM. It embeds all four business templates and a licensed Noto Sans SC weight-400 font; vector PDF output uses PDFsharp Core rather than GDI.
- Offline smoke tests passed for adding/removing equipment, all four parameter groups and their original choice lists, saving/restoring equipment input, print options, six-row calculation and six GA/PID/SEC DWG/PDF pairs for two selected rows. Each saved DWG was reopened and its equipment/material attributes checked by the shared drawing writer.
- A separate desktop-Excel comparison passed for 546 calculated cells across the six supplied rows, including dimensions, weights, output names and the corrected SEC/material-list lookup. Source workbooks were never edited. All 91 original non-workbook VBA/form/OLE streams remained byte-identical after mobile edits and generation; all 16 original reference-file hashes remain unchanged.
- A 390-pixel browser frame displayed the real mobile UI, created a project, loaded the motor/baseplate dialog and completed an offline calculation. The parameter screenshot is docs/images/mobile-parameters.png. This is a browser layout check of the shared UI, not evidence of physical phone testing.
- A generated SEC vector PDF was rendered and visually checked for Chinese/English material descriptions, values, section geometry, title block and table alignment with the bundled font.
- The Android ARM64 Release APK was built and its signature verified. Its manifest requests no network permissions. It is development-signed with the local Android build identity; no private keystore is included in Git. Android installation/airplane-mode testing on a physical phone remains pending because no test device was connected.
- iPhone C# API compilation passed with the installed iOS 26.0 references. Native linking/signing was skipped on Windows. The Mac build kit was extracted separately and compiled to check source/resource completeness. The owner chose to run the supplied build/signing steps on their own Mac. No signed IPA is claimed or supplied by this Windows build; tools/build-mobile-ios.sh creates it using the owner's certificate/profile.
- Local deliverables are under dist/CFL-Mobile: the Android APK, iPhone Mac build kit, installation instructions and package hashes. Signing credentials, build caches and generated results remain excluded from Git. Windows CI also runs the shared offline workbook/drawing tests without requiring Office.
