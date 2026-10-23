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
