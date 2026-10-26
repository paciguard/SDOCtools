# Calculation response investigation — 2026-10-26

The supplied project selects `7版` for its first pump, `ZE80-2200 / 160L` on input row 21. Its ZE7 worksheet contains no ZE80-2200 record. The ZE worksheet does contain that model, but substituting it would violate the seventh-edition database requirement. The saved Sheet1 output still contains earlier ZE8 names; those are stale outputs and must not be treated as a successful calculation of the current inputs.

The published Windows application reproduced a generic Excel VLOOKUP exception at the first weight lookup. The calculation request returned in about eight seconds and subsequent state requests succeeded. An indefinite hang was not reproduced. Opening the workbook twice while selecting it took about eleven seconds. The old UI showed an elapsed timer and disabled controls while awaiting the whole operation, without showing the Excel calculation stage.

The generated rules now check for the exact pump/baseplate record before the first VLOOKUP. Missing and duplicate matches report the database, model, baseplate and input row. The generator contains the same change so regeneration preserves it. ZE7 never falls back to ZE. A failed project calculation keeps drawing generation blocked, restores calculated ranges, logs its failure and returns an error to the workbench.

The Windows UI polls a read-only progress endpoint backed by the existing concurrent log. This endpoint does not enter the Excel STA queue, so progress remains available while Excel is opening or calculating. It reports opening, per-pump lookups, validation/save and result loading. This does not add a timeout or cancellation for a genuinely blocked Excel COM call.

Validation on a private copy of the supplied project returned HTTP 400 in 6.4 seconds with `ZE7`, `ZE80-2200`, `160L` and input row 21 in the Chinese error. Thirty concurrent progress requests remained responsive. Saved input/output formulas matched before and after the failed calculation. Excel changes an OLE directory timestamp when opening an XLS, so unchanged project data is checked separately from file-container metadata. The original supplied workbook's SHA-256 remained unchanged. The same server then calculated the six valid built-in input rows successfully in about twenty seconds.

The offline regression also exercised mixed ZE/ZE7/ZA editions, dimensions, weights, SEC material names, GA main/auxiliary block selections, DWG/PDF generation and rejection of missing ZE7 models. All sixteen original reference files retain their recorded hashes. The Windows package was rebuilt under `dist/CFL-SubmitTools`; existing mobile installation packages were not rebuilt for this Windows diagnostic update.

Reproduce the Windows request/rollback/progress check with:

```powershell
pwsh -NoProfile -File tests/calculation-response.ps1 -Workbook '<project workbook path>'
```

Without `-Workbook`, the test uses a private copy of the original template and selects `7版` on its first row to exercise the same missing-record case. It starts its own server, keeps all test outputs under the ignored `tests/calculation-output` directory, and verifies successful calculation of a separate valid project after the failure.
