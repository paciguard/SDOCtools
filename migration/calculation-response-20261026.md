# Calculation speed and completion — 2026-10-26

The supplied project `Test2026102601/ZTools.xls` selects `7版` for `ZE80-2200 / 160L` at input row 21. Its ZE7 worksheet lacks that model. The screenshot shows HTTP 400 after 8.7 seconds, rather than a successful calculation. ZE contains the model but must not substitute for ZE7. The project and original database were not modified to bypass this error.

## Faster calculation

Windows previously started Excel for calculation, waited for COM cleanup and reopened Excel to refresh the UI. A valid six-row calculation took about twenty seconds locally. Windows now uses `ZTools.Calculation`, a small NPOI-based library linking exactly the offline phone workbook adapter and migrated business rules. It does not include the phone PDF renderer or replace any business database. The calculation starts no Excel process. Windows Excel editing, parameter forms and DWG/PDF generation remain available.

Both Windows and mobile avoid evaluating each formula twice. Pump database ranges are read once per calculation, shared between validation and per-pump lookups, then discarded before the next calculation so later database edits are respected. Missing/ambiguous records remain errors. Sheet1 SEC results, exact List-all matching and ZE7 main/auxiliary drawing rules remain intact.

Measured locally: six-row managed Windows calculation including child-process startup, private-copy save and UI-state return took 1.7 seconds; the offline phone engine took about 0.7 seconds for six rows and 0.5 seconds for a warm 30-row run. These are development-PC timings, not physical phone benchmarks. Independent Excel COM calculation matched 546 checked cells on each managed engine, including mixed ZE/ZE7/ZA dimensions, weights and SEC names. Saved custom print settings survived calculation.

## Safe completion and UI

Windows uses a bounded hidden child on a private workbook copy. It returns rows/options directly after saving and closing the copy. The parent checks the original SHA-256 and atomically publishes only a successful result, rejecting concurrent edits. Cancellation, 60 seconds without progress, or three minutes total stop only the private calculation process. No process-wide GC/finalizer wait is used for workbook disposal.

Android/iPhone calculate in a background task on a private workbook copy. Progress/cancel commands bypass the native command gate. Cooperative cancellation is checked during workbook operations; automatic cancellation after 60 idle seconds or three total minutes is observed at these checkpoints. Failures and cancellations preserve the original XLS bytes and keep generation disabled.

On success, both interfaces clear busy state and enable **生成 DWG＋PDF**. Failure messages stay visible, including beside the Windows calculation button; the main Windows message remains visible when scrolled. A missing ZE7 record therefore returns a specific Chinese error instead of appearing to calculate indefinitely. It does not enable generation from stale output.

## Checks and delivery

Worker smoke checks cover idle/total stalls, cancellation, errors, missing response, concurrent edits and successful publication. Offline smoke checks cover cancellation, byte-identical rollback, missing ZE7 rejection, six/30-row completion and DWG/PDF output. Independent Excel parity, original-reference hashes and VBA/OLE-stream preservation are checked separately. UI checks exercise success and failure, including Generate-button state.

Windows is published under `dist/CFL-SubmitTools`. Mobile version 1.2/build 3 is rebuilt under `dist/CFL-Mobile/Android`; the iPhone ZIP is a refreshed Mac build/signing kit, not an installable IPA. Native iPhone linking/signing and physical-device checks require the owner's Mac/phones. No signed IPA or physical phone benchmark is claimed.

Reproduce the Windows request/rollback/progress check on private copies with:

```powershell
pwsh -NoProfile -File tests/calculation-response.ps1 -Workbook '<project workbook path>'
```

Without `-Workbook`, it copies the reference workbook and selects `7版` at row 21 to exercise the same missing-record case, then calculates a separate valid project. All test artifacts stay in ignored output folders.
