# Drawing creation completion prompt — 2026-10-26

Windows, Android and iPhone now show one Chinese **图纸生成完成** dialog after the whole requested DWG/PDF batch has been saved. Windows includes the output folder; mobile includes the DWG/PDF count and preview/export guidance. **确定** dismisses the dialog. It remains visible until acknowledged, replacing the earlier short mobile toast and adding a visible Windows conclusion.

The UI waits for the successful generation response and releases its busy controls before opening the dialog. Failed generation never sets its completion result and cannot show the success prompt. A subsequent Windows preview error does not hide a successful file creation. No calculation, workbook, template or drawing rules changed. Mobile packages advance to version 1.3/build 4; the iPhone deliverable remains a Mac build/signing kit, not a signed IPA.
