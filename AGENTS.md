# Project instructions

- Communicate with the owner in English. Keep application screens, dialogs, buttons, and messages in Chinese and use the supplied CFL logo.
- The active application is the standalone C# / Three.js / WebGL application. Use DWG templates directly. Do not introduce a DXF conversion requirement or an AutoCAD runtime dependency.
- Preserve `Referrings/Ztools/ZTools.xls` and every original reference file byte for byte. Excel remains the input template and database. Edit only a project workbook copy, with VBA macros disabled.
- Preserve the SEC calculation fix: use worksheet `Sheet1` by name, return the actual SEC block name separately from the material-list name, and match the material-list name exactly against `List-all` row 2. Report missing or ambiguous data instead of guessing.
- The owner explicitly requested on 2026-10-23 that this local project and `https://github.com/paciguard/SDOCtools` stay synchronized. After completing and checking future changes, commit and push the project to that repository. Preserve remote history; never force-push or overwrite unrelated remote changes. If Git transport is unavailable, the connected GitHub tools may publish the same commit; verify the resulting tree and branch head against local Git.
- Include source code, required original references, migration records, and third-party notices. Exclude credentials, private tool caches, build folders, temporary workbooks, generated drawings, and transient test output from Git. Do not change repository visibility.
