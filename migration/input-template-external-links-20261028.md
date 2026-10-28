# Excel input template external-link cleanup

The cleaned input template is `outputs/2026-10-28-template-cleanup/CFL资料工具输入模板_无外链.xlsx`.
Use this copy for future manual input-template work. The original reference at
`Referrings/CFL资料工具输入模板.xlsx` is retained byte for byte, with SHA-256
`75afa463f30bd908d9fe679e4c1960dfce93471efcdde8252370e768929efaa5`.
The application database `Referrings/Ztools/ZTools.xls` was not edited.

Removed three external workbook sources: an old equipment-list template,
`DS Notes.xls`, and the missing `ZTools.xls` shown in Excel's Edit Links dialog.
Deleted nine obsolete externally referring defined names and disconnected three
hidden combo boxes from the external `Sheet1!C12:C14` cells, in both OOXML and
legacy VML representations. Their local dropdown source lists and controls remain.
Removed all six external-link package parts, their three workbook relationships,
three content-type declarations, and the external-reference table.

`tools/clean_input_template_links.py` reproduces the cleanup without opening
Excel, enabling macros, resolving network sources, or modifying the original.
It rejects an unexpected template structure rather than guessing. ZIP-part edits
preserve legacy ActiveX/VML and other native objects that a general spreadsheet
round trip may lose. No worksheet cell formula needed conversion to values.

Validation: ZIP integrity and XML parsing passed; no external workbook reference,
external relationship, or connection remains. Only five retained package parts
changed, solely for the link metadata and bindings. All worksheet cell data,
the `L15=TODAY()` formula, styles, validations, images, and unrelated object parts
are preserved. Artifact Tool imported both files and produced byte-identical
previews of `input!A1:L16`. The workbook has not been checked in native Excel.
Unrelated pre-existing legacy broken/self-referencing names remain outside this
cleanup's scope. No application source or installation package changes are needed.
