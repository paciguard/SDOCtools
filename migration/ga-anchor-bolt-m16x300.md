# GA anchor bolt details

`Resources/Drawings/GA.dwg` is the application template derived from `Referrings/Ztools/TEMPLATE/GA.dwg`. The new `M16X300` block copies the existing `M16X400` detail, including its insertion point, layers and linework. The original M16 variants use identical schematic geometry across lengths. The new detail follows that convention and uses a separate dimension graphics block so both the DWG dimension and the PDF display show 300.

`M12X220` uses the same approach with `M12X250` as its reference. The editable length dimension and its independent graphics cache both display 220.

The GA generator uses the existing anchor bolt lookup to select the matching block. The CZ32-250 database returns `4-M16x300` for baseplate `180M` and `4-M12x220` for baseplate `132S`. Recreate the derived template from the repository root with `dotnet run --project tools/DrawingTemplates`. Windows and the offline mobile engine embed the derived DWG as `GA.dwg`.
