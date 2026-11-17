# M16X300 GA anchor bolt detail

`Resources/Drawings/GA.dwg` is the application template derived from `Referrings/Ztools/TEMPLATE/GA.dwg`. The new `M16X300` block copies the existing `M16X400` detail, including its insertion point, layers and linework. The original M16 variants use identical schematic geometry across lengths. The new detail follows that convention and uses a separate dimension graphics block so both the DWG dimension and the PDF display show 300.

The GA generator uses the existing anchor bolt lookup and inserts the new block when its calculated anchor bolt callout ends in `M16X300`. Recreate the derived template from the repository root with `dotnet run --project tools/DrawingTemplates`. Windows and the offline mobile engine embed the derived DWG as `GA.dwg`.
