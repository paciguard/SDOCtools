# Linked SEC part materials

The owner requested that parts 360.01 and 360.02 follow a specified bearing-house material. `Sec-List!O10:O39` now supplies the same override to 330.01, 330.02, 330.03, 360.01 and 360.02. Blank input retains each part's material from the selected grade in `List-detail`. No reference workbook or database values are changed.

The shared calculation/drawing engine applies this rule on Windows and mobile. Parameter labels identify the affected parts. Regression coverage checks both 360 parts in the prepared material table, their matching PN/MAT attributes in the saved SEC DWG, their material rows in the PDF scene, and database fallback after clearing the override.
