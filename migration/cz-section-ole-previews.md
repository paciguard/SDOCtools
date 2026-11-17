# CZ shaft seal details in SEC output

The original SEC DWG contains two OLE pictures in each of `CZWBALHS`, `CZWBALHD`, `CZWBALHL`, `CZWOBALHS` and `CZWOBALHD`. They show the stuffing box gland and double mechanical seal. The PDF renderer previously stopped at the first visible picture.

`Resources/Drawings/SectionDetails` contains two monochrome PNGs derived from the original pictures. All ten OLE payloads have a compound file at offset 128. Their `CONTENTS` streams contain the same two enhanced metafiles, with the 4-byte stream prefix excluded. The metafiles were rendered at 1,000 pixels wide onto white and saved as grayscale after checking that the source had equal red, green and blue channels. `SectionDetailImages.cs` maps the exact ten source payload hashes to the corresponding image.

The model keeps the original OLE pictures in the saved DWG. PDF and preview use the matching PNGs at the frame's transformed corners, under the vector linework and labels. An unrecognized visible OLE payload still reports an unsupported entity. Each known source block is checked and both images are verified in saved CZ DWG/PDF/preview output.
