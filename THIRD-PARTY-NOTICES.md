# Third-party components

- ACadSharp 3.8.0: https://github.com/DomCR/ACadSharp — MIT license. Direct DWG reading, editing and writing.
- CSMath geometry types are included in ACadSharp; no separate CSMath package is used.
- PDFsharp / PDFsharp-GDI 6.2.4: https://github.com/empira/PDFsharp — MIT license. Vector PDF output.
- Three.js 0.180.0: https://github.com/mrdoob/three.js — MIT license. The offline modules and original license are under `ZTools.Standalone/wwwroot/vendor`.
- Microsoft .NET / ASP.NET Core 8: https://github.com/dotnet — runtime components are supplied by the installed Microsoft runtime.
- Packaged Microsoft extensions and cryptography libraries use the Microsoft MIT license included under `licenses/Microsoft-NET-LICENSE.txt`.

Original Excel, DWG, DVB, EXB, and image references are supplied by the project owner. No ownership transfer or license for those files is implied by this notice. The supplied CFL logo is retained as a project asset.

The Windows application uses a suitable TrueType font already installed on Windows; PDF embedding follows that font's installed permissions.

The offline mobile engine also uses NPOI 2.7.5 (Apache-2.0, https://github.com/nissl-lab/npoi) to retain/read/write the original XLS workbook without Excel. It intentionally uses the Apache-licensed 2.7.5 package, without accepting the separate binary-release EULA in later packages. Its transitive libraries include Bouncy Castle, SharpZipLib, Math.NET, Enums.NET, ZString, NSax, ExtendedNumerics.BigDecimal, Microsoft.IO.RecyclableMemoryStream and SixLabors image/font components; their upstream notices are included with the mobile license bundle. Android/iOS platform runtimes use Microsoft's .NET licenses.

Mobile PDFs embed a regular-weight static instance of Noto Sans SC from the Google Fonts source at https://github.com/google/fonts/tree/main/ofl/notosanssc. The original copyright and SIL Open Font License 1.1 are retained in `ZTools.Offline/Fonts/OFL.txt`. The source variable TrueType font was instantiated at weight 400 with fontTools; all other variations were removed. No Microsoft Windows font is redistributed in the mobile application.

License texts for packaged libraries are included under `licenses` in the published application.
