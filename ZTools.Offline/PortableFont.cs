namespace ZTools.Standalone;
internal sealed class CflFontResolver:PdfSharp.Fonts.IFontResolver
{
    public byte[] GetFont(string faceName)
    {
        using var stream=typeof(CflFontResolver).Assembly.GetManifestResourceStream("ZTools.Fonts.NotoSansSC.ttf")??throw new InvalidOperationException("缺少内置中文字体。");
        using var memory=new MemoryStream();stream.CopyTo(memory);return memory.ToArray();
    }
    public PdfSharp.Fonts.FontResolverInfo ResolveTypeface(string familyName,bool bold,bool italic)=>new("cfl-cjk");
}
