using System.Security.Cryptography;
using ACadSharp.Entities;

namespace ZTools.Standalone;

internal static class SectionDetailImages
{
    // The five legacy CZ section blocks each carry the same two enhanced
    // metafiles in distinct OLE payloads. Match the exact source bytes so an
    // unfamiliar OLE object still produces the unsupported-entity error.
    private static readonly Dictionary<string, char> SourceImages = new(StringComparer.OrdinalIgnoreCase)
    {
        ["babb9f7cf8b361938bcfce99f25359f304b4be4812865db2b2baa773dfcf339a"] = 'A',
        ["62eb931b9121464ac1e780d0f95799018d2fc8ea99417236a373ec212a1655aa"] = 'B',
        ["9700b0c7dbfc74b3c48b8c351849df3fef8db1b8596f4bf3f14baefd36d7fe8b"] = 'A',
        ["31a279a7c6621701d2915e0f2903dc992aa97a244865ad29c783c8ee7465364b"] = 'B',
        ["7a0cd10d62ac98dacc5eb4e07354a62e29fd8d7f8760318969f96b3f706d06f4"] = 'A',
        ["d0f710cc755ed92a3e51546543af362d662851af4adc341ce8e058fcc6f9c6ae"] = 'B',
        ["420b3224d94f0e3e8cc1bc836ed4b0c40547d1ee1ce5b7b898813d3e58efcf90"] = 'A',
        ["ae112d1699f264e171dfb6cb096a265e4dffb574384d1ac37abfc5c14e80a525"] = 'B',
        ["f116c604c9d42732c94182128e62e7326f389a12c6f235cbd43414273b602a52"] = 'A',
        ["c9f9ffc7d491cb50ecec05c8c598ad03299bad740de105a4678cd468832e1e81"] = 'B'
    };

    private static readonly Lazy<string> DetailA = new(() => Load('A'));
    private static readonly Lazy<string> DetailB = new(() => Load('B'));

    internal static string GetImage(Ole2Frame frame)
    {
        if (frame.BinaryData is not { Length: > 0 } data
            || !SourceImages.TryGetValue(Convert.ToHexString(SHA256.HashData(data)), out char detail))
            return null;

        return detail == 'A' ? DetailA.Value : DetailB.Value;
    }

    private static string Load(char detail)
    {
        using var stream = typeof(SectionDetailImages).Assembly.GetManifestResourceStream($"ZTools.SectionDetails.CZ-Detail-{detail}.png")
            ?? throw new InvalidOperationException("缺少内置的 CZ 剖面局部图，请重新安装应用程序。");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Convert.ToBase64String(buffer.ToArray());
    }
}
