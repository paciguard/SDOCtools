using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace ZTools.Engine
{
    internal sealed class TemplateResourceInfo
    {
        public string Name { get; }
        public long Bytes { get; }
        public string Sha256 { get; }
        public TemplateResourceInfo(string name, long bytes, string sha256) { Name = name; Bytes = bytes; Sha256 = sha256; }
    }

    internal static class TemplateResources
    {
        private const string Prefix = "ZTools.Templates.";
        private static readonly string[] Names = { "GA.dwg", "PID.dwg", "SEC.dwg", "ZTools.xls", "Input.xls" };
        private static readonly Lazy<IReadOnlyList<TemplateResourceInfo>> Catalog = new Lazy<IReadOnlyList<TemplateResourceInfo>>(() =>
            Array.AsReadOnly(Names.Select(name =>
            {
                using (var stream = Open(name))
                using (var hash = SHA256.Create())
                    return new TemplateResourceInfo(name, stream.Length, BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant());
            }).ToArray()));

        public static IReadOnlyList<TemplateResourceInfo> Inventory => Catalog.Value;
        public static bool Contains(string name) => Names.Contains(name) && typeof(TemplateResources).Assembly.GetManifestResourceInfo(Prefix + name) != null;
        public static Stream Open(string name)
        {
            if (!Names.Contains(name)) throw new ArgumentException("未知内置模板。", nameof(name));
            return typeof(TemplateResources).Assembly.GetManifestResourceStream(Prefix + name)
                ?? throw new FileNotFoundException("缺少内置模板，请重新生成完整应用程序：" + name);
        }
    }
}
