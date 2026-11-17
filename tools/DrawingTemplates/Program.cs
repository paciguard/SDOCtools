using ACadSharp.Entities;
using ACadSharp.IO;

// Run from the repository root. Only the derived GA template is written.
const string sourcePath = "Referrings/Ztools/TEMPLATE/GA.dwg";
const string destinationPath = "Resources/Drawings/GA.dwg";
const string sourceName = "M16X400";
const string destinationName = "M16X300";

var document = DwgReader.Read(sourcePath);
var originalBlocks = document.BlockRecords.ToDictionary(
    block => block.Name, block => block.Entities.Select(entity => entity.ObjectName).ToArray());
var source = document.BlockRecords[sourceName];
var bolt = (ACadSharp.Tables.BlockRecord)source.Clone();
bolt.Name = destinationName;

// The reference details for M16x250 and M16x400 are schematic. Preserve
// their shared geometry and change the displayed length in both representations.
var length = bolt.Entities.OfType<Dimension>().Single(dimension => dimension.Text == "400");
length.Block.Entities.OfType<MText>().Single(text => text.PlainText == "400").Value = "300";
length.Text = "300";

// Assign unique names to all copied caches before attaching the new block.
// Otherwise the DWG library can resolve the clone to an original cache.
int cacheIndex = 1;
foreach (var dimension in bolt.Entities.OfType<Dimension>())
{
    while (document.BlockRecords.Contains($"*D{cacheIndex}"))
        cacheIndex++;
    dimension.Block.Name = $"*D{cacheIndex++}";
    document.BlockRecords.Add(dimension.Block);
}
document.BlockRecords.Add(bolt);

Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
string temporaryPath = Path.ChangeExtension(destinationPath, $"{Guid.NewGuid():N}.dwg");
try
{
    DwgWriter.Write(temporaryPath, document);
    var saved = DwgReader.Read(temporaryPath);
    foreach (var (name, entityTypes) in originalBlocks)
        if (!saved.BlockRecords.TryGetValue(name, out var block)
            || !entityTypes.SequenceEqual(block.Entities.Select(entity => entity.ObjectName)))
            throw new InvalidOperationException($"Original GA block changed in the derived DWG: {name}.");

    var savedOriginal = saved.BlockRecords[sourceName];
    var savedBolt = saved.BlockRecords[destinationName];
    var originalLength = savedOriginal.Entities.OfType<Dimension>().Single(dimension => dimension.Text == "400");
    var savedLength = savedBolt.Entities.OfType<Dimension>().Single(dimension => dimension.Text == "300");
    if (saved.Entities.Count() != document.Entities.Count()
        || savedBolt.BlockEntity.BasePoint != savedOriginal.BlockEntity.BasePoint
        || !savedBolt.Entities.Select(entity => entity.ObjectName).SequenceEqual(savedOriginal.Entities.Select(entity => entity.ObjectName))
        || savedLength.Block.Name == originalLength.Block.Name
        || savedLength.Block.Entities.OfType<MText>().Single().PlainText != "300"
        || originalLength.Block.Entities.OfType<MText>().Single().PlainText != "400")
        throw new InvalidOperationException("The derived GA template failed DWG round-trip verification.");

    File.Move(temporaryPath, destinationPath, overwrite: true);
    Console.WriteLine($"Added {destinationName} to {destinationPath}.");
}
finally
{
    File.Delete(temporaryPath);
}
