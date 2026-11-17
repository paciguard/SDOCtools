using ACadSharp.Entities;
using ACadSharp.IO;

// Run from the repository root. Only the derived GA template is written.
const string sourcePath = "Referrings/Ztools/TEMPLATE/GA.dwg";
const string destinationPath = "Resources/Drawings/GA.dwg";
var boltVariants = new[]
{
    (Source: "M16X400", Target: "M16X300", OldLength: "400", NewLength: "300"),
    (Source: "M12X250", Target: "M12X220", OldLength: "250", NewLength: "220")
};

var document = DwgReader.Read(sourcePath);
var originalBlocks = document.BlockRecords.ToDictionary(
    block => block.Name, block => block.Entities.Select(entity => entity.ObjectName).ToArray());
int cacheIndex = 1;
foreach (var (sourceName, targetName, oldLength, newLength) in boltVariants)
{
    var bolt = (ACadSharp.Tables.BlockRecord)document.BlockRecords[sourceName].Clone();
    bolt.Name = targetName;

    // Preserve the schematic geometry and update the length in the editable
    // dimension and its cached graphics.
    var length = bolt.Entities.OfType<Dimension>().Single(dimension => dimension.Text == oldLength);
    length.Block.Entities.OfType<MText>().Single(text => text.PlainText == oldLength).Value = newLength;
    length.Text = newLength;

    // Assign unique names before attaching the new block so the DWG library
    // cannot resolve the clones to original dimension caches.
    foreach (var dimension in bolt.Entities.OfType<Dimension>())
    {
        while (document.BlockRecords.Contains($"*D{cacheIndex}"))
            cacheIndex++;
        dimension.Block.Name = $"*D{cacheIndex++}";
        document.BlockRecords.Add(dimension.Block);
    }
    document.BlockRecords.Add(bolt);
}

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

    if (saved.Entities.Count() != document.Entities.Count())
        throw new InvalidOperationException("The derived GA template changed the model space entity count.");

    foreach (var (sourceName, targetName, oldLength, newLength) in boltVariants)
    {
        var savedOriginal = saved.BlockRecords[sourceName];
        var savedBolt = saved.BlockRecords[targetName];
        var originalLength = savedOriginal.Entities.OfType<Dimension>().Single(dimension => dimension.Text == oldLength);
        var savedLength = savedBolt.Entities.OfType<Dimension>().Single(dimension => dimension.Text == newLength);
        if (savedBolt.BlockEntity.BasePoint != savedOriginal.BlockEntity.BasePoint
            || !savedBolt.Entities.Select(entity => entity.ObjectName).SequenceEqual(originalBlocks[sourceName])
            || savedLength.Block.Name == originalLength.Block.Name
            || savedLength.Block.Entities.OfType<MText>().Single().PlainText != newLength
            || originalLength.Block.Entities.OfType<MText>().Single().PlainText != oldLength)
            throw new InvalidOperationException($"The {targetName} block failed DWG round-trip verification.");
    }

    File.Move(temporaryPath, destinationPath, overwrite: true);
    Console.WriteLine($"Added {string.Join(", ", boltVariants.Select(variant => variant.Target))} to {destinationPath}.");
}
finally
{
    File.Delete(temporaryPath);
}
