using ACadSharp.Entities;
using ZTools.Standalone;

if(args.Length!=2)throw new ArgumentException("Pass generated drawing directory and review directory.");
Directory.CreateDirectory(args[1]);
int count=0;
foreach(var path in Directory.GetFiles(args[0],"*.dwg",SearchOption.AllDirectories))
{
    var document=DrawingOutput.Read(path,new());
    var tag=document.Entities.OfType<Insert>().SelectMany(i=>i.Attributes).First(a=>a.Tag=="SHEBEIWEIHAO").Value;
    var scene=DrawingOutput.Scene(document,new());
    if(!scene.Items.Any(i=>i.Type=="text"&&i.Text==tag))throw new Exception("Missing equipment tag: "+path);
    string output=Path.Combine(args[1],Path.GetFileNameWithoutExtension(path));
    DrawingOutput.Pdf(output+".pdf",scene);
    File.WriteAllText(output+".scene.json",System.Text.Json.JsonSerializer.Serialize(scene,DrawingOutput.JsonOptions));
    Console.WriteLine("PASS vector PDF and scene from saved DWG: "+Path.GetFileName(path));count++;
}
if(count!=6)throw new Exception("Expected six previously generated DWG files.");
Console.WriteLine("DRAWING_REVIEW_OK");
