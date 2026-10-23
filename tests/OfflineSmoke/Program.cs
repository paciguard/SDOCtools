using ZTools.Offline;
using System.Text.Json;
using ZTools.Standalone;
var folder=args.Length>0?args[0]:"tests/offline-output";
var app=new OfflineProject(folder);
app.Create();
app.AddEquipment();app.RemoveEquipment(6);
foreach(var group in new[]{"project","equipment","motor","pid","materials"})
{
    var form=app.Parameters(group,0);Console.WriteLine($"PASS {group}: {form.Fields.Count} fields, {form.Fields.Count(f=>f.Choices.Length>0)} choice lists");
}
var equipment=app.Parameters("equipment",0);var tag=equipment.Fields.Single(f=>f.Address=="D21");
app.SaveParameters("equipment",0,new[]{new InputChange(tag.Sheet,tag.Address,"00123")});
if(app.Parameters("equipment",0).Fields.Single(f=>f.Address=="D21").Value!="00123")throw new Exception("Workbook identifier did not retain its leading zeros");
app.SaveParameters("equipment",0,new[]{new InputChange(tag.Sheet,tag.Address,tag.Value)});
app.SaveOptions(new SheetOptions(Margin:7,Copies:1));app.Calculate();
var state=JsonSerializer.SerializeToElement(app.State(),DrawingOutput.JsonOptions);
if(state.GetProperty("rows").GetArrayLength()!=6||state.GetProperty("rows")[0].GetProperty("list").GetString()!="ZE8_QWL")throw new Exception("Calculated SEC lookup differs");
Console.WriteLine("PASS six rows calculate without desktop Excel; SEC matches List-all row 2");
app.Generate(new[]{"GA","PID","SEC"},new[]{0,5});
if(app.Files().Length!=6)throw new Exception("Expected six drawing pairs");
var sceneReply=MobileBridge.Reply(19,JsonSerializer.Deserialize<JsonElement>(app.Scene(0)));
if(MobileBridge.Scripts(19,sceneReply).Count()<2)throw new Exception("Large mobile previews must use chunked messages");
foreach(var file in app.Files())foreach(var format in new[]{"dwg","pdf","scene.json"})if(new FileInfo(app.File(file.Id,format)).Length==0)throw new Exception("Empty output");
System.IO.File.WriteAllText(Path.Combine(folder,"last-workbook.txt"),app.Workbook);
Console.WriteLine("OFFLINE_SMOKE_OK "+app.Workbook);
