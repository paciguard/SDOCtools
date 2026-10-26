using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using ZTools.Offline;
using ZTools.Standalone;

var project=new OfflineProject(Path.GetFullPath("tests/offline-calculation-output/"+Guid.NewGuid().ToString("N")));
project.Create();
byte[] original=SHA256.HashData(File.ReadAllBytes(project.Workbook));
var pending=Task.Run(()=>project.Calculate());var wait=Stopwatch.StartNew();
while(!pending.IsCompleted&&wait.Elapsed.TotalSeconds<60)
{
    var progress=JsonSerializer.SerializeToElement(project.Progress(),DrawingOutput.JsonOptions);
    if(progress.GetProperty("log").EnumerateArray().Any(line=>(line.GetString()??"").Contains("正在查询")))break;
    await Task.Delay(30);
}
if(pending.IsCompleted)throw new Exception("Calculation finished before cancellation could be tested.");
var timer=Stopwatch.StartNew();project.CancelCalculation();
try{await pending;throw new Exception("Cancellation did not stop the calculation.");}
catch(InvalidOperationException error) when(error.Message.Contains("已取消计算")){}
if(timer.Elapsed.TotalSeconds>5||!SHA256.HashData(File.ReadAllBytes(project.Workbook)).SequenceEqual(original))throw new Exception("Cancellation changed the workbook or did not return promptly.");
if(!JsonSerializer.SerializeToElement(project.State(),DrawingOutput.JsonOptions).GetProperty("needsCalculation").GetBoolean())throw new Exception("Cancellation enabled drawing generation.");
Console.WriteLine($"PASS offline cancellation returns in {timer.Elapsed.TotalSeconds:N1}s; original XLS byte-identical; generation remains blocked");
project.SaveParameters("equipment",0,new[]{new InputChange("input","E21","7版")});
original=SHA256.HashData(File.ReadAllBytes(project.Workbook));
try{project.Calculate();throw new Exception("Missing ZE7 record was not rejected.");}
catch(InvalidOperationException error) when(error.Message.Contains("ZE7")&&error.Message.Contains("ZE80-2200")){}
if(!SHA256.HashData(File.ReadAllBytes(project.Workbook)).SequenceEqual(original))throw new Exception("Failed lookup changed the original XLS.");
Console.WriteLine("PASS offline missing-record failure preserves original XLS byte for byte");
project.SaveParameters("equipment",0,new[]{new InputChange("input","E21","8版")});
timer.Restart();var result=JsonSerializer.SerializeToElement(project.Calculate(),DrawingOutput.JsonOptions);
if(result.GetProperty("needsCalculation").GetBoolean()||result.GetProperty("rows").GetArrayLength()!=6)throw new Exception("Successful calculation did not enable generation.");
if(Directory.GetDirectories(Path.GetDirectoryName(project.Workbook),".cfl-calculation-*").Any())throw new Exception("Temporary calculation workbook remains.");
Console.WriteLine($"PASS offline six-row calculation returns ready-for-generation state in {timer.Elapsed.TotalSeconds:N1}s after cancellation/error");
using(var stream=File.OpenRead(project.Workbook))
using(var book=new ZTools.Engine.ExcelBook(stream))
{
    var input=book.Worksheets["input"];
    var rows=(object[,])input.Range["D21:AD26"].Value2;
    for(int row=27;row<=50;row++)
        for(int col=4;col<=30;col++)input.Cells[row,col].Value2=rows[(row-21)%6+1,col-3];
    book.Save(project.Workbook);
}
timer.Restart();result=JsonSerializer.SerializeToElement(project.Calculate(),DrawingOutput.JsonOptions);
if(result.GetProperty("needsCalculation").GetBoolean()||result.GetProperty("rows").GetArrayLength()!=30)throw new Exception("Thirty-row calculation did not return ready state.");
Console.WriteLine($"PASS offline thirty-row calculation returns ready-for-generation state in {timer.Elapsed.TotalSeconds:N3}s");
Console.WriteLine("OFFLINE_CALCULATION_OK");
