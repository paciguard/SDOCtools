using ZTools.Offline;
using System.Text.Json;
using ZTools.Standalone;
using ZTools.Engine;
using ACadSharp.IO;
using ACadSharp.Entities;
var folder=args.Length>0?args[0]:"tests/offline-output";
var app=new OfflineProject(folder);
app.Create();
app.AddEquipment();app.RemoveEquipment(6);
foreach(var group in new[]{"project","equipment","motor","pid","materials","utility"})
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

var housingStorage=Path.Combine(folder,"bearing-house-regression");
var housing=new OfflineProject(housingStorage);housing.Create();
housing.SaveParameters("materials",0,new[]{new InputChange("Sec-List","O10","BH-330-CHECK")});
var savedId=Path.GetFileName(Path.GetDirectoryName(housing.Workbook));
housing=new OfflineProject(housingStorage);housing.Open(savedId);
if(housing.Parameters("materials",0).Fields.Single(f=>f.Address=="O10").Value!="BH-330-CHECK")throw new Exception("Bearing-house input did not survive project reopen");
housing.Calculate();housing.Generate(new[]{"SEC"},new[]{0});
var housingDrawing=DwgReader.Read(housing.File(0,"dwg"));
if(!housingDrawing.Entities.OfType<Insert>().SelectMany(i=>i.Attributes).Any(a=>a.Value=="BH-330-CHECK"))throw new Exception("Saved SEC DWG did not contain the bearing-house material override");
if(!housing.Scene(0).Contains("BH-330-CHECK"))throw new Exception("SEC PDF scene did not contain the bearing-house material override");
using(var session=new WorkbookSession(housing.Workbook))
{
    new SubmissionEngine(session,_=>{}).PrepareSection(0);int houses=0;
    for(int r=72;r<=123;r++)if(new[]{"330.01","330.02","330.03"}.Contains(session.Cell("Sheet1","DB"+r)))
    {houses++;if(session.Cell("Sheet1","DD"+r)!="BH-330-CHECK")throw new Exception("SEC material list did not apply the bearing-house material");}
    if(houses==0)throw new Exception("Bearing-house regression list contains no bearing-house parts");
}
using(var source=TemplateResources.Open("ZTools.xls"))using(var internalModel=new NPOI.HSSF.UserModel.HSSFWorkbook(source,true))
    if(internalModel.GetSheet("map")!=null||internalModel.GetSheet("List-detail")==null)throw new Exception("Embedded model must remove map and retain the material database");
using(var saved=System.IO.File.OpenRead(housing.Workbook))using(var publicProject=new NPOI.HSSF.UserModel.HSSFWorkbook(saved,true))
{
    if(publicProject.NumberOfSheets!=1||publicProject.GetSheetName(0)!="input")throw new Exception("Public project exposes internal data sheets");
    if(publicProject.GetSheet("input").GetDataValidations().Count<8)throw new Exception("Input-only workbook lost its dropdowns");
}
string originalGrade=housing.Parameters("equipment",0).Fields.Single(f=>f.Address=="O21").Value;
string sectionGrade=housing.Parameters("materials",0).Fields.Single(f=>f.Address=="F10").Choices.First(g=>g!=originalGrade);
housing.SaveParameters("materials",0,new[]{new InputChange("Sec-List","F10",sectionGrade)});
housing.Calculate();
using(var session=new WorkbookSession(housing.Workbook))
{
    if(session.Cell("input","O21")!=originalGrade||session.Cell("Sec-List","F10")!=sectionGrade||session.Cell("Sheet1","CE40")!=sectionGrade)
        throw new Exception("Explicit Sec-List grade was overwritten by the input grade");
    new SubmissionEngine(session,_=>{}).PrepareSection(0);
    int gradeColumn=Enumerable.Range(6,11).Single(c=>session.Cell("List-detail",NPOI.SS.Util.CellReference.ConvertNumToColString(c-1)+"2")==sectionGrade);
    int shaftRow=Enumerable.Range(3,151).First(r=>session.Cell("List-detail","C"+r)=="210.01");
    string expected=session.Cell("List-detail",NPOI.SS.Util.CellReference.ConvertNumToColString(gradeColumn-1)+shaftRow);
    int outputRow=Enumerable.Range(72,52).Single(r=>session.Cell("Sheet1","DB"+r)=="210.01");
    if(session.Cell("Sheet1","DD"+outputRow)!=expected)throw new Exception("SEC material list ignored the explicit Sec-List grade");
}
Console.WriteLine("PASS explicit Sec-List grade survives calculation and selects shaft material from the embedded database");
housing.SaveParameters("materials",0,new[]{new InputChange("Sec-List","N10","UNSUPPORTED-BEARING")});
bool rejectedHousingType=false;try{housing.Calculate();}catch(InvalidOperationException){rejectedHousingType=true;}
if(!rejectedHousingType)throw new Exception("Sec-List bearing type was ignored by SEC selection");
Console.WriteLine("INPUT_ONLY_AND_BEARING_HOUSE_OK persisted overrides, map-free embedded data, input dropdowns, SEC list/DWG/PDF scene, and bearing-type validation");

// Mixed editions exercise the same shared calculation rules used on Windows.
var seventh=new OfflineProject(Path.Combine(folder,"ze7-regression"));seventh.Create();
foreach(var (row,standard) in new[]{(1,"7版"),(3," 7版 "),(4,"7th")})
    seventh.SaveParameters("equipment",row,new[]{new InputChange("input","E"+(21+row),standard)});
seventh.Calculate();
using(var file=System.IO.File.OpenRead(seventh.Workbook))
using(var book=new ZTools.Engine.ExcelBook(file))
{
    for(int k=0;k<6;k++)
    {
        string family=k==5?"ZA":(k==1||k==3||k==4?"ZE7":"ZE");
        string model=Convert.ToString(book.Worksheets["input"].Cells[21+k,14].Value2);
        string plate=Convert.ToString(book.Worksheets["input"].Cells[21+k,21].Value2);
        var data=(object[,])book.Worksheets[family].Range["A32:AE2000"].Value2;
        int match=Enumerable.Range(1,data.GetLength(0)).First(r=>Convert.ToString(data[r,1])==model&&Convert.ToString(data[r,6])==plate);
        for(int c=0;c<13;c++)
            if(Convert.ToString(book.Worksheets["Sheet1"].Cells[40+k,24+c].Value2)!=Convert.ToString(data[match,7+c]))throw new Exception($"{family} GA dimension mismatch at row {40+k}, column {24+c}");
        foreach(var (target,source) in new[]{(6,24),(7,25),(8,26),(37,20)})
            if(Convert.ToString(book.Worksheets["Sheet1"].Cells[40+k,target].Value2)!=Convert.ToString(data[match,source]))throw new Exception($"{family} GA mass/f mismatch at row {40+k}, column {target}");
    }
}
seventh.Generate(new[]{"GA"},new[]{1,3,4});
if(seventh.Files().Length!=3)throw new Exception("Expected three ZE7 GA drawing pairs");
foreach(var output in seventh.Files())
{
    var drawing=DwgReader.Read(seventh.File(output.Id,"dwg"));
    var names=drawing.Entities.OfType<Insert>().Select(i=>i.Block.Name).ToArray();
    if(!names.Contains("ZE7")||names.Contains("ZE"))throw new Exception("Seventh-edition GA must contain the ZE7 main block");
    if(!names.Contains("jibenchicun8"))throw new Exception("ZE7 GA must retain the ZE8 auxiliary dimension block");
}
// Hold all calculated data/options fixed and compare the actual insertion
// paths. Changing the edition must change only the main block name.
using(var session=new WorkbookSession(seventh.Workbook))
{
    session.Book.Worksheets["Sheet1"].Range["C19"].Value2=41;
    string[] Blocks(string standard)
    {
        session.Book.Worksheets["input"].Range["E22"].Value2=standard;
        using var template=TemplateResources.Open("GA.dwg");var drawing=DwgReader.Read(template);
        var engine=new SubmissionEngine(session,_=>{});engine.Drawing=new CadAdapter(drawing);engine.BuildGA();
        return drawing.Entities.OfType<Insert>().Select(i=>i.Block.Name).ToArray();
    }
    var seven=Blocks("7版");var eight=Blocks("8版");
    if(!seven.Contains("ZE7")||!eight.Contains("ZE"))throw new Exception("Unexpected GA main blocks");
    if(!seven.Where(n=>n!="ZE7").Order().SequenceEqual(eight.Where(n=>n!="ZE").Order()))throw new Exception("ZE7 auxiliary blocks differ from the ZE8 path");
}
Console.WriteLine("ZE7_GA_BLOCKS_OK saved DWGs use ZE7; all auxiliary block names match ZE8 for the same data/options");

seventh.SaveParameters("equipment",1,new[]{new InputChange("input","N22","ZE80-2200")});
bool missingRejected=false;try{seventh.Calculate();}catch(InvalidOperationException ex){missingRejected=ex.Message.Contains("ZE80-2200");}
if(!missingRejected)throw new Exception("A missing ZE7 model must fail instead of using ZE");
Console.WriteLine("ZE7_GA_REGRESSION_OK Chinese/legacy editions, mixed ZE/ZE7/ZA rows, 13 dimensions, weights, f, DWG/PDF and missing-record rejection");
