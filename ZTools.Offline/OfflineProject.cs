using System.Globalization;
using System.Collections.Concurrent;
using System.Text.Json;
using NPOI.SS.Util;
using ZTools.Engine;
using ZTools.Standalone;
namespace ZTools.Offline;

public sealed record EquipmentRow(int Index,string Tag,string Model,string Material,string Template,string List);
public sealed record InputField(string Sheet,string Address,string Label,string Value,string[] Choices,bool ReadOnly=false);
public sealed record ParameterForm(string Title,List<InputField> Fields);
public sealed record InputChange(string Sheet,string Address,string Value);
public sealed record ProjectEntry(string Id,string Name);
public sealed record OutputEntry(int Id,string Name,string Kind);
public sealed record OfflineState(string Job,bool NeedsCalculation,EquipmentRow[] Rows,SheetOptions Options,OutputEntry[] Files,string[] Log,bool Offline=true);

// Native hosts call this directly: no listener, Excel process or remote service.
public sealed partial class OfflineProject
{
    readonly string storage;readonly ConcurrentQueue<string> log=new();string workbook,output;bool needsCalculation=true;
    public OfflineProject(string storage){this.storage=Path.GetFullPath(storage);Directory.CreateDirectory(this.storage);System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);}
    public string Workbook=>workbook;
    void Report(string message){log.Enqueue(message);while(log.Count>100)log.TryDequeue(out _);}
    public object Progress()=>new{log=log.ToArray()};
    static string ProjectFile(string folder)=>System.IO.File.Exists(Path.Combine(folder,"ZTools.xlsx"))?Path.Combine(folder,"ZTools.xlsx"):Path.Combine(folder,"ZTools.xls");
    public ProjectEntry[] Projects()=>Directory.GetDirectories(storage).Where(d=>System.IO.File.Exists(ProjectFile(d))).OrderDescending().Select(d=>new ProjectEntry(Path.GetFileName(d),Path.GetFileName(d))).ToArray();
    public object Create()
    {
        string dir=Path.Combine(storage,"项目-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")[..6]);Directory.CreateDirectory(dir);
        workbook=Path.Combine(dir,"ZTools.xlsx");ProjectWorkbook.Create(workbook);
        Reset();Report("项目已从内置 Excel 模板创建。");return State();
    }
    public object Open(string id)
    {
        if(!Projects().Any(p=>p.Id==id))throw new ArgumentException("项目不存在。");
        workbook=ProjectFile(Path.Combine(storage,id));
        if(Path.GetExtension(workbook).Equals(".xls",StringComparison.OrdinalIgnoreCase))
        {
            using var input=System.IO.File.OpenRead(workbook);using var source=NPOI.SS.UserModel.WorkbookFactory.Create(input);using var model=ProjectWorkbook.Materialize(source);
            string converted=Path.ChangeExtension(workbook,".xlsx");ProjectWorkbook.WriteInput(model,converted);workbook=converted;
        }
        Reset();return State();
    }
    public object Import(Stream input)
    {
        string dir=Path.Combine(storage,"导入-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")[..6]);Directory.CreateDirectory(dir);
        string file=Path.Combine(dir,"ZTools.xlsx");using(var copy=System.IO.File.Create(file)){byte[] buffer=new byte[81920];int count;long total=0;while((count=input.Read(buffer))>0){total+=count;if(total>20*1024*1024)throw new InvalidOperationException("项目文件不能超过 20 MB。");copy.Write(buffer,0,count);}}
        using(var validation=new WorkbookSession(file)){validation.Save();}workbook=file;Reset();return State();
    }
    void Reset(){output=null;needsCalculation=true;log.Clear();}
    public object AddEquipment()
    {
        using(var b=Book())
        {
            var current=(ExcelBook)b.Book;int count=0;while(count<30&&!string.IsNullOrWhiteSpace(b.Cell("input","N"+(21+count))))count++;
            if(count==30)throw new InvalidOperationException("每个项目最多支持 30 台设备。");
            if(count>0)CopyEquipment(current,count-1,current,count);
            else{using var stream=TemplateResources.Open("ZTools.xls");using var defaults=new ExcelBook(stream);CopyEquipment(defaults,0,current,0);}
            b.Book.Worksheets["input"].Range["D"+(21+count)].Value2="设备-"+(count+1);b.Save();
        }
        needsCalculation=true;output=null;return State();
    }
    public object RemoveEquipment(int index)
    {
        if(index<0||index>=30)throw new ArgumentOutOfRangeException(nameof(index));
        using(var b=Book())
        {
            if(string.IsNullOrWhiteSpace(b.Cell("input","N"+(21+index))))throw new ArgumentException("设备不存在。");
            var current=(ExcelBook)b.Book;int last=index;
            while(last<29&&!string.IsNullOrWhiteSpace(b.Cell("input","N"+(22+last)))){CopyEquipment(current,last+1,current,last);last++;}
            for(int col=4;col<=30;col++){var target=current.Data.GetSheet("input").GetRow(20+last)?.GetCell(col-1);if(target!=null&&target.CellType!=NPOI.SS.UserModel.CellType.Formula)target.SetCellType(NPOI.SS.UserModel.CellType.Blank);}
            current.Evaluator.ClearAllCachedResultValues();b.Save();
        }
        needsCalculation=true;output=null;return State();
    }
    static void CopyEquipment(ExcelBook source,int from,ExcelBook target,int to)
    {
        foreach(var (sheet,offset,start,end) in new[]{("input",21,4,30),("PID",10,6,9),("Sec-List",10,6,15)})
        for(int col=start;col<=end;col++)
        {
            var src=source.Data.GetSheet(sheet).GetRow(offset+from-1)?.GetCell(col-1);var dest=target.Data.GetSheet(sheet).GetRow(offset+to-1)?.GetCell(col-1);
            if(src?.CellType==NPOI.SS.UserModel.CellType.Formula||dest?.CellType==NPOI.SS.UserModel.CellType.Formula)continue;
            target.Worksheets[sheet].Cells[offset+to,col].Value2=source.Value(src);
        }
    }
    WorkbookSession Book()=>new(workbook??throw new InvalidOperationException("请先新建或打开项目。"));
    public object State()
    {
        if(workbook!=null){using var book=Book();return Snapshot(book);}
        return new OfflineState(null,needsCalculation,Array.Empty<EquipmentRow>(),new(),Files(),log.ToArray());
    }
    OfflineState Snapshot(WorkbookSession b)
    {
        var rows=new List<EquipmentRow>();SheetOptions options=new();
        for(int i=0;i<30;i++)if(!string.IsNullOrWhiteSpace(b.Cell("input","N"+(21+i))))rows.Add(new(i,b.Cell("input","D"+(21+i)),b.Cell("input","N"+(21+i)),b.Cell("input","O"+(21+i)),b.Cell("Sheet1","CA"+(40+i)),b.Cell("Sheet1","CC"+(40+i))));
        options=ReadOptions(b);
        return new OfflineState(Path.GetFileName(Path.GetDirectoryName(workbook)),needsCalculation,rows.ToArray(),options,Files(),log.ToArray());
    }
    public ParameterForm Parameters(string group,int index)
    {
        if(index<0||index>=30)throw new ArgumentOutOfRangeException(nameof(index));using var b=Book();var fields=new List<InputField>();
        void Field(string sheet,string address,string label){var point=new CellReference(address);var source=((ExcelBook)b.Book).Data.GetSheet(sheet).GetRow(point.Row)?.GetCell(point.Col);bool readOnly=source?.CellType==NPOI.SS.UserModel.CellType.Formula&&!(sheet=="Sec-List"&&(point.Col==5||point.Col==13));string value=b.Cell(sheet,address);if(source!=null&&(source.CellType==NPOI.SS.UserModel.CellType.Numeric||(source.CellType==NPOI.SS.UserModel.CellType.Formula&&source.CachedFormulaResultType==NPOI.SS.UserModel.CellType.Numeric))&&NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(source))value=((ExcelBook)b.Book).Formatter.FormatCellValue(source,((ExcelBook)b.Book).Evaluator);fields.Add(new(sheet,address,string.IsNullOrWhiteSpace(label)?address:label,value,Choices((ExcelBook)b.Book,sheet=="Sec-List"&&(point.Col==5||point.Col==13)?"input":sheet,sheet=="Sec-List"&&(point.Col==5||point.Col==13)?(point.Col==5?"O":"AA")+(point.Row+12):address),readOnly));}
        string title;
        switch(group)
        {
            case "project":title="项目与计算设置";foreach(string address in new[]{"C4","C5","C6","C10"})Field("Sheet1",address,b.Cell("Sheet1","B"+address[1..]));foreach(string address in new[]{"I53","I56"})Field("input",address,address=="I53"?"法兰标准":"管口载荷系数");break;
            case "equipment":title="设备参数";foreach(string col in new[]{"D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","X","Y","Z","AA","AB","AC","AD"})Field("input",col+(21+index),b.Cell("input",col+"19"));break;
            case "motor":title="电机与底座";foreach(string col in new[]{"S","T","U","V","W"})Field("input",col+(21+index),b.Cell("input",col+"19"));break;
            case "pid":title="管仪参数";foreach(string col in new[]{"F","G","H","I"})Field("PID",col+(10+index),b.Cell("PID",col+"8"));foreach(string col in new[]{"Y","Z"})Field("input",col+(21+index),b.Cell("input",col+"19"));break;
            case "materials":title="剖面与材料清单参数";foreach(string col in new[]{"F","G","H","I","J","K","L","M","N","O"})Field("Sec-List",col+(10+index),col=="O"?"轴承架材质（330.01 / 330.02 / 330.03）":b.Cell("Sec-List",col+"8"));break;
            case "utility":title="公用工程参数";foreach(string col in new[]{"E","F","G","H","I","J","K"})Field("Utility",col+(13+index),b.Cell("Utility",col+"12"));break;
            default:throw new ArgumentException("未知参数窗口。");
        }
        return new(title,fields);
    }
    static string[] Choices(ExcelBook book,string sheet,string address)
    {
        var cell=new CellReference(address);
        IList<NPOI.SS.UserModel.IDataValidation> validations;
        try{validations=book.Data.GetSheet(sheet).GetDataValidations();}
        catch(InvalidOperationException error)when(error.Message.StartsWith("Unexpected (UnknownRecord) while looking for DV Table insert pos",StringComparison.Ordinal)){return Array.Empty<string>();}
        var validation=validations.FirstOrDefault(v=>v.Regions.CellRangeAddresses.Any(a=>a.IsInRange(cell.Row,cell.Col)));
        var constraint=validation?.ValidationConstraint;if(constraint==null)return Array.Empty<string>();
        if(constraint.ExplicitListValues?.Length>0)return constraint.ExplicitListValues;
        string formula=constraint.Formula1?.TrimStart('=');if(string.IsNullOrEmpty(formula))return Array.Empty<string>();
        var named=book.Data.GetName(formula);if(named!=null)formula=named.RefersToFormula;
        if(!System.Text.RegularExpressions.Regex.IsMatch(formula,@"^(?:'[^']+'|[\w-]+)!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$|^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$",System.Text.RegularExpressions.RegexOptions.IgnoreCase))return Array.Empty<string>();
        var area=new AreaReference(formula,NPOI.SS.SpreadsheetVersion.EXCEL97);var values=new List<string>();string target=area.FirstCell.SheetName??sheet;
        foreach(var reference in area.GetAllReferencedCells()){var value=book.Value(book.Data.GetSheet(target).GetRow(reference.Row)?.GetCell(reference.Col));if(value!=null&&!string.IsNullOrWhiteSpace(Vb.Text(value)))values.Add(Vb.Text(value));}
        return values.Distinct().ToArray();
    }
    public object SaveParameters(string group,int row,InputChange[] changes)
    {
        var allowed=Parameters(group,row).Fields.Select(f=>f.Sheet+"!"+f.Address).ToHashSet();
        if(changes==null||changes.Any(c=>!allowed.Contains(c.Sheet+"!"+c.Address)))throw new ArgumentException("参数地址无效。");
        using(var b=Book())
        {
            foreach(var c in changes)
            {
                if(c.Value?.TrimStart().StartsWith('=')==true)throw new ArgumentException("请输入参数值，不要输入公式。");
                var target=new CellReference(c.Address);var current=((ExcelBook)b.Book).Data.GetSheet(c.Sheet).GetRow(target.Row)?.GetCell(target.Col);
                if(current?.CellType==NPOI.SS.UserModel.CellType.Formula&&!(c.Sheet=="Sec-List"&&(target.Col==5||target.Col==13)))throw new InvalidOperationException("此单元格包含公式，请修改其输入参数："+c.Address);
                b.Book.Worksheets[c.Sheet].Range[c.Address].Value2=current?.CellType!=NPOI.SS.UserModel.CellType.String&&double.TryParse(c.Value,NumberStyles.Float,CultureInfo.InvariantCulture,out double number)?(object)number:c.Value;
            }
            b.Save();
        }
        needsCalculation=true;return State();
    }
    internal static SheetOptions ReadOptions(WorkbookSession b)
    {
        var data=((ExcelBook)b.Book).Data;var sheet=data.GetSheet("CFL-App");string value=Convert.ToString(data.DocumentSummaryInformation?.CustomProperties?.Get("CFL_PrintSettings"));if(string.IsNullOrEmpty(value))value=sheet?.GetRow(0)?.GetCell(0)?.StringCellValue;
        return value==null?new():JsonSerializer.Deserialize<SheetOptions>(value)??new();
    }
    public object SaveOptions(SheetOptions options)
    {
        options.Validate();using(var b=Book()){var data=((ExcelBook)b.Book).Data;if(data.DocumentSummaryInformation==null)data.CreateInformationProperties();var properties=data.DocumentSummaryInformation.CustomProperties??new NPOI.HPSF.CustomProperties();properties.Put("CFL_PrintSettings",JsonSerializer.Serialize(options));data.DocumentSummaryInformation.CustomProperties=properties;b.Save();}return State();
    }
    public object Generate(string[] kinds,int[] rows)
    {
        if(needsCalculation)throw new InvalidOperationException("请先计算项目数据。");
        using(var b=Book())output=DrawingOutput.Generate(b,kinds,rows,ReadOptions(b),Report);return State();
    }
    public OutputEntry[] Files()=>output==null?Array.Empty<OutputEntry>():Directory.GetFiles(output,"*.dwg",SearchOption.AllDirectories).Select((p,i)=>new OutputEntry(i,Path.GetFileNameWithoutExtension(p),Path.GetFileName(Path.GetDirectoryName(p)))).ToArray();
    public string File(int id,string format)
    {
        if(format is not ("dwg" or "pdf" or "scene.json"))throw new ArgumentException("未知格式。");var files=output==null?Array.Empty<string>():Directory.GetFiles(output,"*.dwg",SearchOption.AllDirectories);
        if(id<0||id>=files.Length)throw new ArgumentOutOfRangeException(nameof(id));return Path.ChangeExtension(files[id],format);
    }
    public string Scene(int id)=>System.IO.File.ReadAllText(File(id,"scene.json"));
    public static string Probe(string parent)
    {
        var app=new OfflineProject(parent);app.Create();app.Calculate();return JsonSerializer.Serialize(app.State(),DrawingOutput.JsonOptions);
    }
}
