using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text.Json;
using ZTools.Engine;

namespace ZTools.Standalone;

public sealed record CellField(string Sheet,string Address,string Label,string Value,string[] Choices);
public sealed record ParameterForm(string Title,List<CellField> Fields);
public sealed record CellChange(string Sheet,string Address,string Value);
public sealed record ActionRequest(string Path=null,string Sheet=null,string Group=null,int Row=0,string[] Kinds=null,int[] Rows=null,CellChange[] Changes=null,SheetOptions Options=null);

public sealed class StaWorker:IDisposable
{
    readonly BlockingCollection<Action> queue=new();readonly Thread thread;
    public StaWorker() { thread=new Thread(()=>{foreach(var action in queue.GetConsumingEnumerable())action();}){IsBackground=true,Name="凯乐 Excel 操作"};thread.SetApartmentState(ApartmentState.STA);thread.Start(); }
    public Task<T> Run<T>(Func<T> action) {var done=new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);queue.Add(()=>{try{done.SetResult(action());}catch(Exception ex){done.SetException(ex);}});return done.Task;}
    public void Dispose()=>queue.CompleteAdding();
}
public sealed class Workspace:IDisposable
{
    string workbookPath,editorSheet;WorkbookSession editor;object cachedRows=Array.Empty<object>();SheetOptions currentOptions=new();bool needsCalculation=true;
    public ConcurrentQueue<string> Log {get;}=new();
    public string LastOutput {get;private set;}
    public void Report(string text) {Log.Enqueue(text);while(Log.Count>120)Log.TryDequeue(out _);}
    public void EnsureIdle() {if(editor!=null)throw new InvalidOperationException("请先完成 Excel 编辑，并单击“保存并返回”。");}
    WorkbookSession Open() {EnsureIdle();if(workbookPath==null)throw new InvalidOperationException("请先新建或打开项目。");return new WorkbookSession(workbookPath,false);}
    public object State(bool refresh=false)
    {
        SheetOptions options=currentOptions;
        if(workbookPath!=null&&editor==null&&refresh)
        {
            using var book=Open();
            var input=(object[,])book.Book.Worksheets["input"].Range["D21:AD50"].Value2;
            var results=(object[,])book.Book.Worksheets["Sheet1"].Range["CA40:CF69"].Value2;
            var rows=new List<object>();
            for(int k=0;k<30;k++)if(!string.IsNullOrWhiteSpace(Vb.Text(input[k+1,11])))rows.Add(new {index=k,tag=Vb.Text(input[k+1,1]),model=Vb.Text(input[k+1,11]),material=Vb.Text(input[k+1,12]),template=Vb.Text(results[k+1,1]),list=Vb.Text(results[k+1,3])});
            cachedRows=rows;options=currentOptions=ReadOptions(book);
        }
        return new {job=workbookPath,editing=editor!=null,needsCalculation,rows=cachedRows,options,output=LastOutput,log=Log.ToArray(),templates=new[]{"GA","PID","SEC"}.Select(k=>new{kind=k,path=Template(k),exists=File.Exists(Template(k)),format=Path.GetExtension(Template(k)).ToUpperInvariant()})};
    }
    public static string Template(string kind)
    {return Path.Combine(ProjectPaths.Templates,kind+".dwg");}
    public object Create(string parent)
    {
        EnsureIdle();if(string.IsNullOrEmpty(parent)){using var picker=new FolderBrowserDialog {Description="请选择项目保存位置"};if(picker.ShowDialog()!=DialogResult.OK)return State();parent=picker.SelectedPath;}
        workbookPath=Path.Combine(ProjectPaths.CreateJob(parent),"ZTools.xls");LastOutput=null;needsCalculation=true;Report("项目已创建："+workbookPath);return State(true);
    }
    public object Select(string path)
    {
        EnsureIdle();if(string.IsNullOrEmpty(path)){using var picker=new OpenFileDialog{Title="打开项目 Excel 副本",Filter="项目工作簿|ZTools.xls"};if(picker.ShowDialog()!=DialogResult.OK)return State();path=picker.FileName;}
        path=Path.GetFullPath(path);if(ProjectPaths.IsSourceWorkbook(path))throw new InvalidOperationException("请打开项目副本。原始模板通过“新建项目”完整复制后使用。");
        using(var test=new WorkbookSession(path,false)) { }
        workbookPath=path;LastOutput=null;needsCalculation=true;Report("已打开项目："+path);return State(true);
    }
    public object Edit(string sheet)
    {
        EnsureIdle();if(!new[]{"input","PID","Sec-List","Sheet1","Utility"}.Contains(sheet))throw new InvalidOperationException("无效的编辑工作表。");
        if(workbookPath==null)throw new InvalidOperationException("请先新建或打开项目。");
        editor=new WorkbookSession(workbookPath,true);editor.ShowInput(sheet);editorSheet=sheet;Report("已打开 Excel："+sheet);return State();
    }
    public object Finish()
    {
        if(editor!=null){editor.Save();editor.Dispose();editor=null;if(editorSheet=="input")needsCalculation=true;Report("Excel 数据已保存。");}return State(true);
    }
    public object Calculate(bool utility=false)
    {
        using(var book=Open()){var engine=new SubmissionEngine(book,Report);if(utility)engine.CalculateUtility();else{engine.Calculate();needsCalculation=false;}}return State(true);
    }
    public object Refresh(){EnsureIdle();needsCalculation=true;return State(true);}
    public ParameterForm Parameters(string group,int index)
    {
        if(index<0||index>29)throw new InvalidOperationException("设备行号无效。");
        using var book=Open();var fields=new List<CellField>();
        void Field(string sheet,string col,int row,string label)
        {
            dynamic cell=book.Book.Worksheets[sheet].Range[col+row];var choices=new List<string>();
            try
            {
                if(Vb.Int(cell.Validation.Type)==3)
                {
                    string formula=Vb.Text(cell.Validation.Formula1);
                    if(formula.StartsWith("="))
                    {
                        // Resolve validation on the owning sheet so relative names
                        // and INDIRECT formulas keep the original workbook context.
                        dynamic result=book.Book.Worksheets[sheet].Evaluate(formula);
                        object values=Marshal.IsComObject(result)?result.Value2:result;
                        if(values is Array array)foreach(var value in array)choices.Add(Vb.Text(value));else choices.Add(Vb.Text(values));
                    }
                    else choices.AddRange(formula.Split(',', ';'));
                }
            }
            catch(Exception ex) when(ex is COMException or ArgumentException) { }
            fields.Add(new(sheet,col+row,label,Vb.Text(cell.Value2),choices.Where(s=>!string.IsNullOrWhiteSpace(s)).Distinct().Take(2500).ToArray()));
        }
        string title;
        switch(group)
        {
            case "equipment":
                title="设备参数";foreach(var col in new[]{"D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","X","Y","Z","AA","AB","AC","AD"})Field("input",col,21+index,book.Cell("input",col+"19"));break;
            case "motor":
                title="电机与底座选择";foreach(var col in new[]{"S","T","U","V","W"})Field("input",col,21+index,book.Cell("input",col+"19"));break;
            case "pid":
                title="管仪参数";foreach(var col in new[]{"F","G","H","I"})Field("PID",col,10+index,book.Cell("PID",col+"8"));foreach(var col in new[]{"Y","Z"})Field("input",col,21+index,book.Cell("input",col+"19"));break;
            case "materials":
                title="剖面材料选择";foreach(var col in new[]{"F","G","H","I","J","K","L","M"})Field("Sec-List",col,10+index,book.Cell("Sec-List",col+"8"));break;
            default:throw new InvalidOperationException("未知参数窗口。");
        }
        return new(title,fields);
    }
    public object SaveParameters(string group,int row,CellChange[] changes)
    {
        var allowed=Parameters(group,row).Fields.Select(f=>f.Sheet+"!"+f.Address).ToHashSet();
        if(changes==null||changes.Any(c=>!allowed.Contains(c.Sheet+"!"+c.Address)))throw new InvalidOperationException("参数地址无效。");
        using(var book=Open())
        {
            var previous=changes.Select(c=>(object)book.Book.Worksheets[c.Sheet].Range[c.Address].Formula).ToArray();
            try
            {
                foreach(var c in changes)
                {
                    if(c.Value?.TrimStart().StartsWith("=")==true)throw new InvalidOperationException("参数窗口仅接受参数值，公式请在 Excel 中编辑。");
                    book.Book.Worksheets[c.Sheet].Range[c.Address].Value2=double.TryParse(c.Value,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out double number)?(object)number:c.Value;
                }
                book.Save();
            }
            catch {for(int i=0;i<changes.Length;i++)book.Book.Worksheets[changes[i].Sheet].Range[changes[i].Address].Formula=previous[i];throw;}
        }
        if(changes.Any(c=>c.Sheet=="input"))needsCalculation=true;
        Report("参数已保存到项目 Excel；更改输入后请重新计算。");return State(true);
    }
    internal static SheetOptions ReadOptions(WorkbookSession book)
    {try{return JsonSerializer.Deserialize<SheetOptions>(Vb.Text((object)book.Book.CustomDocumentProperties["CFL_PrintSettings"].Value))??new SheetOptions();}catch(Exception ex) when(ex is COMException or ArgumentException){return new();}}
    public object SaveOptions(SheetOptions options)
    {
        options.Validate();using(var book=Open())
        {string value=JsonSerializer.Serialize(options);try{book.Book.CustomDocumentProperties["CFL_PrintSettings"].Value=value;}catch(Exception ex) when(ex is COMException or ArgumentException){book.Book.CustomDocumentProperties.Add("CFL_PrintSettings",false,4,value);}book.Save();}
        Report("打印设置已保存到项目 Excel。");return State(true);
    }
    public object Generate(string[] kinds,int[] rows,SheetOptions options)
    {
        if(needsCalculation)throw new InvalidOperationException("请先计算项目数据，确认当前 Excel 输入与出图结果一致。");
        using(var book=Open())LastOutput=DrawingOutput.Generate(book,kinds??Array.Empty<string>(),rows??Array.Empty<int>(),options??ReadOptions(book),Report);
        return State(true);
    }
    public object Files()
    {
        if(LastOutput==null)return Array.Empty<object>();
        return Directory.GetFiles(LastOutput,"*.dwg",SearchOption.AllDirectories).Select((p,i)=>new{id=i,name=Path.GetFileName(p),kind=Path.GetFileName(Path.GetDirectoryName(p))}).ToArray();
    }
    public string OutputFile(int id,string extension)
    {
        if(LastOutput==null)throw new InvalidOperationException("尚未生成图纸。");var paths=Directory.GetFiles(LastOutput,"*.dwg",SearchOption.AllDirectories);
        if(id<0||id>=paths.Length)throw new InvalidOperationException("图纸不存在。");return Path.ChangeExtension(paths[id],extension);
    }
    public void Dispose(){editor?.Dispose();editor=null;}
}
