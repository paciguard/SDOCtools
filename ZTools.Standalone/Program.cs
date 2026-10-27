using System.Diagnostics;
using System.Security.Cryptography;
using ACadSharp.IO;
using ZTools.Engine;
using ZTools.Standalone;

if(args.Contains("--calculate-worker"))
{Environment.ExitCode=await CalculationProcess.ExecuteWorker(args);return;}
if(args.Contains("--verify-templates"))
{
    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(TemplateResources.Inventory,DrawingOutput.JsonOptions));
    return;
}
if(args.Contains("--verify-input"))
{
    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    string path=args[Array.IndexOf(args,"--verify-input")+1];
    using var source=File.OpenRead(path);using var project=new NPOI.HSSF.UserModel.HSSFWorkbook(source,true);
    if(!ProjectWorkbook.IsInputOnly(project))throw new InvalidOperationException("项目文件必须只包含 input 工作表。");
    using var model=ProjectWorkbook.Materialize(project);
    if(model.GetSheet("map")!=null||model.GetSheet("Sec-List")==null||model.GetSheet("List-detail")==null)throw new InvalidOperationException("内部数据模型不完整。");
    Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new{inputOnly=true,sheets=new[]{"input"},dropdowns=project.GetSheet("input").GetDataValidations().Count,internalSheets=model.NumberOfSheets}));return;
}
if(args.Contains("--probe"))
{
    var folder=Path.Combine(ProjectPaths.Root,"tests","standalone-probe");Directory.CreateDirectory(folder);
    foreach(var kind in new[]{"GA","PID","SEC"})
    {var notices=new List<string>();var doc=DrawingOutput.ReadTemplate(kind,notices);Console.WriteLine($"{kind}: {doc.Entities.Count()} entities, {doc.BlockRecords.Count} blocks, {notices.Count} import notices");doc.Header.Version=ACadSharp.ACadVersion.AC1032;DwgWriter.Write(Path.Combine(folder,kind+".dwg"),doc);}
    Console.WriteLine("WORKBOOK_TEMPLATE_OK "+Path.Combine(ProjectPaths.CreateJob(folder),"ZTools.xls"));
    return;
}
using var sta=new StaWorker();using var workspace=new Workspace();
if(args.Contains("--smoke"))
{
    await sta.Run(()=>
    {
        var parent=Path.Combine(ProjectPaths.Root,"tests","standalone-output");Directory.CreateDirectory(parent);
        int existingIndex=Array.IndexOf(args,"--existing-job");
        var job=existingIndex>=0?Path.GetDirectoryName(Path.GetFullPath(args[existingIndex+1])):ProjectPaths.CreateJob(parent);var log=new List<string>();
        using(var book=new WorkbookSession(Path.Combine(job,"ZTools.xls"),false))
        {var engine=new SubmissionEngine(book,s=>{log.Add(s);Console.WriteLine(s);});if(existingIndex<0)engine.Calculate();string output=DrawingOutput.Generate(book,new[]{"GA","PID","SEC"},new[]{0},new SheetOptions(),s=>{log.Add(s);Console.WriteLine(s);});Console.WriteLine("STANDALONE_SMOKE_OK "+output);}
        File.WriteAllLines(Path.Combine(job,"standalone-smoke.log"),log);return true;
    });return;
}
try
{
string url=StartupSupport.Address(args);bool openBrowser=!args.Contains("--no-browser");
if(await StartupSupport.ReopenExisting(url,openBrowser))return;
string webroot=Path.Combine(AppContext.BaseDirectory,"wwwroot");
if(!Directory.Exists(webroot))webroot=Path.Combine(ProjectPaths.Root,"ZTools.Standalone","wwwroot");
var builder=WebApplication.CreateBuilder(new WebApplicationOptions{Args=Array.Empty<string>(),ContentRootPath=Path.GetDirectoryName(webroot),WebRootPath=webroot});
builder.WebHost.UseUrls(url);
builder.Services.ConfigureHttpJsonOptions(o=>o.SerializerOptions.PropertyNamingPolicy=System.Text.Json.JsonNamingPolicy.CamelCase);
await using var app=builder.Build();string token=Convert.ToHexString(RandomNumberGenerator.GetBytes(24));
app.Use(async(context,next)=>
{
    if(context.Request.Path.StartsWithSegments("/api")&&context.Request.Method!="GET"&&context.Request.Headers["X-CFL-Token"]!=token)
    {context.Response.StatusCode=403;await context.Response.WriteAsJsonAsync(new{error="请求已失效，请刷新应用。"});return;}
    try{await next();}catch(Exception ex){context.Response.StatusCode=400;await context.Response.WriteAsJsonAsync(new{error=ex.Message});}
});
app.UseDefaultFiles();app.UseStaticFiles();
app.MapGet("/api/health",()=>new{application=StartupSupport.ApplicationId,version=typeof(Workspace).Assembly.GetName().Version?.ToString(),processId=Environment.ProcessId,executable=Environment.ProcessPath,baseDirectory=AppContext.BaseDirectory});
app.MapPost("/api/action/close",async()=>
{
    await sta.Run(()=>{workspace.EnsureIdle();return true;});
    _=Task.Run(async()=>{await Task.Delay(500);app.Lifetime.StopApplication();});
    return new{closed=true};
});
app.MapGet("/api/state",async()=>new{token,data=await sta.Run(()=>workspace.State())});
// Progress contains only the thread-safe log. Never queue it behind Excel:
// the browser must still receive updates while a COM operation is running.
app.MapGet("/api/progress",()=>new{log=workspace.Log.ToArray()});
app.MapPost("/api/action/cancel",()=>workspace.CancelCalculation());
app.MapPost("/api/action/{command}",async(string command,ActionRequest request)=>await sta.Run(()=>command switch
{
    "new"=>workspace.Create(request.Path),"open"=>workspace.Select(request.Path),
    "edit"=>workspace.Edit(request.Sheet??"input"),"finish"=>workspace.Finish(),
    "calculate"=>workspace.Calculate(),"utility"=>workspace.Calculate(true),
    "parameters"=>workspace.Parameters(request.Group,request.Row),
    "save-parameters"=>workspace.SaveParameters(request.Group,request.Row,request.Changes),
    "print"=>workspace.SaveOptions(request.Options??new()),
    "generate"=>workspace.Generate(request.Kinds,request.Rows,request.Options),
    "refresh"=>workspace.Refresh(),_=>throw new InvalidOperationException("未知操作。")
}));
app.MapGet("/api/files",()=>workspace.Files());
app.MapGet("/api/file/{id:int}/{format}",(int id,string format)=>
{
    if(!new[]{"dwg","pdf","scene"}.Contains(format))return Results.BadRequest();
    string extension=format=="scene"?"scene.json":format;string path=workspace.OutputFile(id,extension);
    return Results.File(path,format=="pdf"?"application/pdf":format=="scene"?"application/json":"application/octet-stream",format=="scene"?null:Path.GetFileName(path));
});
app.MapGet("/api/template/{kind}",(string kind)=>
{
    if(!new[]{"GA","PID","SEC"}.Contains(kind))throw new InvalidOperationException("未知模板。");
    var notices=new List<string>();return DrawingOutput.Scene(DrawingOutput.ReadTemplate(kind,notices),new());
});
app.Lifetime.ApplicationStopping.Register(()=>sta.Run(()=>{workspace.Dispose();return true;}).GetAwaiter().GetResult());
try{await app.StartAsync();}
catch(IOException)
{
    // Two launches can arrive before either has finished starting its listener.
    // Recheck the identified app after a bind conflict, without closing its job.
    if(await StartupSupport.ReopenExisting(url,openBrowser))return;
    throw;
}
if(openBrowser)StartupSupport.OpenBrowser(url);
await app.WaitForShutdownAsync();
}
catch(Exception error){StartupSupport.Report(error,args.Contains("--no-browser"));Environment.ExitCode=1;}
