using ZTools.Offline;
using System.Text.Json;
string root=Directory.GetCurrentDirectory();
var builder=WebApplication.CreateBuilder(args);builder.WebHost.UseUrls("http://127.0.0.1:55420");var app=builder.Build();
var bridge=new MobileBridge(Path.Combine(root,"tests/mobile-ui-output"));var gate=new SemaphoreSlim(1);
app.MapGet("/",()=>Results.Content(File.ReadAllText(Path.Combine(root,"ZTools.Mobile.Shared/web/index.html")).Replace("<script type=\"module\"", "<script>window.CflNative={postMessage:value=>fetch('/native',{method:'POST',headers:{'Content-Type':'application/json'},body:value}).then(r=>r.text()).then(value=>{if(value.length<24000){window.cflReply(JSON.parse(value));return;}const id=JSON.parse(value).id,total=Math.ceil(value.length/24000);for(let i=0;i<total;i++)window.cflReplyPart(id,i,total,value.slice(i*24000,(i+1)*24000));})};</script><script type=\"module\""),"text/html"));
app.MapGet("/phone",()=>Results.Content("<!doctype html><html><head><meta charset='utf-8'><title>凯乐工具 · 手机界面检查</title></head><body style='margin:0;background:#e7edf2'><iframe title='手机应用界面' src='/' style='width:390px;height:844px;border:1px solid #c8d7e1;border-radius:12px;display:block;margin:14px auto;background:white'></iframe></body></html>","text/html"));
app.MapGet("/{name}",(string name)=>name is "mobile.css" or "mobile.js"?Results.File(Path.Combine(root,"ZTools.Mobile.Shared/web",name),name.EndsWith("css")?"text/css":"text/javascript"):name=="cfl-logo.jpg"?Results.File(Path.Combine(root,"ZTools.Standalone/wwwroot/cfl-logo.jpg"),"image/jpeg"):Results.NotFound());
app.MapGet("/vendor/{name}",(string name)=>name is "three.core.js" or "three.module.js"?Results.File(Path.Combine(root,"ZTools.Standalone/wwwroot/vendor",name),"text/javascript"):Results.NotFound());
app.MapPost("/native",async(JsonElement request)=>
{
    int id=request.GetProperty("id").GetInt32();string command=request.GetProperty("command").GetString();bool entered=false;
    try
    {
        if(!MobileBridge.IsImmediate(command)){await gate.WaitAsync();entered=true;}
        var result=await Task.Run(()=>bridge.Dispatch(command,request.GetProperty("data")));
        return Results.Content(MobileBridge.Reply(id,result),"application/json");
    }
    catch(Exception e){return Results.Content(MobileBridge.Failure(id,e),"application/json");}
    finally{if(entered)gate.Release();}
});
app.Run();
