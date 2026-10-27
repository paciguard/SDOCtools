using Android.App;
using Android.Content;
using Android.OS;
using Android.Webkit;
using Android.Runtime;
using Java.Interop;
using System.Text.Json;
using ZTools.Offline;
namespace ZTools.Mobile.Android;

[Activity(Label="凯乐提交工具",MainLauncher=true,Exported=true,ConfigurationChanges=global::Android.Content.PM.ConfigChanges.Orientation|global::Android.Content.PM.ConfigChanges.ScreenSize,Theme="@android:style/Theme.Material.Light.NoActionBar")]
public sealed class MainActivity:Activity
{
    WebView view;MobileBridge bridge;readonly SemaphoreSlim gate=new(1);TaskCompletionSource<global::Android.Net.Uri> picker;
    protected override void OnCreate(Bundle state)
    {
        base.OnCreate(state);bridge=new(Path.Combine(FilesDir.AbsolutePath,"projects"));view=new WebView(this);view.SetFitsSystemWindows(true);Window.SetSoftInputMode(global::Android.Views.SoftInput.AdjustResize);
        view.Settings.JavaScriptEnabled=true;view.Settings.DomStorageEnabled=true;view.Settings.AllowFileAccess=false;view.Settings.AllowContentAccess=false;
        view.SetWebViewClient(new AssetsClient(this));view.AddJavascriptInterface(new NativeCalls(this),"CflNative");SetContentView(view);
        view.LoadUrl("https://appassets.androidplatform.net/web/index.html");
    }
    async void Receive(string request)
    {
        int id=0;string reply;bool entered=false;
        try
        {
            using var json=JsonDocument.Parse(request);var root=json.RootElement;id=root.GetProperty("id").GetInt32();string command=root.GetProperty("command").GetString();var data=root.GetProperty("data");
            if(!MobileBridge.IsImmediate(command)){await gate.WaitAsync();entered=true;}
            object result;
            if(command=="import")
            {
                var uri=await Pick(false,null);if(uri==null)throw new InvalidOperationException("已取消导入。");
                using var stream=ContentResolver.OpenInputStream(uri);result=await Task.Run(()=>bridge.Project.Import(stream));
            }
            else if(command=="export")
            {
                string format=data.GetProperty("format").GetString();string path=format is "xlsx" or "xls"?bridge.Project.Workbook:bridge.Project.File(data.GetProperty("id").GetInt32(),format);
                if(path==null)throw new InvalidOperationException("请先打开项目。");
                var uri=await Pick(true,Path.GetFileName(path));if(uri==null)throw new InvalidOperationException("已取消导出。");
                using var source=System.IO.File.OpenRead(path);using var output=ContentResolver.OpenOutputStream(uri);await source.CopyToAsync(output);result=new{saved=true};
            }
            else result=await Task.Run(()=>bridge.Dispatch(command,data));
            reply=MobileBridge.Reply(id,result);
        }
        catch(Exception error){reply=MobileBridge.Failure(id,error);}
        finally{if(entered)gate.Release();}
        RunOnUiThread(()=>{foreach(string script in MobileBridge.Scripts(id,reply))view.EvaluateJavascript(script,null);});
    }
    Task<global::Android.Net.Uri> Pick(bool export,string name)
    {
        picker=new(TaskCreationOptions.RunContinuationsAsynchronously);RunOnUiThread(()=>
        {var intent=new Intent(export?Intent.ActionCreateDocument:Intent.ActionOpenDocument);intent.AddCategory(Intent.CategoryOpenable);intent.SetType("*/*");if(export)intent.PutExtra(Intent.ExtraTitle,name);StartActivityForResult(intent,41);});return picker.Task;
    }
    protected override void OnActivityResult(int requestCode,Result resultCode,Intent data){base.OnActivityResult(requestCode,resultCode,data);if(requestCode==41)picker?.TrySetResult(resultCode==Result.Ok?data?.Data:null);}
    sealed class NativeCalls:Java.Lang.Object
    {
        readonly MainActivity owner;public NativeCalls(MainActivity owner)=>this.owner=owner;
        [JavascriptInterface,Export("postMessage")]
        public void PostMessage(string value)=>owner.Receive(value);
    }
    sealed class AssetsClient:WebViewClient
    {
        readonly Activity owner;public AssetsClient(Activity owner)=>this.owner=owner;
        public override bool ShouldOverrideUrlLoading(WebView view,IWebResourceRequest request)=>request.Url.Host!="appassets.androidplatform.net";
        public override WebResourceResponse ShouldInterceptRequest(WebView view,IWebResourceRequest request)
        {
            try
            {
                if(request.Url.Host!="appassets.androidplatform.net")return new("text/plain","utf-8",new MemoryStream());
                string path=Uri.UnescapeDataString(request.Url.Path).TrimStart('/');if(!path.StartsWith("web/")||path.Split('/').Contains(".."))throw new IOException();
                string mime=Path.GetExtension(path) switch{".js"=>"text/javascript",".css"=>"text/css",".jpg"=>"image/jpeg",_=>"text/html"};
                return new(mime,"utf-8",owner.Assets.Open(path));
            }
            catch{return new("text/plain","utf-8",new MemoryStream());}
        }
    }
}
