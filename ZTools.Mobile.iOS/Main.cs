using Foundation;
using UIKit;
using WebKit;
using System.Text.Json;
using ZTools.Offline;
namespace ZTools.Mobile.iOS;

public static class Program{static void Main(string[] args)=>UIApplication.Main(args,null,typeof(AppDelegate));}
[Register("AppDelegate")]
public sealed class AppDelegate:UIApplicationDelegate
{
    public override bool FinishedLaunching(UIApplication app,NSDictionary options)=>true;
    public override UISceneConfiguration GetConfiguration(UIApplication app,UISceneSession session,UISceneConnectionOptions options)=>new("Default",session.Role){DelegateType=typeof(SceneDelegate)};
}
[Register("SceneDelegate")]
public sealed class SceneDelegate:UIWindowSceneDelegate
{
    public override UIWindow Window{get;set;}
    public override void WillConnect(UIScene scene,UISceneSession session,UISceneConnectionOptions options)
    {Window=new UIWindow((UIWindowScene)scene){RootViewController=new MobileController()};Window.MakeKeyAndVisible();}
}
public sealed class MobileController:UIViewController,IWKScriptMessageHandler
{
    WKWebView web;MobileBridge bridge;readonly SemaphoreSlim gate=new(1);TaskCompletionSource<NSUrl> picking;
    public override void ViewDidLoad()
    {
        base.ViewDidLoad();bridge=new(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"projects"));
        var configuration=new WKWebViewConfiguration();configuration.UserContentController.AddScriptMessageHandler(this,"cfl");configuration.SetUrlSchemeHandler(new AssetsHandler(),"cfl");
        web=new WKWebView(View.Bounds,configuration){AutoresizingMask=UIViewAutoresizing.FlexibleWidth|UIViewAutoresizing.FlexibleHeight};View.AddSubview(web);web.LoadRequest(new NSUrlRequest(new NSUrl("cfl://app/web/index.html")));
    }
    public async void DidReceiveScriptMessage(WKUserContentController userContentController,WKScriptMessage message)
    {
        if(message.FrameInfo.Request.Url?.Scheme!="cfl")return;int id=0;string reply;bool entered=false;
        try
        {
            using var json=JsonDocument.Parse(message.Body.ToString());var root=json.RootElement;id=root.GetProperty("id").GetInt32();string command=root.GetProperty("command").GetString();var data=root.GetProperty("data");object result;
            if(!MobileBridge.IsImmediate(command)){await gate.WaitAsync();entered=true;}
            if(command=="import")
            {
                var url=await Pick(null);if(url==null)throw new InvalidOperationException("已取消导入。");bool scoped=url.StartAccessingSecurityScopedResource();
                try{using var stream=System.IO.File.OpenRead(url.Path);result=await Task.Run(()=>bridge.Project.Import(stream));}finally{if(scoped)url.StopAccessingSecurityScopedResource();}
            }
            else if(command=="export")
            {
                string format=data.GetProperty("format").GetString();string file=format is "xlsx" or "xls"?bridge.Project.Workbook:bridge.Project.File(data.GetProperty("id").GetInt32(),format);
                if(file==null)throw new InvalidOperationException("请先打开项目。");if(await Pick(file)==null)throw new InvalidOperationException("已取消导出。");result=new{saved=true};
            }
            else result=await Task.Run(()=>bridge.Dispatch(command,data));reply=MobileBridge.Reply(id,result);
        }
        catch(Exception error){reply=MobileBridge.Failure(id,error);}finally{if(entered)gate.Release();}
        foreach(string script in MobileBridge.Scripts(id,reply))await web.EvaluateJavaScriptAsync(script);
    }
    Task<NSUrl> Pick(string file)
    {
        picking=new(TaskCreationOptions.RunContinuationsAsynchronously);
        var picker=file==null?new UIDocumentPickerViewController(new[]{UniformTypeIdentifiers.UTTypes.Data},true):new UIDocumentPickerViewController(new[]{NSUrl.FromFilename(file)},true);
        picker.DidPickDocumentAtUrls+=(s,e)=>picking.TrySetResult(e.Urls.FirstOrDefault());picker.WasCancelled+=(s,e)=>picking.TrySetResult(null);PresentViewController(picker,true,null);return picking.Task;
    }
    sealed class AssetsHandler:NSObject,IWKUrlSchemeHandler
    {
        public void StartUrlSchemeTask(WKWebView webView,IWKUrlSchemeTask task)
        {
            try
            {
                string relative=Uri.UnescapeDataString(task.Request.Url.Path).TrimStart('/');if(!relative.StartsWith("web/")||relative.Split('/').Contains(".."))throw new IOException();
                string mime=Path.GetExtension(relative) switch{".js"=>"text/javascript",".css"=>"text/css",".jpg"=>"image/jpeg",_=>"text/html"};
                var data=NSData.FromFile(Path.Combine(NSBundle.MainBundle.BundlePath,relative));if(data==null)throw new IOException();task.DidReceiveResponse(new NSUrlResponse(task.Request.Url,mime,(nint)data.Length,"utf-8"));task.DidReceiveData(data);task.DidFinish();
            }
            catch{task.DidFailWithError(new NSError(new NSString("CFLAssets"),404));}
        }
        public void StopUrlSchemeTask(WKWebView webView,IWKUrlSchemeTask task){}
    }
}
