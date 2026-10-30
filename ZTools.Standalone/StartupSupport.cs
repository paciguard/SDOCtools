using System.Diagnostics;
using System.Text.Json;

namespace ZTools.Standalone;

internal static class StartupSupport
{
    internal const string ApplicationId="CFL.SubmitTools";
    internal static string Address(string[] args)
    {
        int option=Array.IndexOf(args,"--port");
        if(option>=0)
        {
            if(option+1>=args.Length||!int.TryParse(args[option+1],out int port)||port<1||port>65535)
                throw new ArgumentException("端口必须为 1 至 65535 的整数。");
            return "http://127.0.0.1:"+port;
        }
        // Visual Studio supplies applicationUrl through this environment value.
        string configured=Environment.GetEnvironmentVariable("ASPNETCORE_URLS");
        if(!string.IsNullOrWhiteSpace(configured))
        {
            foreach(string candidate in configured.Split(';'))
                if(Uri.TryCreate(candidate.Trim(),UriKind.Absolute,out var uri)&&uri.Scheme=="http"&&uri.IsLoopback)
                    return "http://127.0.0.1:"+uri.Port;
            throw new ArgumentException("启动地址必须是本机 HTTP 地址。请使用项目中的 Visual Studio 启动配置。");
        }
        return "http://127.0.0.1:5099";
    }

    internal static async Task<bool> ReopenExisting(string url,bool openBrowser)
    {
        using var client=new HttpClient(new HttpClientHandler{UseProxy=false}){Timeout=TimeSpan.FromSeconds(2)};
        try
        {
            using var response=await client.GetAsync(url+"/api/health");
            if(!response.IsSuccessStatusCode)return false;
            using var data=JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            if(!data.RootElement.TryGetProperty("application",out var name)||name.GetString()!=ApplicationId)return false;
            if(openBrowser)OpenBrowser(url);
            Console.WriteLine("凯乐工具已在运行："+url);
            return true;
        }
        catch(Exception ex) when(ex is HttpRequestException or TaskCanceledException or JsonException){return false;}
    }

    internal static void OpenBrowser(string url)
    {
        try
        {
            // The WebGL viewer needs a modern browser even when Windows has an
            // obsolete or broken default HTTP association (such as IE).
            string edge=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),"Microsoft","Edge","Application","msedge.exe");
            if(File.Exists(edge))
            {
                var start=new ProcessStartInfo(edge){UseShellExecute=false};
                string profile=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"CFL","SubmitTools","EdgeProfile");
                start.ArgumentList.Add("--user-data-dir="+profile);
                start.ArgumentList.Add(url);
                Process.Start(start);
            }
            else Process.Start(new ProcessStartInfo(url){UseShellExecute=true});
        }
        catch(Exception ex){Log(ex);MessageBox.Show("程序已启动，请在浏览器中打开：\n"+url,"凯乐提交文件工具",MessageBoxButtons.OK,MessageBoxIcon.Information);}
    }

    internal static string Log(Exception error)
    {
        string path=Path.Combine(AppContext.BaseDirectory,"application-error.log");
        try{File.AppendAllText(path,DateTimeOffset.Now.ToString("O")+Environment.NewLine+error+Environment.NewLine);}
        catch(IOException){path=FallbackLog(error);}
        catch(UnauthorizedAccessException){path=FallbackLog(error);}
        return path;
    }
    private static string FallbackLog(Exception error)
    {
        string folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"CFL","SubmitTools");
        try{Directory.CreateDirectory(folder);string path=Path.Combine(folder,"application-error.log");File.AppendAllText(path,error+Environment.NewLine);return path;}
        catch{return "无法写入错误日志";}
    }
    internal static void Report(Exception error,bool quiet)
    {
        string log=Log(error);
        Console.Error.WriteLine(error);
        if(!quiet)MessageBox.Show("程序启动失败：\n"+error.Message+"\n\n详细日志："+log,"凯乐提交文件工具",MessageBoxButtons.OK,MessageBoxIcon.Error);
    }
}
