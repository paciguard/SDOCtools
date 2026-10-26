extern alias PortableCalculation;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using ZTools.Engine;

namespace ZTools.Standalone;

internal sealed record CalculationResult(JsonElement Rows, SheetOptions Options);
internal sealed record CalculationMessage(string Type, string Text=null, CalculationResult Result=null);

// Run the managed workbook engine on a disposable copy in a bounded child
// process. Calculation never starts Excel or waits for Excel COM cleanup.
internal sealed class CalculationProcess
{
    readonly TimeSpan idleLimit,totalLimit;
    readonly Func<string,bool,ProcessStartInfo> startInfo;
    internal CalculationProcess(TimeSpan? idleLimit=null,TimeSpan? totalLimit=null,Func<string,bool,ProcessStartInfo> startInfo=null)
    {this.idleLimit=idleLimit??TimeSpan.FromSeconds(60);this.totalLimit=totalLimit??TimeSpan.FromMinutes(3);this.startInfo=startInfo??StartInfo;}

    static ProcessStartInfo StartInfo(string workbook,bool utility)
    {
        string assembly=typeof(Workspace).Assembly.Location,executable=Path.ChangeExtension(assembly,".exe");
        var info=new ProcessStartInfo(File.Exists(executable)?executable:Environment.GetEnvironmentVariable("DOTNET_HOST_PATH")??"dotnet");
        if(!File.Exists(executable))info.ArgumentList.Add(assembly);
        info.ArgumentList.Add("--calculate-worker");info.ArgumentList.Add(workbook);
        if(utility)info.ArgumentList.Add("--utility");
        return info;
    }

    internal CalculationResult Run(string workbook,bool utility,Action<string> report,CancellationToken cancellation)
    {
        if(ProjectPaths.IsSourceWorkbook(workbook))throw new InvalidOperationException("请使用项目 Excel 副本进行计算。");
        string folder=Path.Combine(Path.GetDirectoryName(workbook),".cfl-calculation-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string staged=Path.Combine(folder,"ZTools.xls");
        var messages=new BlockingCollection<string>();
        var errors=new ConcurrentQueue<string>();
        int outputEnded=0,errorEnded=0;
        using var worker=new Process();
        try
        {
            byte[] originalHash;
            using(var input=new FileStream(workbook,FileMode.Open,FileAccess.Read,FileShare.Read))
            {
                originalHash=SHA256.HashData(input);input.Position=0;
                using var output=File.Create(staged);input.CopyTo(output);
            }
            var info=startInfo(staged,utility);
            info.UseShellExecute=false;info.CreateNoWindow=true;info.WindowStyle=ProcessWindowStyle.Hidden;
            info.RedirectStandardOutput=true;info.RedirectStandardError=true;
            worker.StartInfo=info;
            worker.OutputDataReceived+=(_,e)=>{if(e.Data!=null)messages.Add(e.Data);else Volatile.Write(ref outputEnded,1);};
            worker.ErrorDataReceived+=(_,e)=>{if(e.Data!=null){errors.Enqueue(e.Data);while(errors.Count>20)errors.TryDequeue(out string discarded);}else Volatile.Write(ref errorEnded,1);};
            cancellation.ThrowIfCancellationRequested();
            worker.Start();worker.BeginOutputReadLine();worker.BeginErrorReadLine();
            var total=Stopwatch.StartNew();var idle=Stopwatch.StartNew();
            CalculationResult result=null;string failure=null;
            void Receive(string line)
            {
                var message=JsonSerializer.Deserialize<CalculationMessage>(line,DrawingOutput.JsonOptions);
                if(message==null)throw new InvalidOperationException("计算进程未返回有效消息。");
                idle.Restart();
                if(message.Type=="progress")report(message.Text);
                else if(message.Type=="result")result=message.Result;
                else if(message.Type=="error")failure=message.Text;
            }
            while(true)
            {
                cancellation.ThrowIfCancellationRequested();
                if(total.Elapsed>totalLimit||idle.Elapsed>idleLimit)
                    throw new TimeoutException("计算长时间未响应，已停止本次计算并返回工作台，项目 Excel 保持原样。请查看错误日志后重试。");
                if(messages.TryTake(out string line,100))Receive(line);
                if(worker.HasExited&&Volatile.Read(ref outputEnded)==1&&Volatile.Read(ref errorEnded)==1)
                {
                    // EOF follows all async callbacks. Keep waiting for it in
                    // this bounded loop, never in unbounded WaitForExit().
                    while(messages.TryTake(out line))Receive(line);
                    break;
                }
            }
            if(worker.ExitCode!=0||failure!=null||result==null)
                throw new InvalidOperationException(failure??"计算进程异常退出，请查看程序错误日志。");
            cancellation.ThrowIfCancellationRequested();
            report("正在保存计算结果并返回工作台……");
            // Deny other readers/writers while checking and atomically replacing
            // the project. A separately edited workbook must never be overwritten.
            using(var original=new FileStream(workbook,FileMode.Open,FileAccess.Read,FileShare.Delete))
            {
                if(!SHA256.HashData(original).SequenceEqual(originalHash))
                    throw new InvalidOperationException("计算期间项目 Excel 已被其他程序修改，未覆盖原文件。请保存并关闭 Excel 后重新计算。");
                File.Replace(staged,workbook,null);
            }
            return result;
        }
        catch(OperationCanceledException)
        {throw new InvalidOperationException("已取消计算并返回工作台，项目 Excel 保持原样。");}
        catch(Exception error)
        {
            StartupSupport.Log(new InvalidOperationException("项目计算失败："+error.Message+Environment.NewLine+string.Join(Environment.NewLine,errors),error));
            throw;
        }
        finally
        {
            Stop(worker);
            // All paths here were created above, inside this private folder.
            try{Directory.Delete(folder,true);}catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }

    static void Stop(Process process)
    {
        if(process==null)return;
        try{if(!process.HasExited){process.Kill();process.WaitForExit(2000);}}
        catch(InvalidOperationException){}catch(System.ComponentModel.Win32Exception error){System.Diagnostics.Trace.TraceWarning(error.ToString());}
    }

    internal static Task<int> ExecuteWorker(string[] args)
    {
        void Send(CalculationMessage message){Console.WriteLine(JsonSerializer.Serialize(message,DrawingOutput.JsonOptions));Console.Out.Flush();}
        void Report(string text)=>Send(new("progress",text));
        try
        {
            var data=PortableCalculation::ZTools.Calculation.ManagedCalculation.Run(args[1],args.Contains("--utility"),Report);
            var options=string.IsNullOrWhiteSpace(data.PrintSettings)?new SheetOptions():JsonSerializer.Deserialize<SheetOptions>(data.PrintSettings);
            Send(new("result",Result:new(JsonSerializer.SerializeToElement(data.Rows,DrawingOutput.JsonOptions),options)));
            return Task.FromResult(0);
        }
        catch(Exception error){Console.Error.WriteLine(error);Send(new("error",error.Message));return Task.FromResult(1);}
    }
}
