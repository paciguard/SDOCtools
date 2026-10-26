using System.Diagnostics;
using System.Text.Json;
using ZTools.Standalone;

internal static class CalculationProcessChecks
{
    static int Main(string[] args)
    {
        if(args.FirstOrDefault()=="--child")return Child(args);
        string root=Path.GetFullPath("tests/calculation-process-output/"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        foreach(string mode in new[]{"idle","total","cancel","error","empty","conflict","success"})
        {
            string folder=Path.Combine(root,mode);Directory.CreateDirectory(folder);
            string workbook=Path.Combine(folder,"ZTools.xls");File.WriteAllText(workbook,"original");
            int childId=0;
            ProcessStartInfo Start(string staged,bool utility)
            {
                var info=new ProcessStartInfo(Environment.ProcessPath);
                if(Path.GetFileNameWithoutExtension(Environment.ProcessPath).Equals("dotnet",StringComparison.OrdinalIgnoreCase))info.ArgumentList.Add(typeof(CalculationProcessChecks).Assembly.Location);
                foreach(string arg in new[]{"--child",mode,staged,workbook})info.ArgumentList.Add(arg);
                return info;
            }
            using var cancelled=new CancellationTokenSource();
            var runner=new CalculationProcess(TimeSpan.FromMilliseconds(mode=="idle"?500:3000),TimeSpan.FromMilliseconds(mode=="total"?900:5000),Start);
            var timer=Stopwatch.StartNew();bool rejected=false;
            try
            {
                runner.Run(workbook,false,line=>
                {
                    if(line.StartsWith("PID:"))childId=int.Parse(line[4..]);
                    if(mode=="cancel")cancelled.Cancel();
                },cancelled.Token);
            }
            catch(Exception error) when(error is InvalidOperationException or TimeoutException)
            {
                rejected=true;
                if(mode=="idle"||mode=="total")Check(error is TimeoutException,"Blocked worker must time out");
                if(mode=="cancel")Check(error.Message.Contains("取消"),"Cancellation must return a clear error");
                if(mode=="conflict")Check(error.Message.Contains("其他程序修改"),"Concurrent edits must be reported");
            }
            Check(rejected==(mode!="success"),"Unexpected result for "+mode);
            Check(timer.Elapsed.TotalSeconds<8,"Worker did not release its caller: "+mode);
            string expected=mode=="success"?"finished":mode=="conflict"?"original-owner-edit":"original";
            Check(File.ReadAllText(workbook)==expected,"Project data was lost: "+mode);
            Check(!Directory.GetDirectories(folder,".cfl-calculation-*").Any(),"Private calculation folder was not removed");
            if(childId!=0)
            {
                try{using var process=Process.GetProcessById(childId);Check(process.HasExited,"Calculation worker is still running");}
                catch(ArgumentException) { }
            }
            Console.WriteLine($"PASS {mode}: returns in {timer.Elapsed.TotalSeconds:N1}s; original/owner edits preserved; worker stopped");
        }
        Console.WriteLine("CALCULATION_PROCESS_OK");return 0;
    }

    static void Check(bool value,string message){if(!value)throw new Exception(message);}
    static int Child(string[] args)
    {
        void Send(CalculationMessage message){Console.WriteLine(JsonSerializer.Serialize(message,DrawingOutput.JsonOptions));Console.Out.Flush();}
        string mode=args[1],staged=args[2],original=args[3];
        Send(new("progress","PID:"+Environment.ProcessId));
        File.WriteAllText(staged,"partial");
        if(mode=="idle"||mode=="cancel")Thread.Sleep(Timeout.Infinite);
        if(mode=="total")while(true){Send(new("progress","仍在计算"));Thread.Sleep(50);}
        if(mode=="error"){Send(new("error","数据库记录缺失"));return 1;}
        if(mode=="empty")return 0;
        if(mode=="conflict")File.AppendAllText(original,"-owner-edit");
        File.WriteAllText(staged,"finished");
        Send(new("result",Result:new(JsonSerializer.SerializeToElement(Array.Empty<object>()),new())));return 0;
    }
}
