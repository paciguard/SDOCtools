using System.Diagnostics;
using System.Security.Cryptography;
using ZTools.Engine;

namespace ZTools.Offline;

public sealed partial class OfflineProject
{
    CancellationTokenSource calculationCancellation;
    public object CancelCalculation()
    {
        try{Volatile.Read(ref calculationCancellation)?.Cancel();}catch(ObjectDisposedException){}
        return new{cancelled=true};
    }

    public object Calculate(bool utility=false)
    {
        if(workbook==null)throw new InvalidOperationException("请先新建或打开项目。");
        if(!utility)needsCalculation=true;
        using var requested=new CancellationTokenSource();
        using var deadline=new CancellationTokenSource(TimeSpan.FromMinutes(3));
        using var cancellation=CancellationTokenSource.CreateLinkedTokenSource(requested.Token,deadline.Token);
        Volatile.Write(ref calculationCancellation,requested);
        long lastProgress=Stopwatch.GetTimestamp();
        using var idle=new Timer(_=>{try{if(Stopwatch.GetElapsedTime(Volatile.Read(ref lastProgress))>TimeSpan.FromSeconds(60))deadline.Cancel();}catch(ObjectDisposedException){}},null,1000,1000);
        string folder=Path.Combine(Path.GetDirectoryName(workbook),".cfl-calculation-"+Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(folder);string staged=Path.Combine(folder,"ZTools.xls");
            byte[] before=SHA256.HashData(System.IO.File.ReadAllBytes(workbook));System.IO.File.Copy(workbook,staged);
            void Progress(string text){cancellation.Token.ThrowIfCancellationRequested();Volatile.Write(ref lastProgress,Stopwatch.GetTimestamp());Report(text);}
            Progress("正在读取项目计算副本……");OfflineState result;
            using(var book=new WorkbookSession(staged))
            {
                ((ExcelBook)book.Book).Cancellation=cancellation.Token;
                var engine=new SubmissionEngine(book,Progress);
                if(utility)engine.CalculateUtility();else engine.Calculate();
                Progress("正在读取计算结果并返回工作台……");result=Snapshot(book);
            }
            cancellation.Token.ThrowIfCancellationRequested();
            using(var original=new FileStream(workbook,FileMode.Open,FileAccess.Read,FileShare.Delete))
            {
                if(!SHA256.HashData(original).SequenceEqual(before))throw new InvalidOperationException("计算期间项目 Excel 已被修改，未覆盖原文件，请重新计算。");
                System.IO.File.Replace(staged,workbook,null);
            }
            if(!utility)needsCalculation=false;
            Report(utility?"公用工程用量计算完成。":"计算完成，可以生成 DWG 和 PDF 图纸。");
            return result with{NeedsCalculation=needsCalculation,Log=log.ToArray()};
        }
        catch(OperationCanceledException)
        {
            string text=requested.IsCancellationRequested?"已取消计算，项目 Excel 保持原样。":"计算长时间未完成，已停止并返回工作台，项目 Excel 保持原样。";
            Report(text);throw new InvalidOperationException(text);
        }
        catch(Exception error){Report("计算未完成："+error.Message);throw;}
        finally
        {
            Volatile.Write(ref calculationCancellation,null);
            try{Directory.Delete(folder,true);}catch(IOException){}catch(UnauthorizedAccessException){}
        }
    }
}
