using ZTools.Engine;

namespace ZTools.Calculation;

public sealed record CalculatedRow(int Index,string Tag,string Model,string Material,string Template,string List);
public sealed record CalculatedData(CalculatedRow[] Rows,string PrintSettings);

// Uses exactly the portable phone workbook adapter and migrated business rules.
// No Excel activation, DWG/PDF dependencies, or replacement business database.
public static class ManagedCalculation
{
    public static CalculatedData Run(string path,bool utility,Action<string> report,CancellationToken cancellation=default)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        report("正在读取工作簿内置数据库……");
        using var book=new WorkbookSession(path);
        var data=(ExcelBook)book.Book;data.Cancellation=cancellation;
        var engine=new SubmissionEngine(book,report);
        if(utility)engine.CalculateUtility();else engine.Calculate();
        report("正在读取计算结果……");
        var rows=new List<CalculatedRow>();
        for(int k=0;k<30;k++)
            if(!string.IsNullOrWhiteSpace(book.Cell("input","N"+(21+k))))
                rows.Add(new(k,book.Cell("input","D"+(21+k)),book.Cell("input","N"+(21+k)),book.Cell("input","O"+(21+k)),book.Cell("Sheet1","CA"+(40+k)),book.Cell("Sheet1","CC"+(40+k))));
        string options=Convert.ToString(data.Data.DocumentSummaryInformation?.CustomProperties?.Get("CFL_PrintSettings"));
        return new(rows.ToArray(),options);
    }
}
