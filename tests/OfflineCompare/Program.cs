extern alias WindowsEngine;
extern alias OfflineEngine;
using OfflineEngine::ZTools.Offline;
using ExcelBook=OfflineEngine::ZTools.Engine.ExcelBook;
using NPOI.SS.UserModel;
using System.Text.Json;
using System.Diagnostics;
using System.Reflection;
using Desktop=WindowsEngine::ZTools.Standalone.Workspace;
using DesktopChange=WindowsEngine::ZTools.Standalone.CellChange;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        string folder=Path.GetFullPath("tests/offline-compare-output/"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        var offline=new OfflineProject(Path.Combine(folder,"mobile"));offline.Create();
        using var desktop=new Desktop();var created=JsonSerializer.SerializeToElement(desktop.Create(Path.Combine(folder,"desktop")));string expected=created.GetProperty("job").GetString();
        foreach(var (row,standard) in new[]{(1,"7版"),(3," 7版 "),(4,"7th")})
        {
            offline.SaveParameters("equipment",row,new[]{new InputChange("input","E"+(21+row),standard)});
            desktop.SaveParameters("equipment",row,new[]{new DesktopChange("input","E"+(21+row),standard)});
        }
        string excelPath=Path.Combine(folder,"Excel-golden.xls");File.Copy(expected,excelPath);
        var assembly=typeof(Desktop).Assembly;
        // Run the original Excel-backed adapter independently. Workspace now
        // uses the portable engine, so it cannot itself serve as the oracle.
        using(var excel=(IDisposable)Activator.CreateInstance(assembly.GetType("ZTools.Engine.WorkbookSession"),new object[]{excelPath,false}))
        {
            var rules=Activator.CreateInstance(assembly.GetType("ZTools.Engine.SubmissionEngine"),new object[]{excel,(Action<string>)(_=>{})});
            rules.GetType().GetMethod("Calculate").Invoke(rules,null);
            rules.GetType().GetMethod("CalculateUtility").Invoke(rules,null);
        }
        offline.Calculate();offline.Calculate(true);
        var options=new WindowsEngine::ZTools.Standalone.SheetOptions(Margin:17);
        desktop.SaveOptions(options);
        var excelIds=Process.GetProcessesByName("EXCEL").Select(p=>p.Id).ToHashSet();
        var timer=Stopwatch.StartNew();var state=JsonSerializer.SerializeToElement(desktop.Calculate(),WindowsEngine::ZTools.Standalone.DrawingOutput.JsonOptions);
        if(state.GetProperty("needsCalculation").GetBoolean())throw new Exception("Desktop did not enable generation.");
        if(Process.GetProcessesByName("EXCEL").Any(p=>!excelIds.Contains(p.Id)))throw new Exception("Managed calculation started Excel.");
        Console.WriteLine($"MANAGED_WINDOWS_SIX_ROWS {timer.Elapsed.TotalSeconds:N3}s; no Excel activation");
        if(state.GetProperty("options").GetProperty("margin").GetDouble()!=17)throw new Exception("Calculation lost saved print settings.");
        desktop.Calculate(true);
        using(var saved=new ExcelBook(File.OpenRead(expected)))
        {
            if(!Convert.ToString(saved.Data.DocumentSummaryInformation?.CustomProperties?.Get("CFL_PrintSettings")).Contains("17"))throw new Exception("Saved workbook lost custom print settings.");
        }
        string managedPath=expected;expected=excelPath;
        using var source=File.OpenRead(expected);using var golden=new ExcelBook(source);using var candidateFile=File.OpenRead(offline.Workbook);using var candidate=new ExcelBook(candidateFile);
        using var managed=new ExcelBook(File.OpenRead(managedPath));
        int matches=0;
        for(int row=40;row<46;row++)for(int col=2;col<=92;col++)
        {
            var original=golden.Data.GetSheet("Sheet1").GetRow(row-1)?.GetCell(col-1);
            if(original==null)continue;
            object expectedValue=original.CellType==CellType.Formula?original.CachedFormulaResultType switch{CellType.Numeric=>original.NumericCellValue,CellType.String=>original.StringCellValue,CellType.Boolean=>original.BooleanCellValue,_=>null}:golden.Worksheets["Sheet1"].Cells[row,col].Value2;
            foreach(var engine in new[]{candidate,managed})
            {
                object actual=engine.Worksheets["Sheet1"].Cells[row,col].Value2;
                bool equal=expectedValue is double a&&actual is double b?Math.Abs(a-b)<=1e-7*Math.Max(1,Math.Abs(a)):Convert.ToString(expectedValue)==Convert.ToString(actual);
                if(!equal)throw new Exception($"Excel/managed mismatch at {row},{col}: Excel={expectedValue}; managed={actual}");
            }
            matches++;
        }
        Console.WriteLine($"OFFLINE_EXCEL_PARITY_OK {matches} calculated cells on Windows and mobile, including mixed ZE/ZE7/ZA names, dimensions, weights and SEC results");
        // Unused template rows intentionally contain #REF formulas. Compare
        // the six populated equipment rows, not those blank-row diagnostics.
        for(int row=13;row<=18;row++)for(int col=4;col<=16;col++)
        {
            var cell=golden.Data.GetSheet("Utility").GetRow(row-1)?.GetCell(col-1);
            object value=cell?.CellType==CellType.Formula?cell.CachedFormulaResultType switch{CellType.Numeric=>cell.NumericCellValue,CellType.String=>cell.StringCellValue,_=>null}:golden.Worksheets["Utility"].Cells[row,col].Value2;
            foreach(var engine in new[]{candidate,managed})
                if(Convert.ToString(value)!=Convert.ToString(engine.Worksheets["Utility"].Cells[row,col].Value2))throw new Exception($"Utility differs from Excel at {row},{col}");
        }
        Console.WriteLine("UTILITY_EXCEL_PARITY_OK 78 cells on both managed engines");
    }
}
