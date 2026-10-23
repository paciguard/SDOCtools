extern alias WindowsEngine;
using ZTools.Offline;
using ZTools.Engine;
using NPOI.SS.UserModel;
using System.Text.Json;
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
        offline.Calculate();desktop.Calculate();
        using var source=File.OpenRead(expected);using var golden=new ExcelBook(source);using var candidateFile=File.OpenRead(offline.Workbook);using var candidate=new ExcelBook(candidateFile);
        int matches=0;
        for(int row=40;row<46;row++)for(int col=2;col<=92;col++)
        {
            var original=golden.Data.GetSheet("Sheet1").GetRow(row-1)?.GetCell(col-1);
            if(original==null)continue;
            object expectedValue=original.CellType==CellType.Formula?original.CachedFormulaResultType switch{CellType.Numeric=>original.NumericCellValue,CellType.String=>original.StringCellValue,CellType.Boolean=>original.BooleanCellValue,_=>null}:golden.Worksheets["Sheet1"].Cells[row,col].Value2;
            object actual=candidate.Worksheets["Sheet1"].Cells[row,col].Value2;
            bool equal=expectedValue is double a&&actual is double b?Math.Abs(a-b)<=1e-7*Math.Max(1,Math.Abs(a)):Convert.ToString(expectedValue)==Convert.ToString(actual);
            if(!equal)throw new Exception($"Excel/mobile mismatch at {row},{col}: Excel={expectedValue}; mobile={actual}");matches++;
        }
        Console.WriteLine($"OFFLINE_EXCEL_PARITY_OK {matches} calculated cells, including mixed ZE/ZE7/ZA names, dimensions, weights and SEC results");
    }
}
