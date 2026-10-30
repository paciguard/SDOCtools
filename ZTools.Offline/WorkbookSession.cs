using System;
using System.IO;
using System.Linq;
namespace ZTools.Engine;
internal sealed class WorkbookSession:IDisposable
{
 internal dynamic Application=>Book.Application;
 internal dynamic Book {get;}
 internal string FilePath {get;}
 private static readonly string[] RequiredSheets = { "Sheet1", "input", "PID", "Sec-List", "Utility", "UtilityData", "Motor", "CZ", "ZA", "ZE7", "ZE", "ZF", "nozzle-load", "Data1", "CZ-BearingSize", "List-all", "List-detail" };
 public WorkbookSession(string path,bool visible=false) {FilePath=Path.GetFullPath(path);if(Path.GetFileName(Path.GetDirectoryName(FilePath)).Equals("Ztools",StringComparison.OrdinalIgnoreCase)&&Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(FilePath))).Equals("Referrings",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("请使用项目副本。");using var stream=File.OpenRead(FilePath);Book=new ExcelBook(stream,true);foreach(var name in RequiredSheets){var sheet=Book.Worksheets[name];}}
 internal string Cell(string sheet,string address)=>Vb.Text(Book.Worksheets[sheet].Range[address].Value2);
 internal void Save()=>Book.Save(FilePath);
         internal int ValidateInput()
        {
            int count = 0; bool gap = false;
            for (int row = 21; row <= 50; row++)
            {
                string model = Cell("input", "N" + row);
                if (string.IsNullOrWhiteSpace(model)) { gap = true; continue; }
                if (gap) throw new InvalidOperationException("input!N" + row + "：泵数据必须从第 21 行开始连续填写，中间不得留空。");
                foreach (string col in new[] { "E", "O", "U", "V", "AA" })
                    if (string.IsNullOrWhiteSpace(Cell("input", col + row))) throw new InvalidOperationException("必填参数为空：input!" + col + row);
                string family = new string(model.Where(c => c != ' ' && c != '?').ToArray()).Substring(0, Math.Min(2, model.Trim().Length)).ToUpperInvariant();
                if (!new[] { "CZ", "ZA", "ZE", "ZF" }.Contains(family))
                    throw new InvalidOperationException("input!N" + row + "：现有数据库不包含该泵系列：" + family);
                string voltage = Cell("input", "V" + row);
                if (!voltage.EndsWith("V", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("input!V" + row + "：电压必须包含单位 V，例如 380V。");
                count++;
            }
            if (count == 0) throw new InvalidOperationException("请在 input!N21:N50 中至少填写一台泵的数据。");
            string standard = Cell("input", "I53");
            if (!new[] { "ASME", "HG20615", "HG20592", "SH3406", "GB9113.1-2000" }.Contains(standard))
                throw new InvalidOperationException("请在 input!I53 中选择受支持的法兰标准。");
            return count;
        }

 public void Dispose()=>((ExcelBook)Book).Dispose();
}
