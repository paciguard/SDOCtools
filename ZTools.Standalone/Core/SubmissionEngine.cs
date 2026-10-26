using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ZTools.Engine
{
    internal sealed partial class SubmissionEngine
    {
        private readonly WorkbookSession workbook;
        internal dynamic ExcelApp => workbook.Application;
        internal dynamic Book => workbook.Book;
        internal dynamic Drawing { get; set; }
        private readonly Action<string> report;
        private readonly Dictionary<string,object[,]> pumpDatabaseRows = new Dictionary<string,object[,]>();
        public SubmissionEngine(WorkbookSession workbook, Action<string> report)
        { this.workbook = workbook; this.report = report ?? (_ => { }); }
        private void Notify(object message) { report(Vb.Text(message)); }
        private void ChangeField(object name, object value) { Drawing.SetCustomProperty(name, value); }
        private void InsertPidPlan(object point, object value)
        {
            string plan = Vb.Text(value);
            // The supplied PID library represents these with the selected WM/WOM,
            // interference or position variants, inserted later by the drawing rules.
            if (new[] { "52", "53A", "53B", "53C", "72", "C", "S", "A" }.Contains(plan)) return;
            // The original template has no standalone Plan 61 block. Preserve its
            // original label-only handling instead of attempting an invalid insert.
            if (plan == "61") return;
            Drawing.ModelSpace.InsertBlock(point,plan,1,1,1,0);
        }
        private string Get(string sheet, int row, int col) => Vb.Text(Book.Worksheets[sheet].Cells[row, col].Value2);
        private void Put(string sheet, int row, int col, object value) { Book.Worksheets[sheet].Cells[row, col].Value2 = value; }

        public int Calculate()
        {
            // Database tables are read-only during a calculation. Share each
            // range read between preflight and the per-pump drawing lookups.
            // Clear on every run so edits made between calculations take effect.
            pumpDatabaseRows.Clear();
            int count = workbook.ValidateInput();
            // Preserve formulas, formatting, database sheets, and all user inputs.
            // Roll back output values/formulas if any source lookup fails.
            var snapshots = new[] {
                new RangeSnapshot(Book, "Sheet1", "B40:CN69"),
                new RangeSnapshot(Book, "Sheet1", "C15"),
                new RangeSnapshot(Book, "Sheet1", "DB72:DE123"),
                new RangeSnapshot(Book, "Sec-List", "F10:F39") };
            bool screen = ExcelApp.ScreenUpdating;
            object status = ExcelApp.StatusBar;
            object calculation = ExcelApp.Calculation;
            try
            {
                ExcelApp.ScreenUpdating = false;
                ExcelApp.Calculation = -4135; // xlCalculationManual
                ExcelApp.StatusBar = "凯乐工具：正在计算 " + count + " 台泵的数据";
                report("正在查询泵的尺寸和重量数据……");
                TransferData();
                report("正在重新计算工作簿公式……");
                Book.Application.CalculateFull();
                // Validate the list and finalize the material-dependent SEC block
                // now, so Sheet1 contains exactly the values later used by CAD.
                for (int k = 0; k < count; k++)
                {
                    PrepareSection(k);
                    report("剖面图第 " + (k+1) + " 台：Sheet1!CA" + (40+k) + " = " + Get("Sheet1",40+k,79)
                        + "；Sheet1!CC" + (40+k) + " = " + Get("Sheet1",40+k,81));
                }
                string[] required = { "B", "C", "D", "F", "G", "H", "J", "P", "Q", "R", "X", "Y", "AT", "AU", "CA", "CC" };
                for (int k = 0; k < count; k++)
                    foreach (string col in required)
                    {
                        string address = col + (40 + k);
                        object value = Book.Worksheets["Sheet1"].Range[address].Value2;
                        if (value is System.Runtime.InteropServices.ErrorWrapper || string.IsNullOrWhiteSpace(Vb.Text(value)))
                            throw new InvalidOperationException("计算结果缺失：Sheet1!" + address + "。请检查泵型号、电机、底座及数据库匹配情况。");
                    }
                report("正在校验冲洗方案并保存项目 Excel……");
                PreparePid(count);
                workbook.Save();
                report("已校验 " + count + " 台泵的数据，正在完成结果保存……");
                return count;
            }
            catch
            {
                foreach (var snapshot in snapshots) snapshot.Restore(Book);
                throw;
            }
            finally { ExcelApp.Calculation = calculation; ExcelApp.ScreenUpdating = screen; ExcelApp.StatusBar = status; }
        }

        public void PreparePid(int count)
        {
            Book.Worksheets["Sheet1"].Range["CG40:CM69"].ClearContents();
            Book.Worksheets["Sheet1"].Range["CG40:CM69"].NumberFormat = "@"; // plan IDs, including purely numeric codes, are text
            for (int k = 0; k < count; k++)
            {
                string plan = Regex.Replace(Get("Sheet1", 40 + k, 5), @"\s+", "");
                int digit = plan.IndexOfAny("0123456789".ToCharArray());
                if (digit < 0) throw new InvalidOperationException("Sheet1!E" + (40+k) + "：冲洗方案中未找到方案编号。");
                plan = plan.Substring(digit);
                string[] parts = plan.Split('+');
                if (parts.Length > 6 || parts.Any(string.IsNullOrWhiteSpace)) throw new InvalidOperationException("冲洗方案须包含 1 至 6 个非空组成部分：" + plan);
                Put("Sheet1", 40+k, 85, plan); // CG
                for (int p = 0; p < parts.Length; p++) Put("Sheet1", 40+k, 86+p, parts[p]);
            }
        }

        public void CalculateUtility()
        {
            int count = workbook.ValidateInput();
            var snapshot = new RangeSnapshot(Book, "Utility", "D13:P42");
            try
            {
                Book.Worksheets["Utility"].Range["D13:D42"].ClearContents();
                Book.Worksheets["Utility"].Range["F13:G42"].ClearContents();
                Book.Worksheets["Utility"].Range["M13:P42"].ClearContents();
                for (int k = 0; k < count; k++)
                {
                    string plan = Get("input", 21+k, 26).Replace(" ", "");
                    Put("Utility", 13+k, 4, plan);
                    if (plan == "" || plan == "NO" || plan == "无") Put("Utility", 13+k, 7, "无");
                    else
                    {
                        string[] parts = plan.Split('+');
                        if (parts.Length > 4) throw new InvalidOperationException("input!Z" + (21+k) + "：冷却方案最多包含 4 个组成部分。");
                        int lookupColumn = Vb.Int(Book.Worksheets["UtilityData"].Cells[6+k, 2].Value2) + 2;
                        double total = 0;
                        for (int p = 0; p < parts.Length; p++)
                        {
                            Put("Utility", 13+k, 13+p, parts[p]);
                            total += Vb.Num(ExcelApp.WorksheetFunction.VLookup(parts[p], Book.Worksheets["UtilityData"].Range["D8:L33"], lookupColumn, false));
                        }
                        Put("Utility", 13+k, 7, total.ToString(CultureInfo.InvariantCulture) + " m3/h");
                    }
                    Put("Utility", 13+k, 6, Get("Utility",13+k,3).Contains("62") ? "0.06 m3/h" : "无");
                }
                workbook.Save();
            }
            catch { snapshot.Restore(Book); throw; }
        }

        public void PrepareSection(int k)
        {
            int row = 40+k;
            string listName = Get("Sheet1", row, 81); // CC
            string materialClass = Get("Sec-List", 10+k, 6);
            int listColumn = FindSectionListColumn(listName);
            int materialColumn = FindHeader("List-detail", 2, 6, 16, Get("input", 21+k, 15));
            if (listColumn < 0) throw new InvalidOperationException("Sheet1!CC" + row + " 返回的材料清单名“" + listName + "”在 List-all 第 2 行中不存在。请检查泵版本、轴承形式及轴承架号。");
            if (materialColumn < 0) throw new InvalidOperationException("输入表第 " + (21+k) + " 行的材料等级在 List-detail 第 2 行中不存在。");
            var overrides = new Dictionary<string, int> { { "210.01",7 }, { "940.01",8 }, { "230.01",9 }, { "503.01",10 }, { "503.02",10 }, { "502.01",11 }, { "502.02",11 }, { "922.01",12 }, { "525.02",13 } };
            dynamic listSheet = Book.Worksheets["List-all"];
            var listData = (object[,])listSheet.Range[listSheet.Cells[3,listColumn],listSheet.Cells[54,listColumn]].Value2;
            var detailData = (object[,])Book.Worksheets["List-detail"].Range["C3:Q153"].Value2;
            var detailRows = new Dictionary<string,int>();
            for (int d = 1; d <= detailData.GetLength(0); d++)
            {
                string key = PartKey(Vb.Text(detailData[d,1]));
                if (key != "" && !detailRows.ContainsKey(key)) detailRows.Add(key,d);
            }
            var parts = new List<object[]>();
            for (int sourceRow = 3; sourceRow <= 54; sourceRow++)
            {
                string number = Vb.Text(listData[sourceRow-2,1]);
                if (string.IsNullOrWhiteSpace(number)) continue;
                int detailRow;
                if (!detailRows.TryGetValue(PartKey(number),out detailRow)) throw new InvalidOperationException("List-detail 中缺少零件：" + number);
                string material = Vb.Text(detailData[detailRow,materialColumn-2]);
                int overrideCol;
                if (overrides.TryGetValue(number, out overrideCol) && Get("Sec-List", 10+k, overrideCol) != "") material = Get("Sec-List", 10+k, overrideCol);
                parts.Add(new object[] { number, Vb.Text(detailData[detailRow,2]), material, Vb.Text(detailData[detailRow,3]) });
            }
            string pump = Get("Sheet1",row,3).Substring(0,2);
            bool sleeve = false;
            if (pump != "ZA" && pump != "XA" && pump != "CZ")
            {
                object[] shaft = parts.FirstOrDefault(p => Vb.Text(p[0]).StartsWith("210", StringComparison.Ordinal));
                string shaftMaterial = shaft == null ? "" : Vb.Text(shaft[2]);
                string cls = Get("input",21+k,15);
                sleeve = (shaftMaterial.Contains("Cr13") || shaftMaterial == "A276 Type 420") && (cls.Contains("A-7") || cls.Contains("D") || cls.Contains("A-8"));
                parts.RemoveAll(p => Vb.Text(p[0]) == (sleeve ? "525.02" : "400.03"));
            }
            if (parts.Count > 40) throw new InvalidOperationException("剖面零件清单共 " + parts.Count + " 项，现有明细表模板最多支持 40 项。");
            Book.Worksheets["Sheet1"].Range["DB72:DE123"].ClearContents();
            Put("Sheet1",row,83,materialClass); // CE, corrected from old column 79
            Put("Sheet1",row,84,sleeve ? "Y" : "N"); // CF, corrected from old column 82
            if (pump == "ZE" || pump == "ZF")
            {
                string template = Get("Sheet1",row,79);
                if (template.EndsWith("WO",StringComparison.OrdinalIgnoreCase)) template = template.Substring(0,template.Length-2);
                Put("Sheet1",row,79,template + (sleeve ? "" : "WO"));
            }
            for (int p = 0; p < parts.Count; p++)
                for (int c = 0; c < 4; c++) Put("Sheet1",72+p,106+c,parts[p][c]);
        }
        private int FindHeader(string sheet, int row, int first, int last, string value)
        {
            for (int col = first; col <= last; col++) if (Get(sheet,row,col) == value) return col;
            return -1;
        }
        private int FindSectionListColumn(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return -1;
            // Search actual List-all row 2, including columns appended by users.
            // Column A contains the row label; all list headers start at B.
            dynamic sheet = Book.Worksheets["List-all"];
            int last = Vb.Int(sheet.Cells[2,sheet.Columns.Count].End[-4159].Column); // xlToLeft
            if (last < 2) return -1;
            object raw = sheet.Range[sheet.Cells[2,2],sheet.Cells[2,last]].Value2;
            var headers = raw as object[,];
            int found = -1;
            for (int col = 2; col <= last; col++)
            {
                string header = Vb.Text(headers == null ? raw : headers[1,col-1]);
                if (!string.Equals(header.Trim(),(name ?? "").Trim(),StringComparison.OrdinalIgnoreCase)) continue;
                if (found >= 0) throw new InvalidOperationException("List-all 第 2 行存在重复的材料清单名称：" + name);
                found = col;
            }
            return found;
        }
        private static string PartKey(string value)
        {
            // Excel stores some part IDs as numbers (400.1) while lists use text
            // (400.10). Compare decimal IDs consistently, retain display spelling.
            decimal number;
            return decimal.TryParse(value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out number)
                ? number.ToString("0.00",CultureInfo.InvariantCulture) : value;
        }
        private string PumpDatabaseSheet(object family, object pumpIndex)
        {
            string name = Vb.Text(family);
            string standard = Regex.Replace(Get("input",21+Vb.Int(pumpIndex),5), @"\s+", "");
            return name == "ZE" && Regex.IsMatch(standard, @"^7(?=\D|$)") ? "ZE7" : name;
        }
        private int FindDatabaseOffset(object family, object size, object pumpIndex)
        {
            string sheet = Vb.Text(family), model = Vb.Text(size);
            int k = Vb.Int(pumpIndex);
            string baseplate = Get("input",21+k,21);
            if (Get("input",21+k,22) != "380V") baseplate += "(6KV)";
            object[,] data;
            if (!pumpDatabaseRows.TryGetValue(sheet,out data))
            {
                data = (object[,])Book.Worksheets[sheet].Range["A32:F2000"].Value2;
                pumpDatabaseRows.Add(sheet,data);
            }
            int found = -1;
            for (int r = 1; r <= data.GetLength(0); r++)
                if (Vb.Text(data[r,1]) == model && Vb.Text(data[r,6]) == baseplate)
                {
                    if (found >= 0) throw new InvalidOperationException(sheet + " 数据库中存在重复的泵与底座记录：" + model + " / " + baseplate + "（input 第 " + (21+k) + " 行）。");
                    found = r-1;
                }
            if (found >= 0) return found;
            throw new InvalidOperationException(sheet + " 数据库中未找到泵与底座的匹配记录：" + model + " / " + baseplate + "（输入表第 " + (21+k) + " 行）。");
        }
        private sealed class RangeSnapshot
        {
            private readonly string sheet, address;
            private readonly object formulas;
            internal RangeSnapshot(dynamic book, string sheet, string address)
            { this.sheet = sheet; this.address = address; formulas = book.Worksheets[sheet].Range[address].Formula; }
            internal void Restore(dynamic book) { book.Worksheets[sheet].Range[address].Formula = formulas; }
        }
    }
}
