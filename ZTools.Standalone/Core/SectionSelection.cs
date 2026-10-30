using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ZTools.Engine
{
    // DWG blocks and List-all headers use different names for the same assembly.
    internal sealed class SectionSelection
    {
        internal string TemplateName { get; private set; }
        internal string ListName { get; private set; }
        internal bool IsG { get; private set; }
        internal bool LargeBearing { get; private set; }

        internal static SectionSelection Resolve(string model, string revision, string bearing, string czBearingSize)
        {
            model = Regex.Replace(model ?? "", @"\s+", "").ToUpperInvariant();
            bearing = Regex.Replace(bearing ?? "", @"\s+", "").ToUpperInvariant();
            var match = Regex.Match(model, @"^(ZE|ZA|ZF|CZ)([GO]?)(\d+)-(\d{3,4})$");
            if (!match.Success) throw new InvalidOperationException("无法识别剖面图对应的泵型号：" + model);
            string family = match.Groups[1].Value, modifier = match.Groups[2].Value, hydraulic = match.Groups[4].Value;
            var value = new SectionSelection { IsG = modifier == "G" };
            if (family == "CZ")
            {
                var frame = Regex.Match(czBearingSize ?? "", @"^LK(\d)", RegexOptions.IgnoreCase);
                if (!frame.Success) throw new InvalidOperationException("CZ 轴承架号无效：" + czBearingSize);
                value.LargeBearing = int.Parse(frame.Groups[1].Value,CultureInfo.InvariantCulture) >= 6;
                string suffix;
                if (value.LargeBearing) suffix = "WBALHL";
                else
                {
                    if (bearing != "CZ轻轴" && bearing != "CZ重轴")
                        throw new InvalidOperationException("CZ 轴承形式须选择“CZ轻轴”或“CZ重轴”。");
                    int diameter = int.Parse(hydraulic.Substring(hydraulic.Length-3),CultureInfo.InvariantCulture);
                    suffix = (diameter >= 315 ? "W" : "WO") + "BALH" + (bearing == "CZ重轴" ? "D" : "S");
                }
                value.TemplateName = value.ListName = "CZ" + suffix;
                return value;
            }
            value.LargeBearing = hydraulic.Length == 4 && hydraulic[0] >= '6';
            string root, drawingRoot;
            if (family == "ZE")
            {
                var version = Regex.Match(Regex.Replace(revision ?? "", @"\s+", ""), @"^(10|8|7)(?=\D|$)");
                if (!version.Success) throw new InvalidOperationException("ZE 剖面图版本须为 7、8 或 10：" + revision);
                drawingRoot = "ZE" + version.Value;
                root = drawingRoot + (value.IsG ? "G" : "");
            }
            else root = drawingRoot = family + modifier;
            if (bearing.Length == 0) throw new InvalidOperationException("剖面图的轴承形式不能为空。");
            value.TemplateName = drawingRoot + "_" + bearing + (value.LargeBearing ? "_K6" : "");
            value.ListName = root + (value.LargeBearing ? "(LK6)" : "") + "_" + bearing;
            return value;
        }
    }

    internal sealed partial class SubmissionEngine
    {
        internal void CalculateSectionNames(object pumpIndex)
        {
            int k = Vb.Int(pumpIndex);
            int row = 40+k;
            string model = Get("Sheet1",row,3), czSize = "";
            if (model.StartsWith("CZ",StringComparison.OrdinalIgnoreCase))
                czSize = Vb.Text(ExcelApp.WorksheetFunction.VLookup(model,Book.Worksheets["CZ-BearingSize"].Range["A2:B41"],2,false));
            SectionSelection selection;
            try { selection = SectionSelection.Resolve(model,Get("input",21+k,5),Get("Sec-List",10+k,14),czSize); }
            catch (InvalidOperationException ex) { throw new InvalidOperationException("Sec-List!N" + (10+k) + "：" + ex.Message,ex); }
            int listColumn = FindSectionListColumn(selection.ListName);
            if (listColumn < 0) throw new InvalidOperationException("Sheet1!CC" + row + " 对应的材料清单“" + selection.ListName
                + "”在 List-all 第 2 行中未找到。请检查泵版本、轴承形式及轴承架号。");
            Put("Sheet1",row,77,selection.IsG ? "Y" : "N"); // BY
            Put("Sheet1",row,78,selection.LargeBearing ? "Y" : "N"); // BZ
            Put("Sheet1",row,79,selection.TemplateName); // CA, before material-dependent WO suffix
            Put("Sheet1",row,80,Get("Sec-List",10+k,14)); // CB
            Put("Sheet1",row,81,Get("List-all",2,listColumn)); // CC: return the matched row-2 cell, not the constructed search key.
            Put("Sheet1",row,82,czSize); // CD, also clear stale CZ data on other families
        }

        private void InsertSection(object point, object worksheetRow)
        {
            int row = Vb.Int(worksheetRow);
            string name = Get("Sheet1",row,79);
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("Sheet1!CA" + row + " 的剖面图块名称为空，请重新计算项目数据。");
            Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(point,name,1,1,1,0);
            if (Get("Sheet1",row,3).StartsWith("ZE",StringComparison.OrdinalIgnoreCase) && Get("Sheet1",row,77) == "Y")
                Drawing.Application.ActiveDocument.ModelSpace.InsertBlock(point,name.StartsWith("ZE7_",StringComparison.OrdinalIgnoreCase) ? "ZE7G" : "ZEG",1,1,1,0);
        }
    }
}
