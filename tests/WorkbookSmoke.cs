using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ZTools.Engine;

internal static class WorkbookSmoke
{
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); Console.WriteLine("PASS " + message); }
    private static string Hash(string path) { using (var sha = SHA256.Create()) using (var file = new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite)) return BitConverter.ToString(sha.ComputeHash(file)); }
    [STAThread]
    private static int Main()
    {
        try
        {
            Check(Vb.Eq(null, "") && Vb.Eq("6",6), "VBA empty and mixed numeric comparison semantics");
            Check(Vb.Mid("ZE80-2200",3,2) == "80" && Vb.Lenb("泵A") == 4, "VBA one-based Mid and UTF-16 LenB");
            CheckSection("ZE80-2200","8th","QWL","", "ZE8_QWL", "ZE8_QWL");
            CheckSection("ZE80-6200","8th","QWL","", "ZE8_QWL_K6", "ZE8(LK6)_QWL");
            CheckSection("ZE80-6200","7th","QFL","", "ZE7_QFL_K6", "ZE7(LK6)_QFL");
            CheckSection("ZE80-6200","10th","ZFL","", "ZE10_ZFL_K6", "ZE10(LK6)_ZFL");
            CheckSection("ZE80-2200","10版","QWL","", "ZE10_QWL", "ZE10_QWL");
            CheckSection("ZEG80-6200","8th","QWL","", "ZE8_QWL_K6", "ZE8G(LK6)_QWL");
            CheckSection("ZEG80-2200","7th","QFL","", "ZE7_QFL", "ZE7G_QFL");
            CheckSection("ZAG40-6200","7th","QWL","", "ZAG_QWL_K6", "ZAG(LK6)_QWL");
            CheckSection("ZA40-1200","7th","QSL","", "ZA_QSL", "ZA_QSL");
            CheckSection("ZF80-6200","8th","ZFL","", "ZF_ZFL_K6", "ZF(LK6)_ZFL");
            CheckSection("CZ50-200","8th","CZ轻轴","LK30-225", "CZWOBALHS", "CZWOBALHS");
            CheckSection("CZ50-200","8th","CZ重轴","LK30-225", "CZWOBALHD", "CZWOBALHD");
            CheckSection("CZ50-315","8th","CZ轻轴","LK40-315", "CZWBALHS", "CZWBALHS");
            CheckSection("CZ50-315","8th","CZ重轴","LK40-315", "CZWBALHD", "CZWBALHD");
            CheckSection("CZ50-315","8th","CZ重轴","LK60-315", "CZWBALHL", "CZWBALHL");
            string source = ProjectPaths.SourceWorkbook;
            string hash = Hash(source);
            string parent = Path.Combine(ProjectPaths.Root,"tests","output");
            Directory.CreateDirectory(parent);
            string job = ProjectPaths.CreateJob(parent);
            using (var session = new WorkbookSession(Path.Combine(job,"ZTools.xlsx"),false))
            {
                Check(session.Book.Worksheets.Count == 19, "Private working model retains the 19 internal/input sheets and omits map");
                var engine = new SubmissionEngine(session,Console.WriteLine);
                int count = engine.Calculate();
                Check(count == 6, "Six supplied input pump rows calculated");
                for (int k = 0; k < count; k++)
                {
                    Check(session.Cell("Sheet1","CA" + (40+k)) == (k == 5 ? "ZA_QWL" : "ZE8_QWLWO"), "SEC exact block name in CA for pump " + (k+1));
                    Check(session.Cell("Sheet1","CC" + (40+k)) == (k == 5 ? "ZA_QWL" : "ZE8_QWL"), "Material list name in CC for pump " + (k+1));
                    Check(session.Cell("Sheet1","CD" + (40+k)) == "", "Stale CZ bearing data cleared for pump " + (k+1));
                }
                session.Book.Worksheets["input"].Move(session.Book.Worksheets["Sheet1"]);
                Check(engine.Calculate() == 6 && session.Cell("Sheet1","CA40") == "ZE8_QWLWO", "Calculation uses Sheet1 by name after worksheet reordering");
                for (int k = 0; k < count; k++)
                {
                    string address = "C" + (40+k);
                    Check(session.Cell("Sheet1",address) == session.Cell("input","N" + (21+k)), "Pump mapping " + address);
                    string family = session.Cell("input","N" + (21+k)).Substring(0,2);
                    if (family == "ZE" && session.Cell("input","E" + (21+k)).StartsWith("7")) family = "ZE7";
                    var data = (object[,])session.Book.Worksheets[family].Range["A32:AE2000"].Value2;
                    string plate = session.Cell("input","U" + (21+k));
                    if (session.Cell("input","V" + (21+k)) != "380V") plate += "(6KV)";
                    string calculatedModel = session.Cell("Sheet1",address);
                    int match = Enumerable.Range(1,data.GetLength(0)).First(r => Vb.Text(data[r,1]) == calculatedModel && Vb.Text(data[r,6]) == plate);
                    for (int c = 0; c < 13; c++)
                        if (!Vb.Eq(session.Book.Worksheets["Sheet1"].Cells[40+k,24+c].Value2,data[match,7+c])) throw new Exception("Dimension mismatch at pump " + (k+1) + ", column " + (24+c));
                    double expectedMass = Enumerable.Range(6,4).Select(c => (double)Vb.Num((object)session.Book.Worksheets["Sheet1"].Cells[40+k,c].Value2)).Sum();
                    Check(Math.Abs(Vb.Num(session.Book.Worksheets["Sheet1"].Cells[40+k,10].Value2) - expectedMass) < 0.001, "13 dimensions and total mass match source database for pump " + (k+1));
                    engine.PrepareSection(k);
                    Check(session.Cell("Sheet1","DB72") != "", "Section part list populated for pump " + (k+1));
                }
                Check(session.Cell("Sheet1","CH44") == "11" && session.Cell("Sheet1","CI44") == "61", "Combined PID plans split into CH:CM");
                Check(session.Book.Worksheets["Sheet1"].Range["CH41"].Value2 is string, "Numeric PID plan identifiers remain text for mixed plan comparisons");
                engine.CalculateUtility();
                Check(session.Cell("Utility","G13").EndsWith("m3/h"), "Cooling utility looked up from workbook data");
                dynamic book = session.Book;
                // Move a matching list beyond the old BA boundary and preserve
                // the exact header cell's spelling when returning CC40.
                dynamic lists = book.Worksheets["List-all"];
                int originalColumn = Enumerable.Range(2,52).First(c => Vb.Text((object)lists.Cells[2,c].Value2) == "ZE8_QWL");
                object originalHeader = lists.Cells[2,originalColumn].Value2;
                object originalList = lists.Range[lists.Cells[2,originalColumn],lists.Cells[54,originalColumn]].Value2;
                lists.Range["BZ2:BZ54"].Value2 = originalList;
                lists.Range["BZ2"].Value2 = " ze8_qwl ";
                lists.Cells[2,originalColumn].ClearContents();
                engine.CalculateSectionNames(0); engine.PrepareSection(0);
                Check(session.Cell("Sheet1","CC40") == " ze8_qwl " && session.Cell("Sheet1","DB72") != "",
                    "CC40 returns the matching actual List-all row-2 header, including a column beyond BA");
                lists.Cells[2,originalColumn].Value2 = originalHeader;
                bool duplicateRejected = false;
                try { engine.CalculateSectionNames(0); } catch (InvalidOperationException ex) { duplicateRejected = ex.Message.Contains("重复"); }
                Check(duplicateRejected,"Ambiguous duplicate List-all row-2 names are rejected");
                lists.Range["BZ2:BZ54"].ClearContents();
                engine.CalculateSectionNames(0); engine.PrepareSection(0);
                object originalClass = book.Worksheets["input"].Range["O21"].Value2;
                object originalShaft = book.Worksheets["Sec-List"].Range["G10"].Value2;
                book.Worksheets["input"].Range["O21"].Value2 = "A-8-2";
                book.Worksheets["Sec-List"].Range["G10"].Value2 = "A276 Type 420";
                engine.PrepareSection(0);
                Check(session.Cell("Sheet1","CF40") == "Y" && session.Cell("Sheet1","CA40") == "ZE8_QWL", "Sleeve/material overrides select the regular SEC block");
                book.Worksheets["input"].Range["O21"].Value2 = originalClass;
                book.Worksheets["Sec-List"].Range["G10"].Value2 = originalShaft;
                engine.PrepareSection(0); engine.PrepareSection(0);
                Check(session.Cell("Sheet1","CA40") == "ZE8_QWLWO", "WO suffix is restored once when material requirements change");
                string previousTemplate = session.Cell("Sheet1","CA40");
                string previousList = session.Cell("Sheet1","CC40");
                string previousPart = session.Cell("Sheet1","DB72");
                object originalBearing = book.Worksheets["input"].Range["AA21"].Value2;
                book.Worksheets["input"].Range["AA21"].Value2 = "INVALID";
                bool missingListRejected = false;
                try { engine.Calculate(); } catch (InvalidOperationException ex) { missingListRejected = ex.Message.Contains("Sheet1!CC40"); }
                Check(missingListRejected && session.Cell("Sheet1","CA40") == previousTemplate && session.Cell("Sheet1","CC40") == previousList && session.Cell("Sheet1","DB72") == previousPart,
                    "Missing SEC material list reports CC40 and restores names and part-list output");
                book.Worksheets["input"].Range["AA21"].Value2 = originalBearing;
                string before = session.Cell("Sheet1","B40");
                string model = session.Cell("input","N21");
                book.Worksheets["input"].Range["N21"].Value2 = "ZE999-9999";
                bool rejected = false;
                try { engine.Calculate(); } catch (Exception) { rejected = true; }
                Check(rejected && session.Cell("Sheet1","B40") == before, "Missing database lookup fails and restores calculated ranges");
                book.Worksheets["input"].Range["N21"].Value2 = model;
                string second = session.Cell("input","N22");
                book.Worksheets["input"].Range["N22"].Value2 = "";
                rejected = false;
                try { session.ValidateInput(); } catch (InvalidOperationException) { rejected = true; }
                Check(rejected, "Input gaps rejected before calculation");
                book.Worksheets["input"].Range["N22"].Value2 = second;
                session.Save();
            }
            Check(Hash(source) == hash, "Original ZTools.xls unchanged byte for byte");
            Console.WriteLine("WORKBOOK_SMOKE_OK " + job);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static void CheckSection(string model, string revision, string bearing, string frame, string template, string list)
    {
        var result = SectionSelection.Resolve(model,revision,bearing,frame);
        Check(result.TemplateName == template && result.ListName == list,"SEC selection " + model + " / " + revision + " / " + bearing + " -> " + template + " | " + list);
    }
}
