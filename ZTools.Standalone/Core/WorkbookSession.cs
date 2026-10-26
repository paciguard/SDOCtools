using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace ZTools.Engine
{
    internal sealed class WorkbookSession : IDisposable
    {
        internal dynamic Application { get; private set; }
        internal dynamic Book { get; private set; }
        internal string FilePath { get; }
        private static readonly string[] RequiredSheets = { "Sheet1", "input", "PID", "Sec-List", "Utility", "UtilityData", "Motor", "CZ", "ZA", "ZE7", "ZE", "ZF", "nozzle-load", "Data1", "CZ-BearingSize", "List-all", "List-detail" };

        public WorkbookSession(string path, bool visible)
        {
            FilePath = Path.GetFullPath(path);
            if (!File.Exists(FilePath)) throw new FileNotFoundException("未找到 Excel 输入工作簿。", FilePath);
            if (ProjectPaths.IsSourceWorkbook(FilePath))
                throw new InvalidOperationException("请先新建项目副本，再编辑输入数据。");
            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null) throw new InvalidOperationException("读取工作簿内置数据库需要安装 Microsoft Excel 桌面版。");
                Application = Activator.CreateInstance(type);
                Application.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                Application.EnableEvents = false;
                Application.DisplayAlerts = false;
                Application.AskToUpdateLinks = false;
                Application.Visible = visible;
                if (!visible) Application.IgnoreRemoteRequests = true;
                Book = Application.Workbooks.Open(FilePath, UpdateLinks: 0, ReadOnly: false, Password: "123", WriteResPassword: "123", IgnoreReadOnlyRecommended: true);
                if ((bool)Book.ReadOnly) throw new IOException("项目工作簿已被打开或锁定，请先关闭后再操作。");
                foreach (string name in RequiredSheets) { dynamic sheet = Book.Worksheets[name]; Release((object)sheet); }
                // Manual Excel editing must retain its normal save/close prompts.
                Application.DisplayAlerts = visible;
            }
            catch { Dispose(); throw; }
        }

        internal void ShowInput(string sheetName = "input")
        {
            Application.Visible = true;
            Book.Worksheets[sheetName].Visible = -1;
            Book.Worksheets[sheetName].Activate();
        }
        internal void Save() { Book.Save(); }
        internal bool IsOpen
        {
            get
            {
                if (Book == null || Application == null) return false;
                try { return Application.Workbooks.Count > 0 && !string.IsNullOrEmpty((string)Book.FullName); }
                catch (Exception error) when (IsDisconnected(error)) { return false; }
                // A cell edit or a modal Excel dialog temporarily rejects calls.
                // This is still a live editor, so never discard its session.
                catch (COMException error) when (error.HResult == unchecked((int)0x80010001) || error.HResult == unchecked((int)0x8001010A)) { return true; }
            }
        }
        internal static bool IsDisconnected(Exception error)
        {
            if (error is InvalidComObjectException) return true;
            var com = error as COMException;
            if (com == null) return false;
            return com.HResult == unchecked((int)0x80010108) // RPC_E_DISCONNECTED
                || com.HResult == unchecked((int)0x800401FD) // CO_E_OBJNOTCONNECTED
                || com.HResult == unchecked((int)0x800706BA) // RPC server unavailable
                || com.HResult == unchecked((int)0x80010007) // RPC_E_SERVER_DIED
                || com.HResult == unchecked((int)0x80010012); // RPC_E_SERVER_DIED_DNE
        }
        internal string Cell(string sheet, string address) => Vb.Text(Book.Worksheets[sheet].Range[address].Value2);
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
        public void Dispose()
        {
            // Clear owned references first: a disconnected Excel process must not
            // make disposal fail repeatedly or hide the original save/open error.
            object book = Book, application = Application;
            Book = null; Application = null;
            try { if (book != null) ((dynamic)book).Close(false); }
            catch (Exception error) when (error is COMException || error is InvalidComObjectException) { System.Diagnostics.Trace.TraceWarning(error.ToString()); }
            finally { Release(book); }
            try
            {
                // The user may have opened another workbook in the visible
                // instance. Leave it open rather than quitting their Excel.
                if (application != null && ((dynamic)application).Workbooks.Count == 0) ((dynamic)application).Quit();
            }
            catch (Exception error) when (error is COMException || error is InvalidComObjectException) { System.Diagnostics.Trace.TraceWarning(error.ToString()); }
            finally { Release(application); }
            // Do not wait for all process-wide COM finalizers on this STA.
            // Book/Application are explicitly closed and released above.
        }
        private static void Release(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
            catch (Exception error) when (error is COMException || error is InvalidComObjectException) { System.Diagnostics.Trace.TraceWarning(error.ToString()); }
        }
    }

    internal static class ProjectPaths
    {
        public static string Root
        {
            get
            {
                string configured = Environment.GetEnvironmentVariable("ZTOOLS_HOME");
                if (!string.IsNullOrWhiteSpace(configured) && File.Exists(Path.Combine(configured, "Referrings", "Ztools", "ZTools.xls"))) return Path.GetFullPath(configured);
                for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir != null; dir = dir.Parent)
                    if (File.Exists(Path.Combine(dir.FullName, "Referrings", "Ztools", "ZTools.xls"))) return dir.FullName;
                return AppContext.BaseDirectory;
            }
        }
        public static string SourceWorkbook => Path.Combine(Root, "Referrings", "Ztools", "ZTools.xls");
        public static bool IsSourceWorkbook(string path)
        {
            string fullPath = Path.GetFullPath(path);
            if (string.Equals(fullPath, SourceWorkbook, StringComparison.OrdinalIgnoreCase)) return true;
            var folder = new DirectoryInfo(Path.GetDirectoryName(fullPath));
            return string.Equals(Path.GetFileName(fullPath), "ZTools.xls", StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.Name, "Ztools", StringComparison.OrdinalIgnoreCase)
                && string.Equals(folder.Parent?.Name, "Referrings", StringComparison.OrdinalIgnoreCase);
        }
        public static string CreateJob(string parent)
        {
            string path = Path.Combine(parent, "ZTools-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 6));
            Directory.CreateDirectory(path);
            using (var template = TemplateResources.Open("ZTools.xls"))
            using (var copy = new FileStream(Path.Combine(path, "ZTools.xls"), FileMode.CreateNew, FileAccess.Write))
                template.CopyTo(copy);
            return path;
        }
    }
}
