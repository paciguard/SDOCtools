using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.Json;
using ZTools.Standalone;

internal static class EditorLifecycleSmoke
{
    [STAThread]
    private static int Main(string[] args)
    {
        var location=new DirectoryInfo(AppContext.BaseDirectory);
        while(location!=null&&!File.Exists(Path.Combine(location.FullName,"ZTools.sln")))location=location.Parent;
        string root=location?.FullName??throw new DirectoryNotFoundException("Cannot locate the test workspace.");
        if(args.Length==2&&args[0]=="--close-editor")return CloseTestEditor(root,args[1]);
        string source=Path.Combine(root,"Referrings","Ztools","ZTools.xls");
        string folder=Path.Combine(root,"tests","editor-lifecycle-output",Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string job=null;
        var workspace=new Workspace();
        try
        {
            // Public projects contain input only; the source database stays embedded.
            var paths=typeof(Workspace).Assembly.GetType("ZTools.Engine.ProjectPaths",true);
            string projectFolder=(string)paths.GetMethod("CreateJob").Invoke(null,new object[]{folder});
            job=Path.Combine(projectFolder,"ZTools.xls");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using(var input=File.OpenRead(job))using(var project=new NPOI.HSSF.UserModel.HSSFWorkbook(input,true))
                if(project.NumberOfSheets!=1||project.GetSheetName(0)!="input")throw new Exception("New project exposes internal worksheets.");
            Console.WriteLine("PASS new project contains only the public input worksheet");
            workspace.Select(job);
            workspace.Edit("input");
            dynamic book=EditorObject(workspace,"Book");dynamic app=EditorObject(workspace,"Application");
            if(!(bool)app.DisplayAlerts)throw new Exception("Manual Excel save/close prompts must be enabled.");
            SetTag(book,"CFL-SAVED-OPEN");book.Save();
            AssertReturned(workspace.Finish(),"CFL-SAVED-OPEN",false);
            Console.WriteLine("PASS manual save with Excel open, followed by Save and return");

            workspace.Edit("input");book=EditorObject(workspace,"Book");app=EditorObject(workspace,"Application");
            SetTag(book,"CFL-SAVE-CLOSE-RECOVERY");book.Save();book.Close(false);QuitClosedApplication(app);Thread.Sleep(1000);
            AssertReturned(workspace.Finish(),"CFL-SAVE-CLOSE-RECOVERY",true);
            AssertReturned(workspace.Finish(),"CFL-SAVE-CLOSE-RECOVERY",true);
            Console.WriteLine("PASS saved and closed Excel recovers without RPC_E_DISCONNECTED or a stuck editor");

            workspace.Edit("input");book=EditorObject(workspace,"Book");
            SetTag(book,"CFL-WORKBOOK-CLOSED");book.Save();book.Close(false);
            AssertReturned(workspace.State(),"CFL-WORKBOOK-CLOSED",true);
            Console.WriteLine("PASS state refresh detects a manually closed workbook and reloads saved values");

            workspace.Edit("input");book=EditorObject(workspace,"Book");app=EditorObject(workspace,"Application");
            SetTag(book,"CFL-SAVED-BEFORE-EXIT");book.Save();SetTag(book,"UNSAVED-TEST-EDIT");
            if((int)app.Workbooks.Count!=1||!string.Equals((string)book.FullName,job,StringComparison.OrdinalIgnoreCase))throw new Exception("Refusing to terminate a non-isolated Excel process.");
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out uint processId);
            using(var owned=Process.GetProcessById((int)processId)){owned.Kill();if(!owned.WaitForExit(15000))throw new Exception("Test-owned Excel did not exit.");}
            AssertReturned(workspace.State(),"CFL-SAVED-BEFORE-EXIT",true);
            Console.WriteLine("PASS disconnected Excel reloads the saved file and warns about unsaved edits");

            workspace.Edit("input");book=EditorObject(workspace,"Book");SetTag(book,"CFL-RETRY-SAVE");
            object session=EditorSession(workspace);var property=session.GetType().GetProperty("Book",BindingFlags.NonPublic|BindingFlags.Instance);
            try
            {
                property.SetValue(session,new FailingSaveWorkbook(book));
                try{workspace.Finish();throw new Exception("Expected the save failure to propagate.");}
                catch(InvalidOperationException error)when(error.InnerException is COMException com&&com.HResult==unchecked((int)0x80030070)){}
                if(!JsonSerializer.SerializeToElement(workspace.State()).GetProperty("editing").GetBoolean())throw new Exception("A live editor was discarded after a genuine save failure.");
            }
            finally{property.SetValue(session,(object)book);}
            AssertReturned(workspace.Finish(),"CFL-RETRY-SAVE",false);
            Console.WriteLine("PASS a genuine save error preserves the editor and allows a successful retry");

            workspace.Edit("input");app=EditorObject(workspace,"Application");
            dynamic other=app.Workbooks.Add();other.Worksheets[1].Range["A1"].Value2="KEEP-OTHER-WORKBOOK";
            try
            {
                AssertReturned(workspace.Finish(),"CFL-RETRY-SAVE",false);
                if((string)other.Worksheets[1].Range["A1"].Value2!="KEEP-OTHER-WORKBOOK")throw new Exception("Another workbook was changed.");
                Console.WriteLine("PASS returning from the project leaves another open workbook intact");
            }
            finally{dynamic remainingApp=other.Application;other.Close(false);QuitClosedApplication(remainingApp);}
            workspace.Dispose();workspace.Dispose();
            File.WriteAllText(Path.Combine(root,".tools","editor-recovery-job.txt"),job);
            Console.WriteLine("EDITOR_LIFECYCLE_OK "+job);
            return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
        finally{try{workspace.Dispose();}catch(Exception error){Console.WriteLine("CLEANUP: "+error.Message);}}
    }
    private static dynamic EditorObject(Workspace workspace,string name)
    {
        object editor=EditorSession(workspace);
        return editor.GetType().GetProperty(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(editor);
    }
    private static object EditorSession(Workspace workspace)=>typeof(Workspace).GetField("editor",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(workspace);
    private static void SetTag(dynamic book,string tag){book.Worksheets["input"].Range["D21"].Value2=tag;}
    private static void AssertReturned(object result,string tag,bool recovered)
    {
        var state=JsonSerializer.SerializeToElement(result);
        if(state.GetProperty("editing").GetBoolean()||!state.GetProperty("needsCalculation").GetBoolean())throw new Exception("Workbench did not return with calculation required.");
        if(state.GetProperty("rows")[0].GetProperty("tag").GetString()!=tag)throw new Exception("Saved input value did not survive return to the workbench.");
        if(recovered&&!(state.GetProperty("notice").GetString()??"").Contains("最后保存"))throw new Exception("Recovery must explain that only saved file contents are available.");
    }
    private static void QuitClosedApplication(dynamic app)
    {try{app.Quit();}catch(COMException error)when(error.HResult==unchecked((int)0x80010108)){}catch(InvalidComObjectException){}}
    private static int CloseTestEditor(string root,string path)
    {
        path=Path.GetFullPath(path);string permitted=Path.Combine(root,"tests","editor-lifecycle-output")+Path.DirectorySeparatorChar;
        if(!path.StartsWith(permitted,StringComparison.OrdinalIgnoreCase))throw new Exception("Only a private lifecycle-test workbook may be changed by this helper.");
        dynamic book=Marshal.BindToMoniker(path);
        if((bool)book.ReadOnly||!string.Equals((string)book.FullName,path,StringComparison.OrdinalIgnoreCase))throw new Exception("The private test editor was not found read/write.");
        dynamic app=book.Application;SetTag(book,"CFL-BROWSER-RECOVERED");book.Save();book.Close(false);
        try{if((int)app.Workbooks.Count==0)QuitClosedApplication(app);}catch(COMException error)when(error.HResult==unchecked((int)0x80010108)){}
        Console.WriteLine("CLOSED_SAVED_TEST_EDITOR");return 0;
    }
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window,out uint processId);
}

// Fault injection only: a disk-full save failure must not be treated as a
// disconnected application or discard the user's live, unsaved input.
public sealed class FailingSaveWorkbook
{
    private readonly dynamic book;
    public FailingSaveWorkbook(dynamic book){this.book=book;}
    public string FullName=>(string)book.FullName;
    public void Save()=>throw new COMException("Simulated disk-full save failure.",unchecked((int)0x80030070));
}
