using System.Text.Json;
using ZTools.Standalone;
namespace ZTools.Offline;

public sealed class MobileBridge
{
    public OfflineProject Project{get;}
    public MobileBridge(string folder)=>Project=new(folder);
    public static bool IsImmediate(string command)=>command is "progress" or "cancel";
    public object Dispatch(string command,JsonElement data)
    {
        string Text(string key)=>data.GetProperty(key).GetString();int Number(string key)=>data.GetProperty(key).GetInt32();
        return command switch
        {
            "progress"=>Project.Progress(),"cancel"=>Project.CancelCalculation(),
            "state"=>Project.State(),"new"=>Project.Create(),"projects"=>Project.Projects(),"open"=>Project.Open(Text("id")),
            "add"=>Project.AddEquipment(),"remove"=>Project.RemoveEquipment(Number("row")),
            "parameters"=>Project.Parameters(Text("group"),Number("row")),
            "save"=>Project.SaveParameters(Text("group"),Number("row"),data.GetProperty("changes").Deserialize<InputChange[]>(DrawingOutput.JsonOptions)),
            "calculate"=>Project.Calculate(),"utility"=>Project.Calculate(true),
            "print"=>Project.SaveOptions(data.Deserialize<SheetOptions>(DrawingOutput.JsonOptions)),
            "generate"=>Project.Generate(data.GetProperty("kinds").Deserialize<string[]>(),data.GetProperty("rows").Deserialize<int[]>()),
            "scene"=>JsonSerializer.Deserialize<JsonElement>(Project.Scene(Number("id"))),
            _=>throw new ArgumentException("未知操作。")
        };
    }
    public static string Reply(int id,object data)=>JsonSerializer.Serialize(new{id,data},DrawingOutput.JsonOptions);
    public static string Failure(int id,Exception error)=>JsonSerializer.Serialize(new{id,error="操作失败："+error.Message},DrawingOutput.JsonOptions);
    public static IEnumerable<string> Scripts(int id,string reply)
    {
        // GA scenes exceed 2 MB. Keep each native WebView message comfortably
        // below platform IPC limits instead of passing one giant JavaScript call.
        const int size=24000;
        if(reply.Length<=size){yield return "window.cflReply("+reply+")";yield break;}
        int total=(reply.Length+size-1)/size;
        for(int offset=0;offset<reply.Length;offset+=size)
            yield return "window.cflReplyPart("+id+","+(offset/size)+","+total+","+JsonSerializer.Serialize(reply.Substring(offset,Math.Min(size,reply.Length-offset)))+")";
    }
}
