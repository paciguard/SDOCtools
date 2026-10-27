using System.Security.Cryptography;
using System.Text.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ZTools.Engine;
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
string root=Path.GetFullPath(args[0]),source=Path.Combine(root,"Referrings","Updates","ZTools-20261026.xls");
string destination=Path.Combine(root,"Resources","Workbooks","ZTools.xls"),project=Path.Combine(root,"dist","CFL-Input","ZTools.xlsx");
Directory.CreateDirectory(Path.GetDirectoryName(destination)!);Directory.CreateDirectory(Path.GetDirectoryName(project)!);
using var stream=File.OpenRead(source);using var book=new HSSFWorkbook(stream,true);
int map=book.GetSheetIndex("map");if(map>=0)book.RemoveSheetAt(map);
var section=book.GetSheet("Sec-List");section.GetRow(7).GetCell(14).SetCellValue("轴承架材质（可自行输入）");
section.GetRow(8).GetCell(14).SetCellValue("Bearing House Material\n330.01/330.02/330.03");
for(int row=9;row<39;row++)
{
 var cell=section.GetRow(row).GetCell(5)??section.GetRow(row).CreateCell(5);
 string input=book.GetSheet("input").GetRow(row+11).GetCell(14)?.ToString()??"";
 if(string.IsNullOrEmpty(cell.ToString())||cell.ToString()==input)cell.SetCellFormula($"IF(input!O{row+12}=\"\",\"\",input!O{row+12})");
}
using(var output=File.Create(destination))book.Write(output,true);
string inputTemplate=Path.Combine(root,"Resources","Workbooks","Input.xls");
// The archived XLS input template remains a historical migration fixture.
// New public projects use the owner's supplied XLSX template.
ProjectWorkbook.Create(project);
string Hash(string p)=>Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p))).ToLowerInvariant();
var record=new{source="Referrings/Updates/ZTools-20261026.xls",sourceSha256=Hash(source),embedded="Resources/Workbooks/ZTools.xls",embeddedSha256=Hash(destination),inputTemplate="Resources/Workbooks/Input.xls",inputTemplateSha256=Hash(inputTemplate),inputSheets=new[]{"input"},removedSheets=new[]{"map"},bearingHouse="Sec-List!O10:O39"};
File.WriteAllText(Path.Combine(root,"migration","workbook-split-20261026.json"),JsonSerializer.Serialize(record,new JsonSerializerOptions{WriteIndented=true}));
Console.WriteLine($"Embedded model: {destination}\nExternal input project: {project}");
