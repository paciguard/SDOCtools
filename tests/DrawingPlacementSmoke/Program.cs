using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;
using ZTools.Standalone;

string output=Path.GetFullPath("tests/drawing-placement-output/"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(output);
const double dx=12.6394,dy=-4.4978;
XYZ Map(XYZ point,XYZ basePoint,XYZ origin,double sx,double sy,double angle)
{
    double x=(point.X-basePoint.X)*sx,y=(point.Y-basePoint.Y)*sy;
    return new XYZ(origin.X+x*Math.Cos(angle)-y*Math.Sin(angle),origin.Y+x*Math.Sin(angle)+y*Math.Cos(angle),0);
}
bool Near(double[] actual,XYZ expected)=>Math.Abs(actual[0]-expected.X)<1e-6&&Math.Abs(actual[1]-expected.Y)<1e-6;
bool HasLine(PlotScene scene,XYZ start,XYZ end)=>scene.Items.Any(item=>item.Type=="path"&&item.Points.Length==2&&((Near(item.Points[0],start)&&Near(item.Points[1],end))||(Near(item.Points[1],start)&&Near(item.Points[0],end))));
foreach(var (name,origin,sx,sy,angle,nested) in new[]{
    ("za-cz-offset",new XYZ(dx,dy,0),1d,1d,0d,false),
    ("scaled-rotated",new XYZ(30,20,0),1.3,1.3,Math.PI/6,false),
    ("nonuniform-basepoint",new XYZ(40,15,0),1.4,.8,Math.PI/4,false),
    ("nested",new XYZ(dx,dy,0),1d,1d,0d,true)})
{
    var document=DwgReader.Read("Referrings/Ztools/TEMPLATE/GA.dwg");
    foreach(var entity in document.Entities.ToArray())document.Entities.Remove(entity);
    var block=document.BlockRecords["geduan1"];
    if(name=="nonuniform-basepoint")block.BlockEntity.BasePoint=new XYZ(10,5,0);
    XYZ basePoint=block.BlockEntity.BasePoint;
    var dimension=block.Entities.OfType<Dimension>().Single();
    var lines=dimension.Block.Entities.OfType<Line>().ToArray();
    var definition=block.Entities.OfType<AttributeDefinition>().Single(a=>a.Tag=="L1");
    var adapter=new CadAdapter(document);
    adapter.ModelSpace.InsertBlock(new[]{origin.X,origin.Y,0d},"geduan1",sx,sy,1,angle);
    var placed=document.Entities.OfType<Insert>().Single();
    placed.Attributes.Single(a=>a.Tag=="L1").Value="940";
    XYZ Expected(XYZ point)=>Map(point,basePoint,origin,sx,sy,angle);
    XYZ expectedAttribute=Expected(definition.InsertPoint);
    var attribute=placed.Attributes.Single(a=>a.Tag=="L1");
    if(!Near(new[]{attribute.InsertPoint.X,attribute.InsertPoint.Y},expectedAttribute))throw new Exception("DWG attribute placement differs: "+name);
    Func<XYZ,XYZ> expected=Expected;
    if(nested)
    {
        var parent=new BlockRecord("placement-parent");document.BlockRecords.Add(parent);
        document.Entities.Remove(placed);parent.Entities.Add(placed);
        var parentOrigin=new XYZ(25,17,0);const double parentScale=1.1,parentAngle=.25;
        adapter.ModelSpace.InsertBlock(new[]{parentOrigin.X,parentOrigin.Y,0d},parent.Name,parentScale,parentScale,1,parentAngle);
        expected=p=>Map(Expected(p),XYZ.Zero,parentOrigin,parentScale,parentScale,parentAngle);
    }
    void Verify(CadDocument doc)
    {
        var scene=DrawingOutput.Scene(doc,new());
        if(scene.Unsupported.Count>0)throw new Exception("Unexpected unsupported fixture geometry");
        foreach(var line in lines)
        {
            XYZ start=expected(line.StartPoint),end=expected(line.EndPoint);
            if(!HasLine(scene,start,end))
            {
                var closest=scene.Items.Where(item=>item.Type=="path"&&item.Points.Length==2).OrderBy(item=>Math.Pow(item.Points[0][0]-start.X,2)+Math.Pow(item.Points[0][1]-start.Y,2)).First();
                throw new Exception($"Cached dimension graphics lost placement: {name}; expected {start} -> {end}; closest {string.Join(',',closest.Points[0])} -> {string.Join(',',closest.Points[1])}");
            }
        }
        if(!scene.Items.Any(item=>item.Text=="940"&&Near(item.Points[0],expected(definition.InsertPoint))))throw new Exception("Dimension label lost placement: "+name);
    }
    Verify(document);
    string file=Path.Combine(output,name+".dwg");DwgWriter.Write(file,document);
    Verify(DwgReader.Read(file));
    Console.WriteLine("PASS dimension lines and label placement before/after DWG save: "+name);
}
if(args.Length>0)
{
    var original=DwgReader.Read("Referrings/Ztools/TEMPLATE/GA.dwg");int checkedDrawings=0;
    foreach(string file in Directory.GetFiles(args[0],"*.dwg",SearchOption.AllDirectories).Where(f=>Path.GetFileName(Path.GetDirectoryName(f))=="GA"))
    {
        var document=DwgReader.Read(file);var scene=DrawingOutput.Scene(document,new());
        foreach(var insert in document.Entities.OfType<Insert>().Where(i=>i.Block.Name=="geduan1"))
        {
            var source=original.BlockRecords[insert.Block.Name];
            XYZ Expected(XYZ p)=>Map(p,source.BlockEntity.BasePoint,insert.InsertPoint,insert.XScale,insert.YScale,insert.Rotation);
            foreach(var line in source.Entities.OfType<Dimension>().SelectMany(d=>d.Block.Entities.OfType<Line>()))
                if(!HasLine(scene,Expected(line.StartPoint),Expected(line.EndPoint)))throw new Exception("Generated GA dimension line differs from placed original template: "+file);
            var definition=source.AttributeDefinitions.Single(a=>a.Tag=="L1");
            var attribute=insert.Attributes.Single(a=>a.Tag=="L1");
            if(!Near(new[]{attribute.InsertPoint.X,attribute.InsertPoint.Y},Expected(definition.InsertPoint)))throw new Exception("Generated DWG label differs from placed definition");
            if(!scene.Items.Any(i=>i.Text==attribute.Value&&Near(i.Points[0],Expected(definition.InsertPoint))))throw new Exception("Generated GA scene label differs from saved DWG");
            checkedDrawings++;Console.WriteLine("PASS saved generated GA anchor dimensions and label: "+Path.GetFileName(file));
        }
    }
    if(checkedDrawings==0)throw new Exception("Expected at least one generated ZA/CZ GA drawing to audit");
}
Console.WriteLine("DRAWING_PLACEMENT_OK "+output);
