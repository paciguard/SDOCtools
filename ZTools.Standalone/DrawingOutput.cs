using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using CSMath;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Text.Json;
using ZTools.Engine;

namespace ZTools.Standalone;

public sealed record SheetOptions(double MinX=0,double MinY=0,double MaxX=420,double MaxY=297,double Margin=5,int Copies=1)
{
    public void Validate() { if(!double.IsFinite(MinX+MinY+MaxX+MaxY+Margin)||MaxX<=MinX||MaxY<=MinY||Margin<0||Margin>40||Copies<1||Copies>99)throw new InvalidOperationException("打印窗口、页边距或份数无效。"); }
}
public sealed record PlotItem(string Type,double[][] Points,string Text=null,double Height=0,double Angle=0,string Layer="0",bool Closed=false,string Align="left",double WidthFactor=1);
public sealed record PlotScene(List<PlotItem> Items,List<string> Unsupported,SheetOptions Frame);

public static class DrawingOutput
{
    public static CadDocument Read(string path,List<string> notices)
    {
        if(!Path.GetExtension(path).Equals(".dwg",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("仅支持 DWG 图纸模板。");
        return DwgReader.Read(path,(_,e)=>notices.Add(e.Message));
    }
    public static PlotScene Scene(CadDocument doc,SheetOptions frame)
    {
        var items=new List<PlotItem>();var unsupported=new HashSet<string>();
        double[][] Points(IEnumerable<XYZ> points)=>points.Select(p=>new[]{p.X,p.Y}).ToArray();
        void Visit(Entity entity,int depth=0)
        {
            if(depth>32)throw new InvalidOperationException("模板包含过深的嵌套图块。");
            if(entity.IsInvisible||entity.Layer is { IsOn:false })return;
            var layer=entity.Layer?.Name??"0";
            void Path(IEnumerable<XYZ> points,bool closed=false,string type="path") {var p=Points(points);if(p.Length>1)items.Add(new(type,p,Layer:layer,Closed:closed));}
            void Text(string value,XYZ p,double height,double rotation,bool top=false,string align="left",double widthFactor=1)
            {
                if(string.IsNullOrWhiteSpace(value))return;
                if(top)p=new XYZ(p.X+height*Math.Sin(rotation),p.Y-height*Math.Cos(rotation),p.Z);
                items.Add(new("text",new[]{new[]{p.X,p.Y}},value,Math.Max(.1,height),rotation,layer,Align:align,WidthFactor:widthFactor>0?widthFactor:1));
            }
            switch(entity)
            {
                case Insert insert:
                    foreach(var child in insert.Explode())Visit(child,depth+1);
                    foreach(var attribute in insert.Attributes)Visit(attribute,depth+1);
                    break;
                case AttributeDefinition: break;
                case AttributeEntity attribute when (((int)attribute.Flags)&1)!=0:break;
                case MText text: Text(WrapText(text.PlainText,text.Height,text.RectangleWidth),text.InsertPoint,text.Height,text.Rotation,true);break;
                case TextEntity text:
                    int horizontal=(int)text.HorizontalAlignment;
                    // The DWG insert point already contains the resolved baseline.
                    // Recenter changed attribute values along the text direction,
                    // preserving the stored transverse/vertical placement.
                    var anchor=text.InsertPoint;
                    if(horizontal is 1 or 2 or 4)
                    {
                        double cx=Math.Cos(text.Rotation),sy=Math.Sin(text.Rotation);
                        double distance=(text.AlignmentPoint.X-anchor.X)*cx+(text.AlignmentPoint.Y-anchor.Y)*sy;
                        anchor=new XYZ(anchor.X+distance*cx,anchor.Y+distance*sy,anchor.Z);
                    }
                    Text(text.Value,anchor,text.Height,text.Rotation,align:horizontal is 1 or 4?"center":horizontal==2?"right":"left",widthFactor:text.WidthFactor);break;
                case Line line:Path(new[]{line.StartPoint,line.EndPoint});break;
                case Leader leader:
                    Path(leader.Vertices);
                    if(leader.ArrowHeadEnabled&&leader.Vertices.Count>1)
                    {
                        var p=leader.Vertices[0];var q=leader.Vertices[1];double dx=q.X-p.X,dy=q.Y-p.Y,length=Math.Sqrt(dx*dx+dy*dy);
                        double size=Math.Abs(leader.Style.ArrowSize*leader.Style.ScaleFactor);
                        if(length>1e-9&&size>0){dx=dx/length*size;dy=dy/length*size;Path(new[]{p,new XYZ(p.X+dx-dy*.17,p.Y+dy+dx*.17,p.Z),new XYZ(p.X+dx+dy*.17,p.Y+dy-dx*.17,p.Z)},true,"fill");}
                    }
                    break;
                case Arc arc:Path(arc.PolygonalVertexes(Math.Max(12,(int)(Math.Abs(arc.Sweep)*40))));break;
                case Circle circle:Path(circle.PolygonalVertexes(128),true);break;
                case Ellipse ellipse:Path(ellipse.PolygonalVertexes(128),ellipse.IsFullEllipse);break;
                case LwPolyline poly:Path(Polyline(poly.Vertices.Select(v=>(new XYZ(v.Location.X,v.Location.Y,poly.Elevation),v.Bulge)).ToArray(),poly.IsClosed),poly.IsClosed);break;
                case Polyline2D poly:Path(Polyline(poly.Vertices.Select(v=>(v.Location,v.Bulge)).ToArray(),poly.IsClosed),poly.IsClosed);break;
                case Hatch hatch:
                    if(hatch.IsSolid) foreach(var boundary in hatch.Paths)Path(boundary.GetPoints(128),true,"fill");
                    else foreach(var line in hatch.ExplodePattern())Visit(line,depth+1);
                    break;
                case Dimension dimension when dimension.Block!=null:
                    foreach(var child in new Insert(dimension.Block).Explode())Visit(child,depth+1);break;
                case Solid solid:Path(new[]{solid.FirstCorner,solid.SecondCorner,solid.FourthCorner,solid.ThirdCorner},true,"fill");break;
                case Spline spline:Path(spline.PolygonalVertexes(128));break;
                case ACadSharp.Entities.Point:break;
                default:unsupported.Add(entity.GetType().Name);break;
            }
        }
        foreach(var entity in doc.Entities)Visit(entity);
        return new(items,unsupported.Order().ToList(),frame);
    }
    private static IEnumerable<XYZ> Polyline((XYZ P,double Bulge)[] vertices,bool closed)
    {
        for(int i=0;i<vertices.Length;i++)
        {
            var (p,b)=vertices[i];yield return p;
            if((i==vertices.Length-1&&!closed)||Math.Abs(b)<1e-10)continue;
            var q=vertices[(i+1)%vertices.Length].P;
            double dx=q.X-p.X,dy=q.Y-p.Y,length=Math.Sqrt(dx*dx+dy*dy);
            if(length<1e-10)continue;
            double cx=(p.X+q.X)/2-dy*(1-b*b)/(4*b),cy=(p.Y+q.Y)/2+dx*(1-b*b)/(4*b);
            double start=Math.Atan2(p.Y-cy,p.X-cx),sweep=4*Math.Atan(b),radius=length*(1+b*b)/(4*Math.Abs(b));
            int steps=Math.Max(4,(int)Math.Ceiling(Math.Abs(sweep)*40));
            for(int j=1;j<steps;j++)yield return new XYZ(cx+radius*Math.Cos(start+sweep*j/steps),cy+radius*Math.Sin(start+sweep*j/steps),p.Z);
        }
    }
    private static string WrapText(string value,double height,double width)
    {
        if(string.IsNullOrWhiteSpace(value)||height<=0||width<=0)return value;
        // Use the same embedded font metrics as PDF, so the WebGL scene and the
        // exported drawing share paragraph breaks within the CAD text rectangle.
        PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new CflFontResolver();
        using var document=new PdfDocument();var page=document.AddPage();using var measure=XGraphics.FromPdfPage(page);
        var font=new XFont("Microsoft YaHei",height,XFontStyleEx.Regular);var result=new List<string>();
        foreach(string paragraph in value.Replace("\r","").Split('\n'))
        {
            string rest=paragraph;
            while(rest.Length>0&&measure.MeasureString(rest,font).Width>width)
            {
                int length=1;while(length<rest.Length&&measure.MeasureString(rest[..(length+1)],font).Width<=width)length++;
                int space=rest.LastIndexOf(' ',Math.Max(0,length-1),length);if(space>0)length=space;
                result.Add(rest[..length]);rest=rest[length..].TrimStart();
            }
            result.Add(rest);
        }
        return string.Join("\n",result);
    }
    public static void Pdf(string path,PlotScene scene)
    {
        PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new CflFontResolver();
        scene.Frame.Validate();
        if(scene.Unsupported.Count>0)throw new InvalidOperationException("PDF 尚未支持这些模板实体："+string.Join("、",scene.Unsupported)+"。请先转换或检查模板。");
        using var pdf=new PdfDocument();pdf.Info.Title="凯乐提交图纸";
        for(int copy=0;copy<scene.Frame.Copies;copy++)
        {
            var page=pdf.AddPage();page.Width=PdfSharp.Drawing.XUnit.FromMillimeter(420);page.Height=PdfSharp.Drawing.XUnit.FromMillimeter(297);
            using var gfx=XGraphics.FromPdfPage(page);
            double margin=scene.Frame.Margin*72/25.4;
            double scale=Math.Min((page.Width.Point-2*margin)/(scene.Frame.MaxX-scene.Frame.MinX),(page.Height.Point-2*margin)/(scene.Frame.MaxY-scene.Frame.MinY));
            double ox=(page.Width.Point-(scene.Frame.MaxX-scene.Frame.MinX)*scale)/2;
            double oy=(page.Height.Point-(scene.Frame.MaxY-scene.Frame.MinY)*scale)/2;
            XPoint Map(double[] p)=>new(ox+(p[0]-scene.Frame.MinX)*scale,page.Height.Point-oy-(p[1]-scene.Frame.MinY)*scale);
            gfx.IntersectClip(new XRect(ox,oy,(scene.Frame.MaxX-scene.Frame.MinX)*scale,(scene.Frame.MaxY-scene.Frame.MinY)*scale));
            foreach(var item in scene.Items)
            {
                var points=item.Points.Select(Map).ToArray();
                if(item.Type=="text")
                {
                    var saved=gfx.Save();gfx.TranslateTransform(points[0].X,points[0].Y);gfx.RotateTransform(-item.Angle*180/Math.PI);
                    gfx.ScaleTransform(item.WidthFactor,1);
                    var font=new XFont("Microsoft YaHei",Math.Max(.5,item.Height*scale),XFontStyleEx.Regular);
                    var lines=(item.Text??"").Replace("\r","").Split('\n');
                    for(int n=0;n<lines.Length;n++){double width=gfx.MeasureString(lines[n],font).Width;double x=item.Align=="center"?-width/2:item.Align=="right"?-width:0;gfx.DrawString(lines[n],font,XBrushes.Black,new XPoint(x,n*item.Height*scale*1.3));}
                    gfx.Restore(saved);
                }
                else if(points.Length>1)
                {
                    if(item.Type=="fill"&&points.Length>=3)gfx.DrawPolygon(XBrushes.Black,points,XFillMode.Alternate);
                    else { gfx.DrawLines(new XPen(XColors.Black,.25),points);if(item.Closed)gfx.DrawLine(new XPen(XColors.Black,.25),points[^1],points[0]); }
                }
            }
        }
        pdf.Save(path);
    }
    internal static string Generate(WorkbookSession book,string[] kinds,int[] selected,SheetOptions options,Action<string> report)
    {
        options.Validate();var valid=new[]{"GA","PID","SEC"};
        if(kinds.Length==0||kinds.Any(k=>!valid.Contains(k)))throw new InvalidOperationException("请选择需要生成的图纸类型。");
        int count=book.ValidateInput();
        if(selected.Length==0||selected.Any(k=>k<0||k>=count))throw new InvalidOperationException("请选择有效的设备行。");
        var engine=new SubmissionEngine(book,report);engine.PreparePid(count);
        string folder=Path.Combine(Path.GetDirectoryName(book.FilePath),"Output-Standalone",DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(folder);var notices=new List<string>();object previous=book.Book.Worksheets["Sheet1"].Range["C19"].Value2;
        try
        {
            foreach(string kind in kinds.Distinct())foreach(int k in selected.Distinct())
            {
                string template=Path.Combine(ProjectPaths.Templates,kind+".dwg");
                if(!File.Exists(template))throw new FileNotFoundException("缺少图纸模板："+kind+".dwg",template);
                string name=book.Cell("Sheet1",(kind=="GA"?"P":kind=="PID"?"Q":"R")+(40+k));
                name=new string(name.Where(c=>!Path.GetInvalidFileNameChars().Contains(c)&&!char.IsControl(c)).ToArray()).Trim().TrimEnd('.');
                if(name.Length==0)throw new InvalidOperationException("请先计算项目数据，再生成图纸。");
                var document=Read(template,notices);
                book.Book.Worksheets["Sheet1"].Range["C19"].Value2=40+k;
                if(kind=="SEC")engine.PrepareSection(k);
                engine.Drawing=new CadAdapter(document);
                if(kind=="GA")engine.BuildGA();else if(kind=="PID")engine.BuildPID();else engine.BuildSEC();
                if(kind=="PID")foreach(var e in document.Entities.ToArray())
                {var box=e.GetBoundingBox();if(box.Min.X<-.01||box.Min.Y<-.01||box.Max.X>420.01||box.Max.Y>297.01)document.Entities.Remove(e);}
                var scene=Scene(document,options);
                string dir=Path.Combine(folder,kind);Directory.CreateDirectory(dir);string stem=Path.Combine(dir,name);
                if(File.Exists(stem+".dwg"))throw new InvalidOperationException("输出文件名重复："+name);
                Pdf(stem+".pdf",scene);
                document.Header.Version=ACadVersion.AC1032;DwgWriter.Write(stem+".dwg",document);
                // Reopen the actual saved DWG to verify serialization before announcing success.
                var check=DwgReader.Read(stem+".dwg");if(check.Entities.Count()!=document.Entities.Count())throw new InvalidOperationException("DWG 保存后实体数量校验失败。");
                string equipment=book.Cell("Sheet1","B"+(40+k));
                if(!check.Entities.OfType<Insert>().SelectMany(i=>i.Attributes).Any(a=>a.Tag.Equals("SHEBEIWEIHAO",StringComparison.OrdinalIgnoreCase)&&a.Value==equipment))throw new InvalidOperationException("DWG 设备位号校验失败。");
                if(!scene.Items.Any(i=>i.Type=="text"&&i.Text==equipment))throw new InvalidOperationException("PDF 设备位号校验失败。");
                if(kind=="SEC")
                {
                    string firstPart=book.Cell("Sheet1","DB72");
                    if(!check.Entities.OfType<Insert>().SelectMany(i=>i.Attributes).Any(a=>a.Tag.Equals("PN1",StringComparison.OrdinalIgnoreCase)&&a.Value==firstPart))throw new InvalidOperationException("DWG 材料明细校验失败。");
                }
                File.WriteAllText(stem+".scene.json",JsonSerializer.Serialize(scene,JsonOptions));
                report(kind+" 第 "+(k+1)+" 台："+stem);
            }
            return folder;
        }
        finally
        {
            book.Book.Worksheets["Sheet1"].Range["C19"].Value2=previous;book.Save();
            File.WriteAllLines(Path.Combine(folder,"compatibility.log"),notices.Distinct());
        }
    }
    public static readonly JsonSerializerOptions JsonOptions=new(){PropertyNamingPolicy=JsonNamingPolicy.CamelCase,WriteIndented=false};
}

internal sealed class CflFontResolver:PdfSharp.Fonts.IFontResolver
{
    readonly string path;
    public CflFontResolver()
    {
        string fonts=Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        path=new[]{"STXIHEI.TTF","simhei.ttf","simfang.ttf"}.Select(n=>Path.Combine(fonts,n)).FirstOrDefault(File.Exists)
            ?? throw new InvalidOperationException("未找到可嵌入的中文 TrueType 字体。请安装黑体、华文细黑或仿宋字体后再导出 PDF。");
    }
    public byte[] GetFont(string faceName)=>File.ReadAllBytes(path);
    public PdfSharp.Fonts.FontResolverInfo ResolveTypeface(string familyName,bool bold,bool italic)=>new("cfl-cjk");
}
