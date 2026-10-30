using ACadSharp;
using ACadSharp.Entities;
using CSMath;
using ZTools.Engine;

namespace ZTools.Standalone;

public sealed class CadAdapter
{
    public CadDocument Document { get; }
    public CadAdapter(CadDocument document) { Document=document; ModelSpace=new CadSpace(this); }
    public CadAdapter Application => this;
    public CadAdapter ActiveDocument => this;
    public CadAdapter Layouts => this;
    public CadSpace Block => ModelSpace;
    public CadAdapter ActiveLayout => this;
    public CadAdapter Item(object name) { if(Vb.Text(name)!="Model")throw new InvalidOperationException("仅支持模型空间。");return this; }
    public CadSpace ModelSpace { get; }
    public CadLayers Layers => new(this);
    public void SetCustomProperty(object name,object value) { if(!string.IsNullOrWhiteSpace(Vb.Text(name)))Document.SummaryInfo.Properties[Vb.Text(name)]=Vb.Text(value); }
    internal static XYZ Point(object value) { var p=(double[])value;return new XYZ(p[0],p[1],p.Length>2?p[2]:0); }
    internal static Transform BlockTransform(Insert insert)
    {
        // Block geometry is relative to its base point; attributes live in
        // the owning space after this placement has been applied once.
        return new Transform(Matrix4.GetArbitraryAxis(insert.Normal)
            * Transform.CreateTranslation(insert.InsertPoint).Matrix
            * Transform.CreateRotation(XYZ.AxisZ,insert.Rotation).Matrix
            * Transform.CreateScaling(new XYZ(insert.XScale,insert.YScale,insert.ZScale)).Matrix
            * Transform.CreateTranslation(XYZ.Zero-insert.Block.BlockEntity.BasePoint).Matrix);
    }
}
public sealed class CadSpace
{
    readonly CadAdapter cad;
    public CadSpace(CadAdapter cad)=>this.cad=cad;
    public int Count=>cad.Document.Entities.Count();
    public CadItem Item(object index)=>new(cad.Document.Entities.ElementAt(Vb.Int(index)));
    public CadItem AddMText(object point,object width,object content)
    {
        var p=CadAdapter.Point(point);
        foreach(var existing in cad.Document.Entities.OfType<MText>().Where(t=>Math.Abs(t.InsertPoint.X-p.X)+Math.Abs(t.InsertPoint.Y-p.Y)+Math.Abs(t.InsertPoint.Z-p.Z)<1e-6).ToArray())cad.Document.Entities.Remove(existing);
        var text=new MText { InsertPoint=p,RectangleWidth=Vb.Num(width),Value=Vb.Text(content),Height=2.3 };
        cad.Document.Entities.Add(text);return new CadItem(text);
    }
    public CadItem InsertBlock(object point,object name,object sx,object sy,object sz,object angle)
    {
        string key=Vb.Text(name);if(string.IsNullOrWhiteSpace(key)||key=="NONONONO")return null;
        if(!cad.Document.BlockRecords.TryGetValue(key,out var block))throw new InvalidOperationException("模板缺少图块："+key);
        var insert=new Insert(block) {InsertPoint=CadAdapter.Point(point),XScale=Vb.Num(sx),YScale=Vb.Num(sy),ZScale=Vb.Num(sz),Rotation=Vb.Num(angle)};
        cad.Document.Entities.Add(insert);
        // Create attributes from definitions after all placement values are
        // known, including rotation, scale and a nonzero block base point.
        insert.Attributes.Clear();var transform=CadAdapter.BlockTransform(insert);
        foreach(var definition in block.AttributeDefinitions)
        {
            var attribute=new AttributeEntity(definition);attribute.ApplyTransform(transform);
            insert.Attributes.Add(attribute);
        }
        return new CadItem(insert);
    }
}
public sealed class CadItem
{
    readonly Entity entity;
    public CadItem(Entity entity)=>this.entity=entity;
    public string ObjectName=>entity is Insert?"AcDbBlockReference":entity.ObjectName;
    public string Name=>((Insert)entity).Block.Name;
    public double Height {get=>((MText)entity).Height;set=>((MText)entity).Height=value;}
    public double Width {get=>((MText)entity).RectangleWidth;set=>((MText)entity).RectangleWidth=value;}
    public void Update() { }
    public void ScaleEntity(object point,object factor)
    {
        var p=CadAdapter.Point(point);double scale=Vb.Num(factor);
        if(entity is MText text) { text.InsertPoint=p+(text.InsertPoint-p)*scale;text.Height*=scale;text.RectangleWidth*=scale; }
        else throw new InvalidOperationException("无法缩放该文字实体。");
    }
    public CadAttribute[] GetAttributes()=>entity is Insert insert?insert.Attributes.Select(a=>new CadAttribute(a)).ToArray():Array.Empty<CadAttribute>();
}
public sealed class CadAttribute
{
    readonly AttributeEntity attribute;
    public CadAttribute(AttributeEntity attribute)=>this.attribute=attribute;
    public string TagString=>attribute.Tag;
    public string TextString {get=>attribute.Value;set=>attribute.Value=value??"";}
}
public sealed class CadLayers
{
    readonly CadAdapter cad;
    public CadLayers(CadAdapter cad)=>this.cad=cad;
    public CadLayer Item(object name) { if(!cad.Document.Layers.TryGetValue(Vb.Text(name),out var layer))throw new InvalidOperationException("模板缺少图层："+name);return new(layer); }
}
public sealed class CadLayer
{
    readonly ACadSharp.Tables.Layer layer;
    public CadLayer(ACadSharp.Tables.Layer layer)=>this.layer=layer;
    public bool LayerOn {get=>layer.IsOn;set=>layer.IsOn=value;}
}
