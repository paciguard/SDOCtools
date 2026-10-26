using System.Collections;
using System.Globalization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ZTools.Engine;

// A deliberately limited Excel object model for the migrated VBA rules. All
// reads and writes remain in the original BIFF8 workbook; no COM or Excel process.
public sealed class ExcelBook : IDisposable
{
    internal readonly HSSFWorkbook Data;
    internal readonly IFormulaEvaluator Evaluator;
    internal readonly DataFormatter Formatter=new(CultureInfo.InvariantCulture);
    internal CancellationToken Cancellation{get;set;}
    public ExcelSheets Worksheets {get;}
    public ExcelApplication Application {get;}
    public ExcelBook(Stream stream)
    {
        Data=new HSSFWorkbook(stream,true); // Preserve VBA and other OLE nodes.
        Evaluator=Data.GetCreationHelper().CreateFormulaEvaluator();
        Worksheets=new(this);Application=new(this);
    }
    internal object Value(ICell cell)
    {
        Cancellation.ThrowIfCancellationRequested();
        if(cell==null)return null;
        if(cell.CellType==CellType.Formula)
        {
            CellType result;
            // Evaluate once and retain the cached BIFF result for Excel readers.
            try{result=Evaluator.EvaluateFormulaCell(cell);}catch(OperationCanceledException){throw;}catch(Exception e){throw new InvalidOperationException($"无法计算 {cell.Sheet.SheetName}!{cell.Address}：{cell.CellFormula}",e);}
            return result switch {CellType.Numeric=>cell.NumericCellValue,CellType.String=>cell.StringCellValue,CellType.Boolean=>cell.BooleanCellValue,CellType.Blank=>null,_=>throw new InvalidOperationException($"公式错误：{cell.Sheet.SheetName}!{cell.Address}（{cell.ErrorCellValue}）")};
        }
        return cell.CellType switch {CellType.Numeric=>cell.NumericCellValue,CellType.String=>cell.StringCellValue,CellType.Boolean=>cell.BooleanCellValue,CellType.Error=>throw new InvalidOperationException($"单元格错误：{cell.Sheet.SheetName}!{cell.Address}"),_=>null};
    }
    public void Save(string path)
    {
        string pending=path+".saving";
        try{Cancellation.ThrowIfCancellationRequested();using(var output=new FileStream(pending,FileMode.Create,FileAccess.Write))Data.Write(output,true);Cancellation.ThrowIfCancellationRequested();File.Move(pending,path,true);}
        finally{if(File.Exists(pending))File.Delete(pending);}
    }
    public void Dispose()=>Data.Close();
}
public sealed class ExcelSheets
{
    readonly ExcelBook book;public ExcelSheets(ExcelBook book)=>this.book=book;
    public ExcelSheet this[object name]
    {get{var sheet=name is string s?book.Data.GetSheet(s):book.Data.GetSheetAt(Convert.ToInt32(name)-1);return new(book,sheet??throw new InvalidOperationException("缺少工作表："+name));}}
}
public sealed class ExcelSheet
{
    internal readonly ExcelBook Book;internal readonly ISheet Data;
    public ExcelRanges Range {get;}
    public ExcelCells Cells {get;}
    public ExcelColumnCount Columns{get;}=new();
    public ExcelSheet(ExcelBook book,ISheet data){Book=book;Data=data;Range=new(this);Cells=new(this);}
    public void Activate() { } // Legacy selection has no effect on calculation.
}
public sealed class ExcelCells
{
    readonly ExcelSheet sheet;public ExcelCells(ExcelSheet sheet)=>this.sheet=sheet;
    public ExcelRange this[object row,object column]=>new(sheet,Convert.ToInt32(row),Convert.ToInt32(column),Convert.ToInt32(row),Convert.ToInt32(column));
}
public sealed class ExcelRanges
{
    readonly ExcelSheet sheet;public ExcelRanges(ExcelSheet sheet)=>this.sheet=sheet;
    public ExcelRange this[object address]
    {get{var area=AreaReference.GenerateContiguous(NPOI.SS.SpreadsheetVersion.EXCEL97,Convert.ToString(address).Replace("$",""));if(area.Length!=1)throw new ArgumentException("不支持不连续区域。");var a=area[0];return new(sheet,a.FirstCell.Row+1,a.FirstCell.Col+1,a.LastCell.Row+1,a.LastCell.Col+1);}}
    public ExcelRange this[ExcelRange first,ExcelRange last]=>new(sheet,first.Row,first.Column,last.LastRow,last.LastColumn);
}
public sealed class ExcelRange : IEnumerable<ExcelRange>
{
    internal readonly ExcelSheet Sheet;
    public int Row{get;}public int Column{get;}public int LastRow{get;}public int LastColumn{get;}
    public ExcelEnd End=>new(this);
    public ExcelRange(ExcelSheet sheet,int row,int column,int lastRow,int lastColumn){Sheet=sheet;Row=row;Column=column;LastRow=lastRow;LastColumn=lastColumn;}
    internal ICell Cell(int row,int col,bool create=false){var r=Sheet.Data.GetRow(row-1);if(r==null&&create)r=Sheet.Data.CreateRow(row-1);return create?r.GetCell(col-1)??r.CreateCell(col-1):r?.GetCell(col-1);}
    object Read(bool formula)
    {
        Sheet.Book.Cancellation.ThrowIfCancellationRequested();
        object One(int r,int c){var cell=Cell(r,c);return formula&&cell?.CellType==CellType.Formula?"="+cell.CellFormula:Sheet.Book.Value(cell);}
        if(Row==LastRow&&Column==LastColumn)return One(Row,Column);
        var values=Array.CreateInstance(typeof(object),new[]{LastRow-Row+1,LastColumn-Column+1},new[]{1,1});
        for(int r=Row;r<=LastRow;r++)for(int c=Column;c<=LastColumn;c++)values.SetValue(One(r,c),r-Row+1,c-Column+1);
        return values;
    }
    void Write(object value,bool formula)
    {
        for(int r=Row;r<=LastRow;r++)for(int c=Column;c<=LastColumn;c++)
        {
            Sheet.Book.Cancellation.ThrowIfCancellationRequested();
            object v=value is Array a?a.GetValue(r-Row+a.GetLowerBound(0),c-Column+a.GetLowerBound(1)):value;
            var cell=Cell(r,c,true);cell.SetCellType(CellType.Blank);
            if(formula&&v is string f&&f.StartsWith('='))cell.SetCellFormula(f[1..]);
            else if(v==null){}else if(v is bool b)cell.SetCellValue(b);else if(v is string text)cell.SetCellValue(text);else cell.SetCellValue(Convert.ToDouble(v,CultureInfo.InvariantCulture));
        }
        Sheet.Book.Evaluator.ClearAllCachedResultValues();
    }
    public object Value2{get=>Read(false);set=>Write(value,false);}
    public object Formula{get=>Read(true);set=>Write(value,true);}
    public string Text=>Sheet.Book.Formatter.FormatCellValue(Cell(Row,Column),Sheet.Book.Evaluator);
    // Values are assigned as strings by the migrated plan rules. Their BIFF cell
    // types remain text without creating thousands of duplicate number styles.
    public string NumberFormat{set{if(value!="@")throw new NotSupportedException("不支持的数字格式："+value);}}
    public void ClearContents()=>Write(null,false);
    public IEnumerator<ExcelRange> GetEnumerator(){for(int r=Row;r<=LastRow;r++)for(int c=Column;c<=LastColumn;c++)yield return new(Sheet,r,c,r,c);}
    IEnumerator IEnumerable.GetEnumerator()=>GetEnumerator();
}
public sealed class ExcelEnd
{
    readonly ExcelRange range;public ExcelEnd(ExcelRange range)=>this.range=range;
    public ExcelRange this[int direction]
    {get{if(direction==-4159){int column=range.Column;while(column>1&&range.Sheet.Book.Value(range.Cell(range.Row,column))==null)column--;return new(range.Sheet,range.Row,column,range.Row,column);}if(direction!=-4162)throw new NotSupportedException("不支持该查找方向。");int row=Math.Min(range.Row,range.Sheet.Data.LastRowNum+1);while(row>1&&range.Sheet.Book.Value(range.Cell(row,range.Column))==null)row--;return new(range.Sheet,row,range.Column,row,range.Column);}}
}
public sealed class ExcelApplication
{
    readonly ExcelBook book;public ExcelFunctions WorksheetFunction{get;}
    public ExcelApplication(ExcelBook book){this.book=book;WorksheetFunction=new();}
    public bool ScreenUpdating{get;set;}=true;public object Calculation{get;set;}=-4105;public object StatusBar{get;set;}=false;
    public ExcelRowCount Rows{get;}=new();
    public void CalculateFull()=>book.Evaluator.ClearAllCachedResultValues();
}
public sealed class ExcelRowCount{public int Count=>65536;}
public sealed class ExcelColumnCount{public int Count=>256;}
public sealed class ExcelFunctions
{
    public int CountA(ExcelRange range)=>range.Count(cell=>cell.Value2!=null);
    public object VLookup(object key,ExcelRange range,object column,bool approximate)=>Lookup(key,range,Convert.ToInt32(column),approximate,false);
    public object HLookup(object key,ExcelRange range,object row,bool approximate)=>Lookup(key,range,Convert.ToInt32(row),approximate,true);
    static object Lookup(object key,ExcelRange range,int offset,bool approximate,bool horizontal)
    {
        if(approximate)throw new NotSupportedException("仅支持精确数据库匹配。");
        if(offset<1||offset>(horizontal?range.LastRow-range.Row+1:range.LastColumn-range.Column+1))throw new ArgumentOutOfRangeException(nameof(offset));
        for(int i=horizontal?range.Column:range.Row;i<=(horizontal?range.LastColumn:range.LastRow);i++)
        {
            var cell=new ExcelRange(range.Sheet,horizontal?range.Row:i,horizontal?i:range.Column,horizontal?range.Row:i,horizontal?i:range.Column);
            object value=cell.Value2;
            bool match=key is string&&value is string?string.Equals((string)key,(string)value,StringComparison.OrdinalIgnoreCase):Equals(key,value);
            if(match)return new ExcelRange(range.Sheet,horizontal?range.Row+offset-1:i,horizontal?i:range.Column+offset-1,horizontal?range.Row+offset-1:i,horizontal?i:range.Column+offset-1).Value2??0d;
        }
        throw new InvalidOperationException("数据库中未找到："+key);
    }
}
