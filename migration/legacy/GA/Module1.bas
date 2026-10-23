Attribute VB_Name = "Module1"
Public Sub PlotFunction(PrinterName As String, Styles As String, MediaName As String, Copies As Integer, _
AutoMedia As Boolean, AutoRotate As Boolean, AutoClose As Boolean, AutoFrame As Boolean, AutoSave As Boolean)
'定义变量
Dim ptMin As Variant, ptMax As Variant
Dim p(0 To 2) As Double
Dim q(0 To 2) As Double
Dim r(0 To 2) As Double
Dim v(0 To 2) As Double
Dim w(0 To 2) As Double
Dim s(0 To 2) As Double
Dim t(0 To 2) As Double
Dim u(0 To 2) As Double
Dim x(0 To 2) As Double
Dim i, j, k, m As Integer
Dim objdoc, objold As AcadDocument
Dim mtext1, mtext2, mtext3, mtext4 As AcadMText
Dim savename, ss, shebeimingcheng As String
Dim yonghudingdanhao, xiangmumingcheng, yonghumingcheng, gongsimingcheng, Ver As String
Dim xlapp As Object
Dim xlbook As Object
Dim xlsheet, xlsheet1 As Object
Dim Count, Index, tempString, attrib, tagname
Dim BlockAttributes
Dim counter As Integer
Dim objLayout As AcadLayout
Dim objPlot As AcadPlot
Dim PathName As String
Dim nam, last As String
Dim jiekou, anchor As String
Dim pump, DrainZA As String
Dim obj As AcadEntity
Dim myset As AcadSelectionSet
Dim myobj As AcadSelectionSet
Dim mscale As AcadEntity
Dim exchanger, side, reservoir As String
Dim cool As String
Dim newpath As String
Dim typeface As String
Dim Bold As Boolean
Dim Italic As Boolean
Dim charset As Long
Dim pitchandfamily  As Long
Dim stz As AcadTextStyle
Dim Drain As String, newpumptype As String
Dim CName1, Cvalue1, CName2, Cvalue2, CName3, Cvalue3 As String
Dim I1 As Integer


On Error Resume Next

'取得当前文件路径
ss = ThisDrawing.Application.ActiveDocument.Path
newpath = VBA.Left(ss, Len(ss) - 9)
'创建文件夹


'取得EXCEL文件数据源
Set xlbook = GetObject(newpath & "\ZTools.xls")
Set xlsheet = xlbook.worksheets("sheet1")
Set xlsheet1 = xlbook.worksheets("input")

'判断CAD模板文件是否打开
nam = ThisDrawing.Application.ActiveDocument.Name
If nam <> ss & "\GA.dwg" Then
Set objdoc = ThisDrawing.Application.Documents.Open(ss & "\GA.dwg")
objdoc.Activate
Set objdoc = ThisDrawing.Application.ActiveDocument
End If
'隐藏Autocad窗口
ThisDrawing.Application.Visible = False
'对程序中变量赋值
i = Val(xlsheet.range(("C" & 19)))
savename = xlsheet.range(("P" & i))
shebeimingcheng = xlsheet.range(("D" & i))
yonghudingdanhao = xlsheet.range(("C" & 9))
xiangmumingcheng = xlsheet.range(("C" & 8))
yonghumingcheng = xlsheet.range(("C" & 7))
If xlsheet.range("CN" & i) <> "" Then
newpumptype = VBA.Replace(xlsheet.range("C" & i), VBA.Left(xlsheet.range("C" & i), 2), xlsheet.range("CN" & i))
Else
newpumptype = xlsheet.range("C" & i)
End If
j = Val(xlsheet.range(("AU" & i)))
jiekou = xlsheet.range(("BD" & i))
pump = VBA.Left(xlsheet.range(("C" & i)), 2)
anchor = VBA.Right(xlsheet.range(("AT" & i)), 7)
exchanger = xlsheet.range(("BE" & i))
reservoir = xlsheet.range(("BF" & i))
side = xlsheet.range(("BI" & i))
cool = xlsheet.range(("AL" & i))
gongsimingcheng = xlsheet1.range(("K" & 10))
Ver = VBA.Val(VBA.Left(xlsheet1.range(("E" & i - 19)), 1))
p(0) = 0
p(1) = 0
p(2) = 0
q(0) = 420
q(1) = 297
q(2) = 0
r(0) = 361
v(0) = 361
w(0) = 361
x(0) = 361
r(1) = 63
v(1) = 67
x(1) = 71.5
w(1) = 53.9
r(2) = 0
v(2) = 0
w(2) = 0
x(2) = 0
t(0) = 0
t(1) = -4.5
t(2) = 0




ptMin = p
ptMax = q

If pump = "ZA" Or pump = "CZ" Then
k = 7
m = 7
s(0) = 12.6394
s(1) = -4.4978
s(2) = 0
End If
If pump = "ZE" Then
 If Ver = 7 Then pump = "ZE7"
k = 8
m = 8
s(0) = 0
s(1) = 0
s(2) = 0
End If
If pump = "ZF" Then
m = 10
k = 8
s(0) = 0
s(1) = 0
s(2) = 0
End If
'插入图形及标注块
'Set stz = ThisDrawing.TextStyles.Add("Sulzer")
'stz.SetFont "Sulzer", False, False, 0, 0

'ThisDrawing.ActiveTextStyle.GetFont typeface, Bold, Italic, charset, pitchandfamily
'typeface = "Sulzer"
'MsgBox "123" & typeface
'ThisDrawing.ActiveTextStyle.SetFont typeface, Bold, Italic, charset, pitchandfamily
'mtext1.Update
'Set mtext1 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(r, 48, shebeimingcheng)



''Set Custom Property Value*********************************************************************************************************************************************


For I1 = 4 To 18
''CName1 = xlsheet.range("G" & I1)
''CName2 = xlsheet.range("I" & I1)
''CName3 = xlsheet.range("K" & I1)
''Cvalue1 = xlsheet.range("H" & I1)
''Cvalue2 = xlsheet.range("J" & I1)
''Cvalue3 = xlsheet.range("L" & I1)
changeFld xlsheet.range("G" & I1), xlsheet.range("H" & I1)
changeFld xlsheet.range("I" & I1), xlsheet.range("J" & I1)
changeFld xlsheet.range("K" & I1), xlsheet.range("L" & I1)

Next I1


''**************************************************************************************************************************************************************************
  
  

Set mtext1 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(r, 48, shebeimingcheng)
'ThisDrawing.ActiveTextStyle.Height = 2.2193
'ThisDrawing.ActiveTextStyle.Width = 48
Select Case LenB(shebeimingcheng)
Case Is < 18:
mtext1.ScaleEntity r, 1
mtext1.Height = 2.3
mtext1.Width = 48
mtext1.Update
Case 18 To 36:
mtext1.ScaleEntity r, 0.85
mtext1.Height = 2
mtext1.Width = 55
mtext1.Update
Case Is > 36:
mtext1.ScaleEntity r, 0.6
mtext1.Height = 1.65
mtext1.Width = 75
mtext1.Update
End Select


Set mtext2 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(v, 48, xiangmumingcheng)
'ThisDrawing.ActiveTextStyle.Height = 2.2193
'ThisDrawing.ActiveTextStyle.Width = 48
Select Case LenB(xiangmumingcheng)
Case Is < 18:
mtext2.ScaleEntity v, 1
mtext2.Height = 2.3
mtext2.Width = 48
mtext2.Update
Case 18 To 36:
mtext2.ScaleEntity v, 0.85
mtext2.Height = 2
mtext2.Width = 55
mtext2.Update
Case 37 To 54:
mtext2.ScaleEntity v, 0.6
mtext2.Height = 1.65
mtext2.Width = 75
mtext2.Update
Case 55 To 60:
mtext2.ScaleEntity v, 0.5
mtext2.Height = 1#
mtext2.Width = 75
mtext2.Update
Case Is > 60:
mtext2.ScaleEntity v, 0.4
mtext2.Height = 0.8
mtext2.Width = 65
mtext2.Update
End Select

Set mtext3 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(w, 48, yonghudingdanhao)
'ThisDrawing.ActiveTextStyle.Height = 2.2193
'ThisDrawing.ActiveTextStyle.Width = 48
Select Case LenB(yonghudingdanhao)
Case Is < 18:
mtext3.ScaleEntity w, 1
mtext3.Height = 2.3
mtext3.Width = 48
mtext3.Update
Case 18 To 36:
mtext3.ScaleEntity w, 0.85
mtext3.Height = 2
mtext3.Width = 55
mtext3.Update
Case Is > 36:
mtext3.ScaleEntity w, 0.6
mtext3.Height = 1.6
mtext3.Width = 70
mtext3.Update
End Select

Set mtext4 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(x, 48, yonghumingcheng)
'ThisDrawing.ActiveTextStyle.Height = 2.2193
'ThisDrawing.ActiveTextStyle.Width = 48
Select Case LenB(yonghumingcheng)
Case Is < 18:
mtext4.ScaleEntity x, 1
mtext4.Height = 2.3
mtext4.Width = 48
mtext4.Update
Case 18 To 36:
mtext4.ScaleEntity x, 0.85
mtext4.Height = 2
mtext4.Width = 55
mtext4.Update
Case Is > 36:
mtext4.ScaleEntity x, 0.6
mtext4.Height = 1.6
mtext4.Width = 70
mtext4.Update
End Select




If Val(xlsheet.range(("C" & 18))) = 1 Then
  Drain = xlsheet.range(("BR" & i))
  If pump <> "CZ" Then
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "Drain" & m, 1, 1, 1, 0
  End If
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, Drain, 1, 1, 1, 0
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "jibenjiekou", 1, 1, 1, 0
  t(1) = 0
  Else
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "jibenjiekouwod", 1, 1, 1, 0
  
    End If

If Val(xlsheet.range(("C" & 16))) = 1 Then
  ThisDrawing.Application.ActiveDocument.Layers.Item("matingflange").LayerOn = True
  Else
  ThisDrawing.Application.ActiveDocument.Layers.Item("matingflange").LayerOn = False
  End If


If exchanger = "Y" Then
 If side <> "L" Then
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "huanreqi" & m & "you", 1, 1, 1, 0
 Else
 If reservoir = "Y" Then
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "huanreqi" & m & "you", 1, 1, 1, 0
 Else
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "huanreqi" & k & "zuo", 1, 1, 1, 0
   End If
   End If
   End If
   
 If Val(xlsheet.range(("C" & 17))) = 1 Then
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "youwurunhua" & m, 1, 1, 1, 0
  End If
   
 
If reservoir = "Y" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "hongxiguan" & m, 1, 1, 1, 0
End If

If cool <> "NO" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "lengqueshui" & k, 1, 1, 1, 0
End If

If xlsheet.range(("AR" & i)) <> "NO" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "tuxingFI" & k, 1, 1, 1, 0
End If

If xlsheet.range(("AP" & i)) <> "NO" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "tuxingQ" & k, 1, 1, 1, 0
End If

If xlsheet.range(("AV" & i)) <> "NO" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "tuxingB" & k, 1, 1, 1, 0
End If

ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump, 1, 1, 1, 0
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "jibenchicun" & k, 1, 1, 1, 0
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock s, "geduan" & j, 1, 1, 1, 0
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "nozzle", 1, 1, 1, 0
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock t, jiekou, 1, 1, 1, 0
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, anchor, 1, 1, 1, 0

ThisDrawing.ActiveTextStyle.SetFont "Standard", False, False, 0, 0
'对所有块的属性赋值
Count = ThisDrawing.ActiveLayout.Block.Count
For Index = 0 To Count - 1
tempString = ThisDrawing.ActiveLayout.Block(Index).ObjectName
If tempString = "AcDbBlockReference" Then
If tempString.Name = "AAA" Then
BlockAttributes = ThisDrawing.ActiveLayout.Block(Index).GetAttributes
For Each attrib In BlockAttributes
tagname = attrib.TagString
Select Case VBA.UCase(tagname)
'Case "YONGHUMINGCHENG":
'Set attribute value
'attrib.TextString = xlsheet.range(("C" & 7))
'xlsheet.range (("AS" & i))
'Case "XIANGMUMINGCHENG":
'Set attribute value
'attrib.TextString = xlsheet.range(("C" & 8))
'Case "SHEBEIMINGCHENG":
'Set attribute value
'attrib.TextString = xlsheet.range(("D" & i))
Case "SHEBEIWEIHAO":
'Set attribute value
attrib.TextString = xlsheet.range(("B" & i))
'Case "YONGHUDINGDANHAO":
'Set attribute value
'attrib.TextString = xlsheet.range(("C" & 9))
Case "DINGDANHAO_SULZER":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 10))
Case "GUANLUFANGAN":
'Set attribute value
attrib.TextString = xlsheet.range(("E" & i))


Case "TYPE_PUMP":
'Set attribute value
attrib.TextString = " " & newpumptype
Case "DATE":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 6))
Case "DESIGN":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 4))
Case "CHECK":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 5))
Case "PUMP":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("F" & i)))
Case "MOTOR":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("G" & i)))
Case "BASE":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("H" & i)))
Case "ACCESSORY":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("I" & i)))
Case "TOTAL":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("J" & i)))
Case "DIANJIXINGHAO":
'Set attribute value
attrib.TextString = xlsheet.range(("K" & i))
Case "DIANYA":
'Set attribute value
attrib.TextString = xlsheet.range(("L" & i))
Case "EDINGGONGLV":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("M" & i)))
Case "DIANJICHANGJIA":
'Set attribute value
attrib.TextString = xlsheet.range(("N" & i))

Case "ZHUANSU":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("O" & i)))
Case "COMPANY":
'Set attribute value  gongsimingcheng
attrib.TextString = gongsimingcheng

Case "SCI":
'Set attribute value
attrib.TextString = xlsheet.range(("AL" & i))
Case "SCO":
'Set attribute value
attrib.TextString = xlsheet.range(("AL" & i))
Case "FCI":
'Set attribute value
attrib.TextString = xlsheet.range(("AO" & i))
Case "FCO":
'Set attribute value
attrib.TextString = xlsheet.range(("AO" & i))
Case "SQI":
'Set attribute value
attrib.TextString = xlsheet.range(("AP" & i))
Case "SQO":
'Set attribute value
attrib.TextString = xlsheet.range(("AP" & i))
Case "FQI":
'Set attribute value
attrib.TextString = xlsheet.range(("AQ" & i))
Case "FQO":
'Set attribute value
attrib.TextString = xlsheet.range(("AQ" & i))
Case "SFI":
'Set attribute value
attrib.TextString = xlsheet.range(("AR" & i))
Case "FFI":
'Set attribute value
attrib.TextString = xlsheet.range(("AS" & i))
Case "SBI":
'Set attribute value
attrib.TextString = xlsheet.range(("AV" & i))
Case "SBO":
'Set attribute value
attrib.TextString = xlsheet.range(("AV" & i))
Case "FBI":
'Set attribute value
attrib.TextString = xlsheet.range(("AW" & i))
Case "FBO":
'Set attribute value
attrib.TextString = xlsheet.range(("AW" & i))
Case "SD":
'Set attribute value
attrib.TextString = xlsheet.range(("S" & i))
Case "SS":
'Set attribute value
attrib.TextString = xlsheet.range(("T" & i))
Case "FSD":
'Set attribute value
attrib.TextString = xlsheet.range(("U" & i))
Case "SPD":
'Set attribute value
attrib.TextString = xlsheet.range(("V" & i))
Case "FPD":
'Set attribute value
attrib.TextString = xlsheet.range(("W" & i))
Case "DFX":
'Set attribute value
attrib.TextString = xlsheet.range(("AX" & i))
Case "DFY":
'Set attribute value
attrib.TextString = xlsheet.range(("AY" & i))
Case "DFZ":
'Set attribute value
attrib.TextString = xlsheet.range(("AZ" & i))
Case "SFX":
'Set attribute value
attrib.TextString = xlsheet.range(("BA" & i))
Case "SFY":
'Set attribute value
attrib.TextString = xlsheet.range(("BB" & i))
Case "SFZ":
'Set attribute value
attrib.TextString = xlsheet.range(("BC" & i))
Case "DMX":
'Set attribute value
attrib.TextString = xlsheet.range(("BK" & i))
Case "DMY":
'Set attribute value
attrib.TextString = xlsheet.range(("BL" & i))
Case "DMZ":
'Set attribute value
attrib.TextString = xlsheet.range(("BM" & i))
Case "SMX":
'Set attribute value
attrib.TextString = xlsheet.range(("BN" & i))
Case "SMY":
'Set attribute value
attrib.TextString = xlsheet.range(("BO" & i))
Case "SMZ":
'Set attribute value
attrib.TextString = xlsheet.range(("BP" & i))
Case "SAD2":
'Set attribute value
attrib.TextString = xlsheet.range(("BJ" & i))
Case "FAD2":
'Set attribute value
attrib.TextString = "NPTF"

Case "LF":
'Set attribute value
attrib.TextString = xlsheet.range(("BT" & i))
Case "LB":
'Set attribute value
attrib.TextString = xlsheet.range(("BV" & i))
Case "LQ":
'Set attribute value
attrib.TextString = xlsheet.range(("BX" & i))
Case "HF":
'Set attribute value
attrib.TextString = xlsheet.range(("BS" & i))
Case "HB":
'Set attribute value
attrib.TextString = xlsheet.range(("BU" & i))
Case "HQ":
'Set attribute value
attrib.TextString = xlsheet.range(("BW" & i))
Case "LEX":
'Set attribute value
attrib.TextString = xlsheet.range(("BG" & i))
Case "LRE":
'Set attribute value
attrib.TextString = xlsheet.range(("BH" & i))
Case "NAME":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 4))

Case "E":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AJ" & i)))
'xlsheet.range (("AS" & i))
Case "F":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AK" & i)))
Case "Y":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AM" & i)))
Case "K":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AC" & i)))
Case "H_DRAIN":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AN" & i)))
Case "H2":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AI" & i)))
Case "L3":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AB" & i)))
Case "L4":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("BQ" & i)))
Case "L":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("Y" & i)))
Case "B1":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AE" & i)))
Case "B":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AD" & i)))
Case "I":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AF" & i)))
Case "H_BASE":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AG" & i)))
Case "H_PUMP":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("X" & i)))
Case "H_ANCHOR":
'Set attribute value
attrib.TextString = VBA.Right(xlsheet.range(("AT" & i)), 3)
Case "BOLTS":
'Set attribute value
attrib.TextString = xlsheet.range(("AT" & i))

Case "L2":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("AA" & i)))
Case "L1":
'Set attribute value
attrib.TextString = Val(xlsheet.range(("Z" & i)))
'xlsheet.range (("AS" & i))
End Select
Next attrib
End If
End If
Next Index

ThisDrawing.Application.Visible = False
  
  

  
'设置PDF打印参数
Set objdoc = ThisDrawing.Application.ActiveDocument
Set objLayout = objdoc.Layouts.Item("Model")
Set objPlot = objdoc.Plot
ThisDrawing.Application.ZoomExtents

' 设置打印机
If Not VBA.Trim(PrinterName) = "" Then
objLayout.ConfigName = PrinterName
Else
Exit Sub
End If

' 设置打印样式表
If Not VBA.Trim(Styles) = "" Then
objLayout.StyleSheet = Styles
Else
objLayout.StyleSheet = "acad.ctb"
End If

' 设置图纸尺寸
If AutoMedia Then
objLayout.CanonicalMediaName = "A3"
Else
If Not VBA.Trim(MediaName) = "" Then
objLayout.CanonicalMediaName = MediaName
Else
objLayout.CanonicalMediaName = "A3"
End If
End If

' 设置图纸单位
objLayout.PaperUnits = acMillimeters
'objLayout.PaperUnits = acInches

' 设置默认图纸打印方向
'objLayout.PlotRotation = ac0degrees     '纵向
'objLayout.PlotRotation = ac180degrees
objLayout.PlotRotation = ac90degrees   '横向
'objLayout.PlotRotation = ac270degrees

' 设置图纸打印比例
objLayout.StandardScale = acScaleToFit
'objLayout.UseStandardScale = False '使用自定义打印比例

' 设置自定义打印比例
'objLayout.SetCustomScale txtNumerator.Value, txtDenominator.Value

' 设置图纸是否居中打印
objLayout.CenterPlot = True

' 打印时使用图形文件中的线宽
objLayout.PlotWithLineweights = True

' 设置是否应用打印样式
objLayout.PlotWithPlotStyles = True

' 打印时隐藏图纸空间对象
'objLayout.PlotHidden = False

' 设置图纸打印份数
If Copies >= 1 Then
objPlot.NumberOfCopies = CInt(Copies)
Else
objPlot.NumberOfCopies = 1
End If

' 将打印错误报告切换为静默错误模式，以便不间断地执行打印任务
objPlot.QuietErrorMode = True
ThisDrawing.PurgeAll
' 重新生成当前图形
objdoc.Regen acAllViewports

' 设置前台打印，使打印任务按打印顺序依次发送到打印机
objdoc.SetVariable "BACKGROUNDPLOT", 0
ReDim Preserve ptMin(0 To 1)
ReDim Preserve ptMax(0 To 1)
ThisDrawing.ActiveLayout.SetWindowToPlot ptMin, ptMax
objLayout.PlotType = acWindow
objPlot.PlotToFile newpath & "\GA\" & savename & ".pdf"
'另存为dwg格式文件
objdoc.SaveAs newpath & "\GA\" & savename & ".dwg"

'关闭另存后的活动文件
ThisDrawing.Application.ActiveDocument.Close

'xlbook.Application.Close
'Set xlbook = Nothing


End Sub
Private Function custFldAry() As String
'makes a comma delimted string of existing
'custom properties and values
On Error GoTo errCaseH
Dim docInfo As AcadSummaryInfo
Set docInfo = ThisDrawing.SummaryInfo
Dim fName As String
Dim fVal As String
Dim i As Integer
For i = 0 To (docInfo.NumCustomInfo - 1)
docInfo.GetCustomByIndex i, fName, fVal
fName = VBA.Trim(fName)
fVal = "," & VBA.Trim(fVal)
custFldAry = custFldAry & "," & VBA.Trim(fName) & fVal
Next i
If VBA.Trim(Replace(custFldAry, ",", "")) = "" Then custFldAry = ""
Set docInfo = Nothing
Exit Function
errCaseH:
Set docInfo = Nothing
custFldAry = ""
End Function

Private Function mkFldCode(inputStr) As String
'makes a field code for a custom property
mkFldCode = "%<\AcVar CustomDP." & inputStr & ">%"
End Function

Private Function changeFld(fldName As String, newVal As String) As Boolean
'changes a custom property value, or creates custom prop
'and assigns a value
On Error GoTo errCaseG
changeFld = True
Dim custProps As AcadSummaryInfo
Set custProps = ThisDrawing.SummaryInfo
If VBA.InStr(custFldAry, fldName) < 1 Then GoTo createIt
custProps.SetCustomByKey fldName, newVal
Set custProps = Nothing
Exit Function
createIt:
custProps.AddCustomInfo fldName, newVal
Set custProps = Nothing
Exit Function
errCaseG:
changeFld = False
Set custProps = Nothing
End Function



Public Sub GA()
Call PlotFunction("DWG to PDF.pc3", "acad.ctb", "A3", 1, True, True, False, True, True)
End Sub


