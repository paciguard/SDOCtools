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
Dim yonghudingdanhao, xiangmumingcheng, yonghumingcheng As String
Dim xlapp As Object
Dim xlbook As Object
Dim xlsheet, xlsheet1, xlsheet2 As Object
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
Dim th As String
Dim bearing, LK, newpumptype, Ver As String
Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15 As String
Dim lp1, lp2, lp3, lp4, lp5, lp6 As String
Dim NAMlast As String

On Error Resume Next

'取得当前文件路径
ss = ThisDrawing.Application.ActiveDocument.Path
newpath = VBA.Left(ss, Len(ss) - 9)
'创建文件夹


'取得EXCEL文件数据源
Set xlbook = GetObject(newpath & "\ZTools.xls")
Set xlsheet = xlbook.worksheets("sheet1")
Set xlsheet1 = xlbook.worksheets("PID")
Set xlsheet2 = xlbook.worksheets("input")
'判断CAD模板文件是否打开
nam = ThisDrawing.Application.ActiveDocument.Name
If nam <> "PID.dwg" Then
Set objdoc = ThisDrawing.Application.Documents.Open(ss & "\PID.dwg")

objdoc.Activate
Set objdoc = ThisDrawing.Application.ActiveDocument
End If
'隐藏Autocad窗口
'ThisDrawing.Application.Visible = False
'对程序中变量赋值
i = Val(xlsheet.range(("C" & 19)))
savename = xlsheet.range(("Q" & i))
Ver = VBA.Val(VBA.Left(xlsheet2.range(("E" & i - 19)), 1))
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
If Ver = 7 And pump = "ZE" Then pump = "ZE7"

anchor = VBA.Right(xlsheet.range(("AT" & i)), 7)
p1 = xlsheet.range(("CH" & i))
p2 = xlsheet.range(("CI" & i))
p3 = xlsheet.range(("CJ" & i))
p4 = xlsheet.range(("CK" & i))
p5 = xlsheet.range(("CL" & i))
p6 = xlsheet.range(("CM" & i))
lp1 = VBA.Left(p1, 2)
lp2 = VBA.Left(p2, 2)
lp3 = VBA.Left(p3, 2)
lp4 = VBA.Left(p4, 2)
lp5 = VBA.Left(p5, 2)
lp6 = VBA.Left(p6, 2)

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
u(0) = 3.5
u(1) = 0
u(2) = 0





'插入图形及标注块
'Set stz = ThisDrawing.TextStyles.Add("Sulzer")
'stz.SetFont "Sulzer", False, False, 0, 0

'ThisDrawing.ActiveTextStyle.GetFont typeface, Bold, Italic, charset, pitchandfamily
'typeface = "Sulzer"
'MsgBox "123" & typeface
'ThisDrawing.ActiveTextStyle.SetFont typeface, Bold, Italic, charset, pitchandfamily
'mtext1.Update
'Set mtext1 = ThisDrawing.Application.ActiveDocument.ModelSpace.AddMText(r, 48, shebeimingcheng)

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
Case Is > 36:
mtext2.ScaleEntity v, 0.6
mtext2.Height = 1.65
mtext2.Width = 75
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


'判断泵型
ptMin = p
ptMax = q

If pump = "ZF" Then
pump = "ZE"
End If
 '判断52是否带冷却水
 If p1 = "52" Or p2 = "52" Or p3 = "52" Then
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 p7 = "52WM"
 Else
 p7 = "52WOM"
 End If
 Else
 p7 = ""
 End If
 '判断53A是否带冷却水
 If p1 = "53A" Or p2 = "53A" Or p3 = "53A" Then
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 p8 = "53AWM"
 Else
 p8 = "53AWOM"
 End If
 Else
 p8 = ""
 End If
'判断53B是否带冷却水
 If p1 = "53B" Or p2 = "53B" Or p3 = "53B" Then
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 p14 = "53BWM"
 Else
 p14 = "53BWOM"
 End If
 Else
 p14 = ""
 End If
'判断53C是否带冷却水,并是否存在干涉
If p1 = "53C" Or p2 = "53C" Or p3 = "53C" Or p4 = "53C" Then
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 If p1 = "11" Or p1 = "21" Or p1 = "23" Or p1 = "32" Or p2 = "62" Or p3 = "62" Then
  p9 = "53CWM+"
  Else
  p9 = "53CSWM"
 End If
 
 Else
 If p1 = "11" Or p1 = "21" Or p1 = "23" Or p1 = "32" Or p2 = "62" Or p3 = "62" Then
  p9 = "53CWOM+"
  Else
  p9 = "53CSWOM"
 End If
 End If
 Else
 p9 = ""
 End If
 
 
 
 
 '判断是否有planC并是否有干涉
 If p2 = "C" Or p3 = "C" Or p4 = "C" Then
 If p1 = "21" Or p1 = "11" Or p1 = "23" Then
 Select Case VBA.UCase(p1)
 Case "21":
 p10 = "C+"
 Case "11":
 p10 = "C+"
 Case "23"
 p10 = "C++"
 End Select
 Else
 p10 = "CS"
 End If
 Else
 p10 = ""
 End If
 
 
 
 '判断是否有planS并是否有干涉
 
 If p2 = "S" Or p3 = "S" Or p4 = "S" Or p5 = "S" Or p6 = "S" Then
 If p1 = "11" Or p1 = "21" Or p1 = "23" Then
 Select Case VBA.UCase(p1)
 Case "21":
 p11 = "S+"
 Case "11":
 p11 = "S+"
 Case "23"
 p11 = "S++"
 End Select
 Else
 p11 = "SS"
 End If
 Else
 p11 = ""
 End If
 
 '判断是否有planA
  If p2 = "A" Or p3 = "A" Or p4 = "A" Or p5 = "A" Then
 p12 = "A"
 Else
 p12 = ""
 End If
 '如果有planC/S/A,即存在冷却方案
 If p2 <> "M" And p3 <> "M" And p4 <> "M" And p5 <> "M" And p6 <> "M" Then
 If p10 = "C+" Or p10 = "C++" Or p10 = "CS" Or p11 = "S+" Or p11 = "SS" Or p11 = "S++" Or p12 = "A" Then
 p13 = "M"
 Else
 p13 = ""
 End If
  End If
 '判断72位置是否干涉
 If p1 = "72" Or p2 = "72" Or p3 = "72" Or p4 = "72" Or p5 = "72" Or p6 = "72" Then
 If p2 = "76" Or p3 = "76" Or p4 = "76" Or p5 = "76" Or p6 = "76" Then
 p15 = "72DS"
 Else
 p15 = "72US"
 End If
 Else
 p15 = ""
 End If
 '判断冷却水分支数量,相应增减球阀
 If p1 = "32" Or p2 = "32" Or p3 = "32" Or p1 = "62" Or p2 = "62" Or p3 = "62" Or p4 = "62" Then
 If (VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "AC" And VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 3) <> "3AC") Or (VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "AS" And VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 3) <> "3AS") Or VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "CS" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = False
 ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = True
 End If
 End If
  
  
 
  
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 If p1 = "21" Or p1 = "23" Or p1 = "41" Or p1 = "32" Or p2 = "32" Or p1 = "62" Or p2 = "62" Or p3 = "62" Then
 If p2 = "A" Or p3 = "A" Or p4 = "A" Or p5 = "A" Or p2 = "C" Or p3 = "C" Or p4 = "C" Or p5 = "C" Or p6 = "C" Or p2 = "S" Or p3 = "S" Or p4 = "S" Or p5 = "S" Or p6 = "S" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = False
  ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = True
  End If
 End If
 End If
   
If p1 = "52" Or p2 = "52" Or p3 = "52" Or p4 = "52" Or p5 = "52" Or lp1 = "53" Or lp2 = "53" Or lp3 = "53" Or lp4 = "53" Or lp5 = "53" Or p1 = "54" Or p2 = "54" Or p3 = "54" Or p1 = "72" Or p2 = "72" Or p3 = "72" Or p4 = "72" Or p1 = "74" Or p2 = "74" Or p3 = "74" Or p4 = "74" Or p2 = "75" Or p3 = "75" Or p4 = "75" Or p5 = "75" Or p2 = "76" Or p3 = "76" Or p4 = "76" Or p5 = "76" Then
 If p2 = "M" Or p3 = "M" Or p4 = "M" Or p5 = "M" Or p6 = "M" Then
 If p1 = "21" Or p1 = "23" Or p1 = "41" Or p2 = "A" Or p3 = "A" Or p4 = "A" Or p5 = "A" Or p2 = "C" Or p3 = "C" Or p4 = "C" Or p5 = "C" Or p6 = "C" Or p2 = "S" Or p3 = "S" Or p4 = "S" Or p5 = "S" Or p6 = "S" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = False
 
 Else
  ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = False
  ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = True
  End If
  Else
  If (VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "AC" And VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 3) <> "3AC") Or (VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "AS" And VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 3) <> "3AS") Or VBA.Right(Replace(xlsheet.range(("CG" & i)), "+", ""), 2) = "CS" Then
  ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = False
 
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("WB").LayerOn = False
  ThisDrawing.Application.ActiveDocument.Layers.Item("WOB").LayerOn = True
  End If
  End If
  End If
 
 '判明53A/53B/53C与planA干涉情况
 If p7 <> "" Or p8 <> "" Or p14 = "53BWM" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO1").LayerOn = True
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO1").LayerOn = False
 End If
 
 If p9 = "53CWM+" Or p9 = "53CWOM+" Or p14 = "53BWOM" Or p7 = "52WM" Or p8 = "53AWM" Or p14 = "53BWM" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO2").LayerOn = True
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO2").LayerOn = False
 End If
 
 If p9 = "53CWM+" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO3").LayerOn = True
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AO3").LayerOn = False
 End If
 
 If p7 <> "" Or p8 <> "" Or p14 = "53BWM" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF1").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF1").LayerOn = True
 End If
 
 If p9 = "53CWM+" Or p9 = "53CWOM+" Or p14 = "53BWOM" Or p7 = "52WM" Or p8 = "53AWM" Or p14 = "53BWM" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF2").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF2").LayerOn = True
 End If
 
 If p9 = "53CWM+" Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF3").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("AF3").LayerOn = True
 End If
 
 If Val(xlsheet.range(("C" & 18))) = 1 Then
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "PD", 1, 1, 1, 0
 End If
 
  
 '判断与泵排残管线干涉
 If p1 = "21" Or p1 = "23" Or p1 = "41" Or p2 = "41" Or p11 <> "" Then
 If Val(xlsheet.range(("C" & 18))) = 1 Then
 ThisDrawing.Application.ActiveDocument.Layers.Item("PDO").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("PDF").LayerOn = False
 Else
 ThisDrawing.Application.ActiveDocument.Layers.Item("PDF").LayerOn = True
 ThisDrawing.Application.ActiveDocument.Layers.Item("PDO").LayerOn = False
 End If
 End If
 
 '插入所有冲洗、冷却管线及图形块
 'MsgBox ("Pump=" & pump)
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump, 1, 1, 1, 0
 If p1 = "32" And p2 = "53B" Then
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock u, p1, 1, 1, 1, 0
 Else
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p1, 1, 1, 1, 0
 End If
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p2, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p3, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p4, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p5, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p6, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p7, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p8, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p9, 1, 1, 1, 0
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p10, 1, 1, 1, 0
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p11, 1, 1, 1, 0
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p12, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p13, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p14, 1, 1, 1, 0
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, p15, 1, 1, 1, 0
 
 
 '判断接口并插入接口列表块
 If Val(xlsheet.range(("C" & 18))) = 1 Then
  
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "jibenjiekou", 1, 1, 1, 0
  t(1) = 0
  Else
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "jibenjiekouwod", 1, 1, 1, 0
  
    End If
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock t, jiekou, 1, 1, 1, 0
 
 
 If xlsheet1.range(("G" & (i - 30))) <> "" And xlsheet1.range(("H" & (i - 30))) <> "" Then
 ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "BVBD", 1, 1, 1, 0
 End If
 
 
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
attrib.TextString = "                       " & newpumptype
Case "DATE":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 6))
Case "DESIGN":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 4))
Case "CHECK":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 5))

Case "NAME":
'Set attribute value
attrib.TextString = xlsheet.range(("C" & 4))


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

Case "SBV":
'Set attribute value
attrib.TextString = xlsheet1.range(("G" & (i - 30)))
Case "SBD":
'Set attribute value
attrib.TextString = xlsheet1.range(("H" & (i - 30)))
Case "FBV":
'Set attribute value
attrib.TextString = xlsheet1.range(("I" & (i - 30)))
Case "FBD":
'Set attribute value
attrib.TextString = xlsheet1.range(("I" & (i - 30)))





End Select
Next attrib
End If
End If
Next Index



'ThisDrawing.Application.Visible = False
  
  
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

' 重新生成当前图形
objdoc.Regen acAllViewports
' 设置前台打印，使打印任务按打印顺序依次发送到打印机
objdoc.SetVariable "BACKGROUNDPLOT", 0
ReDim Preserve ptMin(0 To 1)
ReDim Preserve ptMax(0 To 1)
ThisDrawing.ActiveLayout.SetWindowToPlot ptMin, ptMax
objLayout.PlotType = acWindow
objPlot.PlotToFile newpath & "\PID\" & savename & ".pdf"
'另存为dwg格式文件
''objdoc.SaveAs newpath & "\PID\" & savename & ".dwg"
For Each myset In ThisDrawing.SelectionSets
If myset.Name = "aaa" Then
ThisDrawing.SelectionSets.Item("aaa").Delete
Exit For
End If
Next
Set myset = ThisDrawing.SelectionSets.Add("aaa")
myset.Select acSelectionSetWindow, p, q
ThisDrawing.PurgeAll
ThisDrawing.Wblock newpath & "\PID\" & savename & ".dwg", myset '注意，块的保存路径为：D:。文件名为：wbolck.dwg
myset.Delete
Set myset = Nothing
objdoc.Activate

'关闭另存后的活动文件
''ThisDrawing.Application.ActiveDocument.Close

ThisDrawing.Application.ActiveDocument.Close False
NAMlast = ThisDrawing.Application.ActiveDocument.Name
If NAMlast = "" Then
AcadApplication.Documents.Open ss & "\PID.dwg"
End If

'xlbook.Application.Close
'Set xlbook = Nothing


End Sub


Public Sub PID()
Call PlotFunction("DWG to PDF.pc3", "acad.ctb", "A3", 1, True, True, False, True, True)
End Sub








