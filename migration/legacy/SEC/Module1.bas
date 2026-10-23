Attribute VB_Name = "Module1"
Public Sub PlotFunction(PrinterName As String, Styles As String, MediaName As String, Copies As Integer, _
AutoMedia As Boolean, AutoRotate As Boolean, AutoClose As Boolean, AutoFrame As Boolean, AutoSave As Boolean)
'¨??±???
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
Dim yonghudingdanhao, xiangmumingcheng, yonghumingcheng, gongsimingcheng As String
Dim xlapp As Object
Dim xlbook As Object
Dim xlsheet, xlsheet1, xlsheet2, xlsheet3 As Object
Dim Count, Index, tempString, Attrib, Tagname
Dim BlockAttributes
Dim counter As Integer
Dim objLayout As AcadLayout
Dim objPlot As AcadPlot
Dim PathName As String
Dim nam, last As String
Dim jiekou, anchor As String
Dim size, pump, DrainZA, NorDia As String
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
Dim th As String, newpumptype As String
Dim ver As String
Dim bearing, LK As String
Dim n As Integer



On Error Resume Next

'?????±?°?????·?
ss = ThisDrawing.Application.ActiveDocument.Path
newpath = VBA.Left(ss, Len(ss) - 9)
'???¨??????


'????EXCEL??????????
Set xlbook = GetObject(newpath & "\ZTools.xls")
Set xlsheet2 = xlbook.worksheets("input")
Set xlsheet = xlbook.worksheets("sheet1")
Set xlsheet1 = xlbook.worksheets("Sec-List")
Set xlsheet3 = xlbook.worksheets("CZ-BearingSize")
'???CAD??°???????·??ò??
nam = ThisDrawing.Application.ActiveDocument.Name
If nam <> ss & "\SEC.dwg" Then
Set objdoc = ThisDrawing.Application.Documents.Open(ss & "\SEC.dwg")
objdoc.Activate
Set objdoc = ThisDrawing.Application.ActiveDocument
End If
'????Autocad?°??
'ThisDrawing.Application.Visible = False
'????ò??±???????
i = Val(xlsheet.range(("C" & 19)))
savename = xlsheet.range(("R" & i))
shebeimingcheng = xlsheet.range(("D" & i))
yonghudingdanhao = xlsheet.range(("C" & 9))
xiangmumingcheng = xlsheet.range(("C" & 8))
yonghumingcheng = xlsheet.range(("C" & 7))
gongsimingcheng = xlsheet2.range(("K" & 10))
If xlsheet.range("CN" & i) <> "" Then
newpumptype = VBA.Replace(xlsheet.range("C" & i), VBA.Left(xlsheet.range("C" & i), 2), xlsheet.range("CN" & i))
Else
newpumptype = xlsheet.range("C" & i)
End If
j = Val(xlsheet.range(("AU" & i)))
jiekou = xlsheet.range(("BD" & i))
size = xlsheet.range("C" & i)
pump = VBA.Left(xlsheet.range(("C" & i)), 2)
NorDia = VBA.Right(xlsheet.range(("C" & i)), 3)
anchor = VBA.Right(xlsheet.range(("AT" & i)), 7)
exchanger = xlsheet.range(("BE" & i))
reservoir = xlsheet.range(("BF" & i))
side = xlsheet.range(("BI" & i))
cool = xlsheet.range(("AL" & i))
th = VBA.Right(VBA.Left(xlsheet.range(("CA" & i)), 3), 1)
bearing = xlsheet.range(("CB" & i))
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



ver = th

ptMin = p
ptMax = q
If th = "A" Or pump = "ZF" Then
ver = ""
Else
If th = "1" And pump = "ZE" Then
ver = 10
End If
End If

If pump = "CZ" Then
LK = VBA.Right(VBA.Left(xlsheet.range("CD" & i), 3), 1)
Else
If xlsheet.range(("BZ" & i)) = "Y" Then
LK = "_" & "K6"
Else
LK = ""
End If
End If



'?????????°±ê×??é
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


If Val(xlsheet.range(("C" & 18))) = 1 Then
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "Drain" & pump, 1, 1, 1, 0
   End If

If pump <> "ZA" And pump <> "XA" And pump <> "CZ" Then
  If Val(xlsheet.range(("C" & 17))) = 1 Then
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "PUREOIL", 1, 1, 1, 0
  End If
  
End If

'MsgBox ("LK=" & LK & "/// Size=" & size)
If pump = "CZ" Then

If VBA.Val(LK) >= 6 Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & "W" & "BALHL", 1, 1, 1, 0

Else

If Val(NorDia) >= 315 Then

If bearing = "CZ重轴" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & "W" & "BALHD", 1, 1, 1, 0
ElseIf bearing = "CZ轻轴" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & "W" & "BALHS", 1, 1, 1, 0
End If

Else

If bearing = "CZ重轴" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & "WO" & "BALHD", 1, 1, 1, 0
ElseIf bearing = "CZ轻轴" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & "WO" & "BALHS", 1, 1, 1, 0
End If

End If

End If

End If




If pump <> "ZA" And xlsheet.range(("CF" & i)) <> "Y" Then
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & ver & "_" & bearing & LK & "WO", 1, 1, 1, 0
 
 Else
ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, pump & ver & "_" & bearing & LK, 1, 1, 1, 0
End If
If pump = "ZE" And xlsheet.range(("BY" & i)) = "Y" Then
  ThisDrawing.Application.ActiveDocument.ModelSpace.InsertBlock p, "ZEG", 1, 1, 1, 0
End If

ThisDrawing.ActiveTextStyle.SetFont "Standard", False, False, 0, 0
'??ù???é??????????
''Count = ThisDrawing.ActiveLayout.Block.Count
Count = ThisDrawing.Layouts.Item("Model").Block.Count
For Index = 0 To Count - 1
''tempString = ThisDrawing.ActiveLayout.Block(Index).ObjectName
tempString = ThisDrawing.Layouts.Item("Model").Block(Index).ObjectName
If tempString = "AcDbBlockReference" Then
If tempString.Name = "AAA" Then
''BlockAttributes = ThisDrawing.ActiveLayout.Block(Index).GetAttributes
BlockAttributes = ThisDrawing.Layouts.Item("Model").Block(Index).GetAttributes
For Each Attrib In BlockAttributes
Tagname = Attrib.TagString
Select Case VBA.UCase(Tagname)
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
Attrib.TextString = xlsheet.range(("B" & i))
'Case "YONGHUDINGDANHAO":
'Set attribute value
'attrib.TextString = xlsheet.range(("C" & 9))
Case "DINGDANHAO_SULZER":
'Set attribute value
Attrib.TextString = xlsheet.range(("C" & 10))
Case "CAILIAODAIMA":
'Set attribute value
Attrib.TextString = xlsheet1.range(("F" & (i - 30)))


Case "TYPE_PUMP":
'Set attribute value

Attrib.TextString = "   " & newpumptype

Case "DATE":
'Set attribute value
Attrib.TextString = xlsheet.range(("C" & 6))
Case "DESIGN":
'Set attribute value
Attrib.TextString = xlsheet.range(("C" & 4))
Case "CHECK":
'Set attribute value
Attrib.TextString = xlsheet.range(("C" & 5))
Case "COMPANY":
'Set attribute value  gongsimingcheng
Attrib.TextString = gongsimingcheng

End Select
Next Attrib
End If
End If
Next Index

'??????????????é????
For n = 1 To 40
''Count = ThisDrawing.ActiveLayout.Block.Count
Count = ThisDrawing.Layouts.Item("Model").Block.Count

For Index = 0 To Count - 1
''tempString = ThisDrawing.ActiveLayout.Block(Index).ObjectName
tempString = ThisDrawing.Layouts.Item("Model").Block(Index).ObjectName
If tempString = "AcDbBlockReference" Then
If tempString.Name = "BBB" Then
''BlockAttributes = ThisDrawing.ActiveLayout.Block(Index).GetAttributes
BlockAttributes = ThisDrawing.Layouts.Item("Model").Block(Index).GetAttributes
For Each Attrib In BlockAttributes
Tagname = Attrib.TagString
Select Case VBA.UCase(Tagname)

Case "PN" & n:
'Set attribute value
Attrib.TextString = xlsheet.range(("DB" & (n + 71)))
Case "DES" & n:
'Set attribute value
Attrib.TextString = xlsheet.range(("DC" & (n + 71)))
Case "MAT" & n:
'Set attribute value
Attrib.TextString = xlsheet.range(("DD" & (n + 71)))
Case "QTY" & n:
'Set attribute value
Attrib.TextString = xlsheet.range(("DE" & (n + 71)))

End Select
Next Attrib
End If
End If
Next Index

Next n
'ThisDrawing.Application.Visible = False
  
  
'?è??PDF?ò??????
Set objdoc = ThisDrawing.Application.ActiveDocument
Set objLayout = objdoc.Layouts.Item("Model")
Set objPlot = objdoc.Plot
ThisDrawing.Application.ZoomExtents

' ?è???ò??"ú
If Not VBA.Trim(PrinterName) = "" Then
objLayout.ConfigName = PrinterName
Else
Exit Sub
End If

' ?è???ò???ù??±í
If Not VBA.Trim(Styles) = "" Then
objLayout.StyleSheet = Styles
Else
objLayout.StyleSheet = "acad.ctb"
End If

' ?è??????????
If AutoMedia Then
objLayout.CanonicalMediaName = "A3"
Else
If Not VBA.Trim(MediaName) = "" Then
objLayout.CanonicalMediaName = MediaName
Else
objLayout.CanonicalMediaName = "A3"
End If
End If

' ?è?????????"
objLayout.PaperUnits = acMillimeters
'objLayout.PaperUnits = acInches

' ?è??????????ò??·??ò
'objLayout.PlotRotation = ac0degrees     '×??ò
'objLayout.PlotRotation = ac180degrees
objLayout.PlotRotation = ac90degrees   '?á?ò
'objLayout.PlotRotation = ac270degrees

' ?è???????ò??±???
objLayout.StandardScale = acScaleToFit
'objLayout.UseStandardScale = False '????×?¨???ò??±???

' ?è??×?¨???ò??±???
'objLayout.SetCustomScale txtNumerator.Value, txtDenominator.Value

' ?è????????·??????ò??
objLayout.CenterPlot = True

' ?ò???±???????????????????í
objLayout.PlotWithLineweights = True

' ?è????·??????ò???ù??
objLayout.PlotWithPlotStyles = True

' ?ò???±??????????????ó
'objLayout.PlotHidden = False

' ?è???????ò??·???
If Copies >= 1 Then
objPlot.NumberOfCopies = CInt(Copies)
Else
objPlot.NumberOfCopies = 1
End If

' ?"?ò???í?ó±¨????""??????í?ó???????±??"??????????ò??????
objPlot.QuietErrorMode = True
ThisDrawing.PurgeAll
' ?????ú???±?°????
objdoc.Regen acAllViewports

' ?è???°?¨?ò??????ò??????°??ò?????ò????·??????ò??"ú
objdoc.SetVariable "BACKGROUNDPLOT", 0
ReDim Preserve ptMin(0 To 1)
ReDim Preserve ptMax(0 To 1)
ThisDrawing.ActiveLayout.SetWindowToPlot ptMin, ptMax
objLayout.PlotType = acWindow
objPlot.PlotToFile newpath & "\SEC\" & savename & ".pdf"
'?í????dwg????????
objdoc.SaveAs newpath & "\SEC\" & savename & ".dwg"

'??±??í???ó??"??????
ThisDrawing.Application.ActiveDocument.Close

'xlbook.Application.Close
'Set xlbook = Nothing


End Sub


Public Sub SEC()
Call PlotFunction("DWG to PDF.pc3", "acad.ctb", "A3", 1, True, True, False, True, True)
End Sub






