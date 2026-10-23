Attribute VB_Name = "Module04"
'Public Sub SecCreate()
'
'Dim i, x, y, k, a, b, c As Integer
'Dim List_Name, Pump_Type_Sec As String
'
''CAD part1
''========================================
''CAD part1
'
'Dim cap As AutoCAD.AcadApplication
'Dim xlapp  As Excel.Application
'Dim xlbook As Object
'Dim xlsheet As Object
'Dim ss As String
'Dim j, Counter As Integer
'
'
'Dim caddoc As AcadDocument, mospace As AcadModelSpace
'
'
'    Set xlapp = GetObject(, "Excel.application")
'
'
'
'ss = xlapp.ActiveWorkbook.Path
'Set xlbook = GetObject(ss & "\ZTools.xls")
'Set xlsheet = xlbook.Worksheets("sheet1")
'
'    Sheets(1).Visible = True
'
'    If ActiveSheet.name = "Sheet1" Then
'        Sheets(1).Activate
'    Else
'        ActiveSheet.Visible = False
'        Sheets(1).Activate
'    End If
'
'
'If Dir(ss & "\SEC", vbDirectory) = "" Then
'MkDir ss & "\SEC"
'Else
'MsgBox "SEC文件夹已存在，请删除！"
'Exit Sub
'End If
'
'    i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))
'
''check availablity of list name
'    'For k = 0 To i - 1
'
'        'List_Name = Sheets(1).Range("$cc$40").Offset(k, 0)
'            'For x = 0 To 34
'                'If Sheets(1).Range("$cc$40").Offset(k, 0) = Sheets("List-all").Range("$b$2").Offset(0, x) Then Exit For
'            'Next x
'
'        'If x = 34 Then
'            'MsgBox ("第" & k + 1 & "行，泵标准号或轴承冷却形式选择错误 或 需要补充数据源，请与产品负责人联系，谢谢。")
'            'Exit Sub
'        'End If
'    'Next k
'
''transfer section part list
'
''On Error Resume Next '忽略错误
'
'   'Set cap = GetObject(, "Autocad.application")
'
'
'
'
''If Err.Number <> 0 Then '如果 Excel 没有运行则
'
'  'Err.Clear
'
'  'Set cap = CreateObject("Autocad.application")
'
'
''End If
'
'
'Set cap = New AutoCAD.AcadApplication
'
'
'Set caddoc = cap.ActiveDocument
'
'
'Set mospace = caddoc.ModelSpace
''Counter = Val(xlsheet.Range(("C" & 15)))
'
'
'Dim dwgname As String, aada As Object
'
'dwgname = ss & "\TEMPLATE\SEC.dwg"
'Set caddoc = cap.Documents.Open(ss & "\TEMPLATE\SEC.dwg")
'
'
' cap.Visible = False
'
''CAD part1
''======================
''CAD part1
'
'    Sheets(1).Range("$cf$40:$cf$69").ClearContents
'
'    For k = 0 To i - 1
'
'        Sheets(1).Range("$db$72:$de$111").ClearContents
'
'        List_Name = Sheets(1).Range("$cc$40").Offset(k, 0)
'
'            'list material class from sheet "Sec-List" to sheet 1
'            Sheets(1).Range("$ce$40").Offset(k, 0) = Sheets("Sec-List").Range("$f$10").Offset(k, 0)
'
'            'get column number (x) of the section name
'            For x = 0 To 36
'                If Sheets(1).Range("$cc$40").Offset(k, 0) = Sheets("List-all").Range("$b$2").Offset(0, x) Then Exit For
'            Next x
'
'        For a = 0 To 39
'
'
'            'list the part number according the current section name column
'            Sheets(1).Range("$db$72").Offset(a, 0) = Sheets("List-all").Range("$b$3").Offset(a, x)
'
'
'            For y = 0 To 150
'                'find current line number
'                If Sheets(1).Range("$db$72").Offset(a, 0) = Sheets("List-detail").Range("$c$3").Offset(y, 0) Then Exit For
'
'            Next y
'                'get the part name and amount list using "offset"
'                Sheets(1).Range("$dc$72").Offset(a, 0) = Sheets("List-detail").Range("$c$3").Offset(y, 1)
'                Sheets(1).Range("$de$72").Offset(a, 0) = Sheets("List-detail").Range("$c$3").Offset(y, 2)
'
'                'get the material information using "offset"
'                For b = 0 To 13
'
'                    If Sheets("input").Range("$o$21").Offset(k, 0) = Sheets("List-detail").Range("$c$2").Offset(0, b) Then Exit For
'                Next b
'
'                'specil shaft material if necessary
'                If Sheets(1).Range("$db$72").Offset(a, 0) = "210.01" And Sheets("Sec-List").Range("$g$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$g$10").Offset(k, 0)
'                'specil key (impeller to shaft) material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "940.01" And Sheets("Sec-List").Range("$h$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$h$10").Offset(k, 0)
'                'specil impeller material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "230.01" And Sheets("Sec-List").Range("$i$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$i$10").Offset(k, 0)
'                 'specil impeller nut material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "922.01" And Sheets("Sec-List").Range("$l$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$l$10").Offset(k, 0)
'                 'specil spacer sleeve material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "525.02" And Sheets("Sec-List").Range("$m$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$m$10").Offset(k, 0)
'                  'specil impeller ring material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "503.01" And Sheets("Sec-List").Range("$j$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$j$10").Offset(k, 0)
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "503.02" And Sheets("Sec-List").Range("$j$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$j$10").Offset(k, 0)
'                   'specil stationary ring material if necessary
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "502.01" And Sheets("Sec-List").Range("$k$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$k$10").Offset(k, 0)
'                ElseIf Sheets(1).Range("$db$72").Offset(a, 0) = "502.02" And Sheets("Sec-List").Range("$k$10").Offset(k, 0) <> "" Then
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("Sec-List").Range("$k$10").Offset(k, 0)
'                Else
'                    Sheets(1).Range("$dd$72").Offset(a, 0) = Sheets("List-detail").Range("$c$3").Offset(y, b)
'                End If
'
'        Next a
''===========================
''judge spacer sleeve should be existed or not, if not, adjust part list in sheet 1
'
'        Pump_Type_Sec = VBA.Mid(Sheets(1).Range("$c$40").Offset(k, 0), 1, 2)
'
'        If Pump_Type_Sec <> "ZA" Then
'
'
'            For a = 0 To 39
'
'                If VBA.Mid(Sheets(1).Range("$db$72").Offset(a, 0), 1, 3) = "210" Then Exit For
'
'            Next a
'
'            If InStr(Sheets(1).Range("$db$72").Offset(a, 2), "Cr13") <> 0 Or Sheets(1).Range("$db$72").Offset(a, 2) = "A276 Type 420" Then
'
'                If InStr(Sheets("input").Range("$o$21").Offset(k, 0), "A-7") <> 0 Or InStr(Sheets("input").Range("$o$21").Offset(k, 0), "D") <> 0 Or InStr(Sheets("input").Range("$o$21").Offset(k, 0), "A-8") <> 0 Then
'                    Sheets(1).Range("$cf$40").Offset(k, 0) = "Y"
'                    For c = 0 To 39
'
'                        If Sheets(1).Range("$db$72").Offset(c, 0) = "525.02" Then Exit For
'
'                    Next c
'
'                    Sheets(1).Range("$db$" & (72 + c + 1) & ":" & "$de$112").Copy
'
'                    Sheets(1).Range("$db$" & (72 + c)).Select
'                    Sheets(1).Paste
'
'                Else
'
'
'                    For c = 0 To 39
'
'                        If Sheets(1).Range("$db$72").Offset(c, 0) = "400.03" Then Exit For
'
'                    Next c
'
'                    Sheets(1).Range("$db$" & (72 + c + 1) & ":" & "$de$112").Copy
'
'                    Sheets(1).Range("$db$" & (72 + c)).Select
'                    Sheets(1).Paste
'
'                End If
'
'            Else
'                    For c = 0 To 39
'
'                        If Sheets(1).Range("$db$72").Offset(c, 0) = "400.03" Then Exit For
'
'                    Next c
'
'                    Sheets(1).Range("$db$" & (72 + c + 1) & ":" & "$de$112").Copy
'
'                    Sheets(1).Range("$db$" & (72 + c)).Select
'                    Sheets(1).Paste
'
'            End If
'
'
'
'        End If
'
''CAD part2
'
''===========================
'
''CAD part2
'
' j = k + 40
'
'    'Dim FileName As String
'
'    'FileName = ss & "\Project.dvb"
'
'
'   'cap.LoadDVB FileName
'
'
'   Dim FileName As String
'
'    FileName = ss & "\TEMPLATE\SEC.dvb"
'
'
'   cap.LoadDVB FileName
'
'    UserForm4.Show 0
'    On Error Resume Next
'    With UserForm4.ProgressBar1
'    .Min = 0
'    .Max = i
'    .Scrolling = 1
'    .Value = j - 39
'    UserForm4.Caption = "已完成 " & Int((j - 39) * 100 / i) & "%, 请稍等！"
'    xlsheet.Range(("C" & 19)).Value = j
'    cap.RunMacro ("SEC")
'   cap.Visible = False
'
'   Application.ScreenUpdating = True
'   Unload UserForm4
'
'   End With
'
'
'
'    Next k
'
'
'
'caddoc.Close False
'Set caddoc = Nothing
'AcadApplication.ActiveDocument.Close False
'AutoCAD.Application.Quit
'
'Set cap = GetObject(, "AutoCAD.AcadApplication")
'Set caddoc = cap.ActiveDocument
'caddoc.Close False
'cap.Quit
'
'
'Sheets(1).Range("a1").Select
'
'
'
'
'
'MsgBox "剖面图生成完毕"
''CAD part2
''============================
''CAD part2
'AutoCAD.Application.Quit
'
'End Sub
