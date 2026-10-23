Attribute VB_Name = "Module4"
Option Explicit

Public Sub SecCreate()

    Dim cap As AutoCAD.AcadApplication
    Dim caddoc As AcadDocument
    
    Dim xlapp As Excel.Application
    Dim xlbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet
    
    Dim ss As String
    Dim i As Long, k As Long
    Dim List_Name As String, Pump_Type_Sec As String, CZBearingsize As String, pumpsize As String
    Dim X As Long, a As Long, b As Long, c As Long, Y As Long
    Dim Counter As Long
    Dim dwgTemplate As String, dvbFile As String
    Dim GAfolder As String        ' Actually SEC folder
    
    On Error GoTo ErrorHandler
    
    '=== Excel Setup ===
    Set xlapp = GetObject(, "Excel.Application")
    Set xlbook = xlapp.ActiveWorkbook
    Set xlsheet = xlbook.Worksheets("Sheet1")
    
    ss = xlbook.Path
    
    ' Make sure Sheet1 is visible and active
    xlbook.Sheets(1).Visible = True
    xlbook.Sheets(1).Activate
    
    '=== Create SEC folder ===
    Dim SECfolder As String
    SECfolder = ss & "\SEC"
    
    If Dir(SECfolder, vbDirectory) = "" Then
        MkDir SECfolder
    Else
        MsgBox "SEC文件夹已存在，请先删除该文件夹后再运行！", vbCritical, "提示"
        Exit Sub
    End If
    
    '=== Get number of pumps ===
    i = Application.CountA(xlbook.Sheets("input").Range("N21:N50"))
    If i = 0 Then
        MsgBox "输入表中没有数据！", vbExclamation
        Exit Sub
    End If
    
    '=== Template paths ===
    dwgTemplate = ss & "\TEMPLATE\SEC.dwg"
    dvbFile = ss & "\TEMPLATE\SEC.dvb"
    
    If Dir(dwgTemplate) = "" Then
        MsgBox "模板文件不存在: " & dwgTemplate, vbCritical
        Exit Sub
    End If
    If Dir(dvbFile) = "" Then
        MsgBox "DVB文件不存在: " & dvbFile, vbCritical
        Exit Sub
    End If
    
    '=== Start AutoCAD ===
    Set cap = New AutoCAD.AcadApplication
    cap.Visible = False
    
    Set caddoc = cap.Documents.Open(dwgTemplate)
    
    ' Clear previous temporary data
    xlbook.Sheets(1).Range("CF40:CF69").ClearContents
    
    '=== Main Loop ===
    For k = 0 To i - 1
        
        Dim r As Long
        r = 40 + k                     ' Current data row in Sheet1
        
        ' Clear temporary area for this pump
        xlbook.Sheets(1).Range("DB72:DE111").ClearContents
        pumpsize = xlbook.Sheets(1).Cells(r, 3).Value
        List_Name = xlbook.Sheets(1).Cells(r, 81).Value          ' CC column = 81
        
        ' Material class
        xlbook.Sheets(1).Cells(r, 79).Value = xlbook.Sheets("Sec-List").Cells(10 + k, 6).Value   ' CE column
        
        ' Find column number of current section list name in "List-all"
        For X = 0 To 51
            If List_Name = xlbook.Sheets("List-all").Cells(2, 2 + X).Value Then Exit For
        Next X
        
        ' Check if list name was found
        If X > 39 Then
            MsgBox "第" & (k + 1) & "行，泵标准号或轴承冷却形式选择错误，请检查数据！", vbCritical
            GoTo CleanExit
        End If
        
        '=== Fill part list for this pump ===
        For a = 0 To 51
            
            Dim rowDB As Long
            rowDB = 72 + a
            
            ' Part number
            xlbook.Sheets(1).Cells(rowDB, 106).Value = xlbook.Sheets("List-all").Cells(3 + a, 2 + X).Value   ' DB column = 106
            
            ' Find detail in List-detail
            For Y = 0 To 150
                If CStr(xlbook.Sheets(1).Cells(rowDB, 106).Value) = CStr(xlbook.Sheets("List-detail").Cells(3 + Y, 3).Value) Then Exit For
            Next Y
            
            ' Part name and amount
            xlbook.Sheets(1).Cells(rowDB, 107).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 4).Value   ' DC   107
            
            
            xlbook.Sheets(1).Cells(rowDB, 109).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 5).Value   ' DE (quantity?)  109
            
            ' Material
            For b = 0 To 13
                If xlbook.Sheets("input").Cells(21 + k, 15).Value = xlbook.Sheets("List-detail").Cells(2, 3 + b).Value Then Exit For
            Next b
            
            ' Special material overrides
            Dim partNo As String
            partNo = xlbook.Sheets(1).Cells(rowDB, 106).Value
            
            Select Case partNo
                Case "210.01"
                    If xlbook.Sheets("Sec-List").Cells(10 + k, 7).Value <> "" Then
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("Sec-List").Cells(10 + k, 7).Value
                    Else
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 3 + b).Value
                    End If
                    
                Case "940.01"
                    If xlbook.Sheets("Sec-List").Cells(10 + k, 8).Value <> "" Then
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("Sec-List").Cells(10 + k, 8).Value
                    Else
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 3 + b).Value
                    End If
                    
                Case "230.01"
                    If xlbook.Sheets("Sec-List").Cells(10 + k, 9).Value <> "" Then
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("Sec-List").Cells(10 + k, 9).Value
                    Else
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 3 + b).Value
                    End If
                    
                Case "922.01", "525.02", "503.01", "503.02", "502.01", "502.02"
                    Dim specialCol As Long
                    Select Case partNo
                        Case "922.01": specialCol = 12  ' L
                        Case "525.02": specialCol = 13  ' M
                        Case "503.01", "503.02": specialCol = 10 ' J
                        Case "502.01", "502.02": specialCol = 11 ' K
                    End Select
                    
                    If xlbook.Sheets("Sec-List").Cells(10 + k, specialCol).Value <> "" Then
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("Sec-List").Cells(10 + k, specialCol).Value
                    Else
                        xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 3 + b).Value
                    End If
                    
                Case Else
                    xlbook.Sheets(1).Cells(rowDB, 108).Value = xlbook.Sheets("List-detail").Cells(3 + Y, 3 + b).Value
            End Select
            
        Next a
        
        '=== Spacer sleeve logic (delete row if needed) ===
        Pump_Type_Sec = Left(xlbook.Sheets(1).Cells(r, 3).Value, 2)   ' Column C
        
        If Pump_Type_Sec <> "ZA" And Pump_Type_Sec <> "XA" And Pump_Type_Sec <> "CZ" Then
            
            ' Find shaft row
            For a = 0 To 51
                If Left(xlbook.Sheets(1).Cells(72 + a, 106).Value, 3) = "210" Then Exit For
            Next a
            
            Dim shaftMaterial As String
            shaftMaterial = xlbook.Sheets(1).Cells(72 + a, 108).Value
            
            If InStr(shaftMaterial, "Cr13") > 0 Or shaftMaterial = "A276 Type 420" Then
                
                If InStr(xlbook.Sheets("input").Cells(21 + k, 15).Value, "A-7") > 0 Or _
                   InStr(xlbook.Sheets("input").Cells(21 + k, 15).Value, "D") > 0 Or _
                   InStr(xlbook.Sheets("input").Cells(21 + k, 15).Value, "A-8") > 0 Then
                    
                    xlbook.Sheets(1).Cells(r, 82).Value = "Y"          ' CF column
                    
                    ' Delete spacer sleeve row (525.02)
                    For c = 0 To 51
                        If xlbook.Sheets(1).Cells(72 + c, 106).Value = "525.02" Then Exit For
                    Next c
                    xlbook.Sheets(1).Range("DB" & (72 + c + 1) & ":DE112").Copy
                    xlbook.Sheets(1).Range("DB" & (72 + c)).PasteSpecial xlPasteAll
                    
                Else
                    ' Delete 400.03 row
                    For c = 0 To 51
                        If xlbook.Sheets(1).Cells(72 + c, 106).Value = "400.03" Then Exit For
                    Next c
                    xlbook.Sheets(1).Range("DB" & (72 + c + 1) & ":DE112").Copy
                    xlbook.Sheets(1).Range("DB" & (72 + c)).PasteSpecial xlPasteAll
                End If
                
            Else
                ' Normal case - delete 400.03
                For c = 0 To 51
                    If xlbook.Sheets(1).Cells(72 + c, 106).Value = "400.03" Then Exit For
                Next c
                xlbook.Sheets(1).Range("DB" & (72 + c + 1) & ":DE112").Copy
                xlbook.Sheets(1).Range("DB" & (72 + c)).PasteSpecial xlPasteAll
            End If
            
        End If
        
        '=== Run AutoCAD SEC Macro ===
        xlbook.Sheets(1).Cells(19, 3).Value = k + 40                   ' C19
        
        UserForm4.Show vbModeless
        
        With UserForm4
            .ProgressBar1.Min = 0
            .ProgressBar1.Max = i
            .ProgressBar1.Value = k + 1
            .Caption = "已完成 " & Format((k + 1) / i, "0%") & "，请稍等..."
        End With
        
        cap.LoadDVB dvbFile
        cap.RunMacro "SEC"
        
        DoEvents
        Unload UserForm4
        
    Next k
    
    '=== Cleanup ===
CleanExit:
    If Not caddoc Is Nothing Then
        caddoc.Close False
        Set caddoc = Nothing
    End If
    
    If Not cap Is Nothing Then
        cap.Quit
        Set cap = Nothing
    End If
    
    xlbook.Sheets(1).Range("A1").Select
    MsgBox "剖面图生成完毕！" & vbCrLf & "共处理 " & i & " 个位号。", vbInformation
    
    Exit Sub

ErrorHandler:
    MsgBox "发生错误: " & Err.Description, vbCritical, "SecCreate 错误"
    
    On Error Resume Next
    If Not caddoc Is Nothing Then caddoc.Close False
    If Not cap Is Nothing Then cap.Quit
    On Error GoTo 0

End Sub

