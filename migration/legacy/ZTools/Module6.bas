Attribute VB_Name = "Module6"
Sub PIDCreate_For_Blocks()

Dim i, k, X, j, Counter As Integer
Dim PID_Plan, str1, str2 As String
Dim cap As AutoCAD.AcadApplication
Dim xlapp  As Excel.Application
Dim xlbook As Object
Dim xlsheet As Object
Dim ss As String
Dim w As Integer


Dim caddoc As AcadDocument, mospace As AcadModelSpace


    Set xlapp = GetObject(, "Excel.application")
  
   

ss = xlapp.ActiveWorkbook.Path
Set xlbook = GetObject(ss & "\ZTools.xls")
Set xlsheet = xlbook.Worksheets("sheet1")


If Dir(ss & "\PID", vbDirectory) = "" Then
MkDir ss & "\PID"
Else
MsgBox "PID文件夹已存在，请删除！"
Exit Sub
End If

    i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))
    Sheets(1).Range("$cg$40:$cm$69").ClearContents
    
    'On Error Resume Next '忽略错误

   'Set cap = GetObject(, "Autocad.application")
     
     
   

'If Err.Number <> 0 Then '如果 Excel 没有运行则

  'Err.Clear

  'Set cap = CreateObject("Autocad.application")
 
  
'End If

 
Set cap = New AutoCAD.AcadApplication
  
   
Set caddoc = cap.ActiveDocument


Set mospace = caddoc.ModelSpace
'Counter = Val(xlsheet.Range(("C" & 15)))


Dim dwgname As String, aada As Object

dwgname = ss & "\TEMPLATE\PID.dwg"
Set caddoc = cap.Documents.Open(ss & "\TEMPLATE\PID.dwg")


 cap.Visible = False
    
    
    
    
    
For k = 0 To i - 1
    
    
    PID_Plan = Sheets(1).Range("$e$40").Offset(k, 0)
        
    'delete "space"
            j = Len(PID_Plan)
            str2 = ""
                For X = 1 To j
                    
                    str1 = VBA.Mid(PID_Plan, X, 1)
                    

                    If str1 <> " " Then
                          str2 = str2 & str1
                    Else
                         str1 = ""
                         str2 = str2 & str1
                    End If
                Next X
    ' delete "Plan" or "P" or "p"
            j = Len(str2)
            
                For X = 1 To j
                
                    If IsNumeric(VBA.Mid(str2, X, 1)) = True Then Exit For
                    
                Next X
       str2 = VBA.Mid(str2, X, j)
             
      Sheets(1).Range("$cg$40").Offset(k, 0) = str2

'================================================
'split up flush plan
    
    'calculate "+" amount
    Counter = 0
    For X = 1 To j
        If VBA.Mid(str2, X, 1) = "+" Then Counter = Counter + 1
    Next X
        
        If Counter = 0 Then
        
            Sheets(1).Range("$ch$40").Offset(k, 0) = str2
        
        Else
            
            For X = 1 To Counter
            
                Sheets(1).Range("$cg$40").Offset(k, X) = VBA.Mid(str2, 1, InStr(str2, "+") - 1)
                str2 = VBA.Mid(str2, InStr(str2, "+") + 1, j)
                j = Len(str2)
            
            Next X
             Sheets(1).Range("$cg$40").Offset(k, X) = str2
        End If

    

'================================================
'增加CAD程序
'================================================
w = k + 40
    
    'Dim FileName As String
    
    'FileName = ss & "\Project.dvb"
    
   
   'cap.LoadDVB FileName
   
   
   Dim fileName As String
    
    fileName = ss & "\TEMPLATE\PID.dvb"
    
   
   cap.LoadDVB fileName
  
    UserForm4.Show 0
    On Error Resume Next
    With UserForm4.ProgressBar1
    .Min = 0
    .Max = i
    .Scrolling = 1
    .Value = w - 39
    UserForm4.Caption = "已完成 " & Int((w - 39) * 100 / i) & "%, 请稍等！"
    xlsheet.Range(("C" & 19)).Value = w
    cap.RunMacro ("PID")
   cap.Visible = False
   
   Application.ScreenUpdating = True
   Unload UserForm4
   
   End With

Next k



'Sheets(1).Range("a1").Select

Set caddoc = Nothing
AcadApplication.ActiveDocument.Close False
AutoCAD.Application.Quit

Set cap = GetObject(, "AutoCAD.AcadApplication")
Set caddoc = cap.ActiveDocument
caddoc.Close False
cap.Quit

MsgBox "PID生成完毕"


End Sub

