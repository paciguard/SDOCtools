Attribute VB_Name = "Module2"
Public Sub GAcreat()


Dim cap As AutoCAD.AcadApplication
Dim xlapp  As Excel.Application
Dim xlbook As Object
Dim xlsheet As Object
Dim ss As String
Dim i, Counter As Integer


Dim caddoc As AcadDocument, mospace As AcadModelSpace


    Set xlapp = GetObject(, "Excel.application")
  
   

ss = xlapp.ActiveWorkbook.Path
Set xlbook = GetObject(ss & "\ZTools.xls")
Set xlsheet = xlbook.Worksheets("sheet1")

If Dir(ss & "\GA", vbDirectory) = "" Then
MkDir ss & "\GA"
Else
MsgBox "GA文件夹已存在，请删除！"
Exit Sub
End If


'On Error Resume Next '忽略错误

   'Set cap = GetObject(, "Autocad.application")
     
     
   

'If Err.Number <> 0 Then '如果 Excel 没有运行则

  'Err.Clear

  'Set cap = CreateObject("Autocad.application")
 
  
'End If

 
Set cap = New AutoCAD.AcadApplication
  
   
Set caddoc = cap.ActiveDocument


Set mospace = caddoc.ModelSpace
Counter = Val(xlsheet.Range(("C" & 15)))


Dim dwgname As String, aada As Object

dwgname = ss & "\TEMPLATE\GA.dwg"
Set caddoc = cap.Documents.Open(ss & "\TEMPLATE\GA.dwg")


 cap.Visible = False
 
 For i = 40 To Counter + 39
   
    Dim fileName As String
    
    fileName = ss & "\TEMPLATE\GA.dvb"
    
   
   cap.LoadDVB fileName
  
    UserForm4.Show 0
    On Error Resume Next
    With UserForm4.ProgressBar1
    .Min = 0
    .Max = Counter
    .Scrolling = 1
    .Value = i - 39
    UserForm4.Caption = "已完成 " & Int((i - 39) * 100 / Counter) & "%, 请稍等！"
    xlsheet.Range(("C" & 19)).Value = i
    cap.RunMacro ("GA")
   cap.Visible = False
   
   Application.ScreenUpdating = True
   Unload UserForm4
   
   End With
   Next i
   

    



caddoc.Close False
Set caddoc = Nothing
AutoCAD.Application.Quit

Set cap = GetObject(, "AutoCAD.AcadApplication")
Set caddoc = cap.ActiveDocument
caddoc.Close False
cap.Quit



MsgBox "外形图生成完毕"
'Autocad.Application.Quit

 'xlbook.Application.Close
 'Set xlbook = Nothing
 'xlbook.Application.Quit
 
   
AutoCAD.Application.Quit

End Sub


