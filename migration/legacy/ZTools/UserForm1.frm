Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{FF8EAFA6-A871-47A2-BB5A-8317189E3190}{F15C43B1-1BC6-4138-AD53-C69C1D092465}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION As Long = &HC00000
Private Const SWP_NOMOVE = &H2
Dim hwnd


Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Sub Image2_Click()

End Sub

Private Sub CommandButton3_Click()
UserForm2.Show
End Sub

Private Sub Image1_Click()

End Sub

Private Sub CommandButton4_Click()
    Dim a, b As String
    Dim i As Integer
    
     
    If TextBox2.Text = "123" Then
        Application.Visible = False
        
'        For i = 3 To ActiveWorkbook.Sheets.Count - 1
'            Sheets(i).Visible = False
'        Next i
        Sheets("map").Visible = True
        Sheets("background").Visible = False
        Unload UserForm1
        
        Sheets("map").Activate
        Application.Visible = True
        
    Else
        MsgBox ("您的密码输入错误，请重新输入！")
        TextBox2.Text = ""
    End If
End Sub

Private Sub Label7_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
    Application.Quit
End Sub

Private Sub UserForm_Initialize()
    Dim rng As Range
     Dim IStyle As Long
     
    If Val(Application.version) < 9 Then
        hwnd = FindWindow("ThunderXFrame", Me.Caption)
    Else
        hwnd = FindWindow("ThunderDFrame", Me.Caption)
    End If
    SetWindowLong hwnd, GWL_STYLE, IStyle
    DrawMenuBar hwnd
End Sub


'====================
'原UserZone “OK”按钮代码

Private Sub CommandButton1_Click()
    Dim a, b As String
    Dim i As Integer
    Dim SheetCounter As Integer
    
     
    If TextBox1.Text = Sheets("data").Range("$b$1") Then
        Application.Visible = False
        
        SheetCounter = ActiveWorkbook.Sheets.Count
        For i = 3 To SheetCounter
            Sheets(i).Visible = False
        Next i
        Unload UserForm1
        Sheets(1).Visible = True
        
        Sheets("map").Activate
        Sheets("input").Visible = True
        Application.Visible = True
        
    Else
        MsgBox ("您的密码输入错误，请重新输入！")
        TextBox1.Text = ""
    End If
End Sub

Private Sub CommandButton2_Click()
    Sheets("background").Visible = True
    ThisWorkbook.Close False
End Sub

Private Sub Label1_Click()
    Label1.Caption = "HELLO" & Application.UserName
End Sub

