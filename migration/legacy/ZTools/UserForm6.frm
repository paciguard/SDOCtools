Attribute VB_Name = "UserForm6"
Attribute VB_Base = "0{D16640AC-BCF0-4F27-B74F-A13E1C153645}{638B5124-E19D-4DC8-8F27-DA94B2255402}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub CommandButton1_Click()
   Dim a, b As String
    Dim i As Integer
    
     
    If TextBox2.Text = Sheets("Data").Range("$b$2") Then
        
        Unload UserForm6
        Application.Visible = False
        For i = 1 To ActiveWorkbook.Sheets.Count
            Sheets(i).Visible = True
        Next i
            Sheets("Data").Activate
        Application.Visible = True
    Else
        MsgBox ("您的密码输入错误，请重新输入！")
        TextBox2.Text = ""
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload UserForm6
    UserForm5.Show
    
End Sub

Private Sub TextBox2_Change()

End Sub
