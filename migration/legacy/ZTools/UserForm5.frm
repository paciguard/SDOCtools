Attribute VB_Name = "UserForm5"
Attribute VB_Base = "0{4168137F-CCB6-47F1-A928-B1CEAF192916}{B7061813-B2C7-48A6-97F7-9B7022F7BA91}"
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


Private Sub CommandButton3_Click()
    Unload UserForm5

    Call SecCreate
End Sub

Private Sub CommandButton10_Click()
Unload UserForm5
UserForm6.Show
End Sub

Private Sub CommandButton11_Click()
    Unload UserForm5
        Sheets(1).Visible = True
    
    If ActiveSheet.name = "Sheet1" Then
        Sheets(1).Activate
    Else
        ActiveSheet.Visible = False
        Sheets(1).Activate
    End If
    Sheets(1).Activate
End Sub

Private Sub CommandButton12_Click()
    Unload UserForm5

    Sheets("input").Visible = True
    
    If ActiveSheet.name = "input" Then
        Sheets("input").Activate
    Else
        ActiveSheet.Visible = False
        Sheets("input").Activate
    End If
End Sub

Private Sub CommandButton13_Click()
Unload UserForm5
End Sub

Private Sub CommandButton4_Click()
    Unload UserForm5
        Sheets("Utility").Visible = True
    
    If ActiveSheet.name = "Utility" Then
        Sheets("Utility").Activate
    Else
        ActiveSheet.Visible = False
        Sheets("Utility").Activate
    End If
    Call Utility
    MsgBox ("注意：外冲洗流量列、Plan 62列及其表头，可以根据实际订单进行手工修正。")
    
End Sub

Private Sub CommandButton5_Click()

End Sub

Private Sub CommandButton6_Click()
    Unload UserForm5
    Sheets(1).Visible = True
    
    If ActiveSheet.name = "Sheet1" Then
        Sheets(1).Activate
    Else
        ActiveSheet.Visible = False
        Sheets(1).Activate
    End If
    Call tranfer_data
End Sub

Private Sub CommandButton7_Click()
Unload UserForm5

        Sheets("Sec-List").Visible = True
    
    If ActiveSheet.name = "Sec-List" Then
        Sheets("Sec-List").Activate
    Else
        ActiveSheet.Visible = False
        Sheets("Sec-List").Activate
    End If

End Sub

Private Sub CommandButton8_Click()
    Unload UserForm5
        Sheets(1).Visible = True
    
    If ActiveSheet.name = "Sheet1" Then
        Sheets(1).Activate
    Else
        ActiveSheet.Visible = False
        Sheets(1).Activate
    End If
    
    Call GAcreat
End Sub

Private Sub CommandButton9_Click()
Unload UserForm5

        Sheets("PID").Visible = True
    
    If ActiveSheet.name = "PID" Then
        Sheets("PID").Activate
    Else
        ActiveSheet.Visible = False
        Sheets("PID").Activate
    End If
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

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




Private Sub CommandButton1_Click()
Unload UserForm5
Call GAcreat
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub Label1_Click()
    Label1.Caption = "HELLO" & Application.UserName
End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
