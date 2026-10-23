Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub workbook_open()
    
    Dim j As Integer
    
    
    Application.Visible = False
        
        j = Sheets.Count
        For i = 1 To j
            Sheets(i).Visible = True
        Next i
        Unload UserForm1
        
        
        UserForm1.Show
        
    Application.Visible = True
        

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
        Dim a, b As String
        Dim i, j As Integer
        Sheets("background").Visible = True
        j = Sheets.Count
        For i = 1 To j - 1
            Sheets(i).Visible = xlSheetHidden
        Next i
        
        ThisWorkbook.Save
        ThisWorkbook.Close False
End Sub
