Attribute VB_Name = "Module3"

'button
'====================================

Sub button_data_input()
    Sheets("input").Visible = True
    Sheets("input").Activate
    Sheets("map").Visible = False
End Sub


Sub button_utility()
    Sheets("utility").Activate
End Sub

Sub button_sec_list()
    Sheets("Sec-List").Activate
End Sub

Sub button_map()

    
        Sheets("map").Visible = True
    
    If ActiveSheet.name = "map" Then
        Sheets("map").Activate
    Else
        ActiveSheet.Visible = False
        Sheets("map").Activate
    End If
    
End Sub

Sub button_form5()

UserForm5.Show


End Sub

Sub sheet1_input()
    Sheets(1).Activate
End Sub

Sub user_mode()

    Dim i As Integer
                Application.Visible = False
        For i = 1 To ActiveWorkbook.Sheets.Count - 1
            Sheets(i).Visible = False
        Next i
        Sheets("map").Visible = True
        Sheets("background").Visible = False
        Sheets("map").Activate

                Application.Visible = True
        UserForm5.Show
End Sub

