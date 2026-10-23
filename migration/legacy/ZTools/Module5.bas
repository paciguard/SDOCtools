Attribute VB_Name = "Module5"
Sub PIDCreate_For_files()

Dim i, k, X, j As Integer
Dim PID_Plan, str1, str2 As String

    i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))
    Sheets(1).Range("$cg$40:$cg$69").ClearContents
    
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
'增加CAD程序
'================================================


Next k

End Sub
