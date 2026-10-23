Attribute VB_Name = "Module7"
Sub Utility()

    Dim X, k, i, j, Counter As Integer
    Dim Cooling_Plan, str1, str2 As String
    Dim Sum As Double
    
'clear former contents

Sheets("Utility").Range("$d$13:$d$42").ClearContents
Sheets("Utility").Range("$f$13:$g$42").ClearContents
Sheets("Utility").Range("$m$13:$p$42").ClearContents

i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))
    
    For k = 0 To i - 1
    
        Cooling_Plan = Sheets("input").Range("$z$21").Offset(k, 0)
        j = Len(Cooling_Plan)
            str2 = ""
                For X = 1 To j
                    
                    str1 = VBA.Mid(Cooling_Plan, X, 1)
                    

                    If str1 <> " " Then
                          str2 = str2 & str1
                    Else
                         str1 = ""
                         str2 = str2 & str1
                    End If
                Sheets("Utility").Range("$d$13").Offset(k, 0) = str2
        
                Next X
                
If str2 <> "NO" And str2 <> "" And str2 <> "无" Then

'======================================
'split up cooling plan
    
    'calculate "+" amount
    Counter = 0
    For X = 1 To j
        If VBA.Mid(str2, X, 1) = "+" Then Counter = Counter + 1
    Next X
        
        If Counter = 0 Then
        
            Sheets("Utility").Range("$m$13").Offset(k, 0) = str2
        
        Else
            
            For X = 0 To Counter - 1
            
                Sheets("Utility").Range("$m$13").Offset(k, X) = VBA.Mid(str2, 1, InStr(str2, "+") - 1)
                str2 = VBA.Mid(str2, InStr(str2, "+") + 1, j)
                j = Len(str2)
            
            Next X
        Sheets("Utility").Range("$m$13").Offset(k, X) = str2
        End If
        
'split up cooling plan
'======================================
Sum = 0

        For X = 0 To 3
        
            If Sheets("Utility").Range("$m$13").Offset(k, X) <> "" Then
            
                Sum = Sum + Application.VLookup(Sheets("Utility").Range("$m$13").Offset(k, X), Sheets("UtilityData").Range("$d$8:$l$33"), Sheets("UtilityData").Range("$b$6").Offset(k, 0) + 2, False)
            
            Else
                Exit For
            End If
            
            Sheets("Utility").Range("$g$13").Offset(k, 0) = Sum & " m3/h"
        Next X
        
Else
        Sheets("Utility").Range("$g$13").Offset(k, 0) = "无"
                
End If

'==========================================
'Plan 62

        If InStr(Sheets("Utility").Range("$c$13").Offset(k, 0), "62") <> 0 Then
            Sheets("Utility").Range("$f$13").Offset(k, 0) = "0.06 m3/h"
        Else
            Sheets("Utility").Range("$f$13").Offset(k, 0) = "无"
        End If


    Next k


End Sub
