Attribute VB_Name = "Module01"
''Button
''=========================================
'Option Explicit
'Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
'() '  Alias "ShellExecuteA" (ByVal hwnd As Long, _
'  ByVal lpOperation As String, ByVal lpFile As String, _
'  ByVal lpParameters As String, ByVal lpDirectory As String, _
'  ByVal nShowCmd As Long) As Long
'
'
'
'
'Sub tranfer_data()
'
'    Dim Pump_Type, Pump_Type2, Pump_Type_all As Variant
'    Dim i, j, k, x, z, Counter, Counter1, Counter2, y As Integer
'    Dim str1, str2, str3, str4, str5, str6, LK, name As String
'    Dim other_weight As Integer
'    Dim CounterProgress, Progress, CounterSum As Integer
'    Dim PctDone As Single
'    Dim xCheck, yCheck As Integer
'    Dim cell As Range
'    Dim CounterCheck As Integer
'    Dim abcd As String
'
'    'save information
'    ThisWorkbook.Save
'
'    'count available rows
'    i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))
'
'    Sheets(1).Range("$c$15") = i
'
'    'check mandatory filling in sheet "input"
'    For k = 0 To i - 1
'
'
'       '======================
'        If Sheets("input").Range("$i$53") = "GB9113.1-2000" Then
'            For Counter1 = 0 To 2
'
'                If InStr(Sheets(1).Range("$c$12").Offset(Counter1, 0), "P") = 0 Then
'
'                    MsgBox ("请将GB标准的法兰信息填写完整")
'                    Exit Sub
'                Else
'                End If
'
'            Next Counter1
'        Else
'
'        End If
'       '======================
'        If Sheets("input").Range("$u$21").Offset(k, 0) = "" Then
'            MsgBox ("基座号列信息不得为空")
'            Exit Sub
'        Else
'
'        End If
'        '======================
'        If Sheets("input").Range("$v$21").Offset(k, 0) = "" Then
'            MsgBox ("电机电压列信息不得为空")
'            Exit Sub
'        Else
'
'        End If
'        '======================
'
'        If Sheets("input").Range("$v$21").Offset(k, 0) = "6000V" Or Sheets("input").Range("$v$21").Offset(k, 0) = "10000V" Then
'            MsgBox ("请注意进行电机长度及底座长度校核")
'
'        Else
'
'        End If
'        '======================
'         If Sheets("input").Range("$aa$21").Offset(k, 0) = "" Then
'            MsgBox ("轴承形式列信息不得为空")
'            Exit Sub
'        Else
'
'        End If
'
'    Next k
'
'
'Sheets(1).Activate
''===================
'
''Application.Visible = False
'
''====================
'     'clear former contents
'
'    Sheets(1).Range("$b$40:$i$69").ClearContents
'    Sheets(1).Range("$k$40:$cc$69").ClearContents
'    Sheets("Sec-List").Range("$f$10:$f$39").ClearContents
'
''/////////////////////////////////////////////
'
''calculate cells amount for progress
'
'CounterSum = 75 * i
'
'CounterProgress = 1
'
'
'UserForm2.Show 0
'
'
'
''//////////////////////////////////////////////
''===================================================================fill data
'
'
'
'    For k = 0 To i - 1
'
'    'link item no.
'    Sheets(1).Range("$b$40").Offset(k, 0) = Sheets("input").Range("$d$21").Offset(k, 0)
'    'link service name
'    Sheets(1).Range("$d$40").Offset(k, 0) = Sheets("input").Range("$f$21").Offset(k, 0)
'    'link motor
'    Sheets(1).Range("$k$40").Offset(k, 0) = Sheets("input").Range("$s$21").Offset(k, 0)
'    'link motor manufacture
'    Sheets(1).Range("$n$40").Offset(k, 0) = Sheets("input").Range("$t$21").Offset(k, 0)
'    'link motor manufacture
'    Sheets(1).Range("$n$40").Offset(k, 0) = Sheets("input").Range("$t$21").Offset(k, 0)
'    'link motor voltage
'        j = Len(Sheets("input").Range("$v$21").Offset(k, 0))
'
'    Sheets(1).Range("$l$40").Offset(k, 0) = VBA.Mid(Sheets("input").Range("$v$21").Offset(k, 0), 1, j - 1)
'    'link motor power
'    Sheets(1).Range("$m$40").Offset(k, 0) = Sheets("input").Range("$w$21").Offset(k, 0)
'    'link speed
'    Sheets(1).Range("$o$40").Offset(k, 0) = Sheets("input").Range("$r$21").Offset(k, 0)
'    'link Material class to sheet "Sec-List"
'    Sheets("Sec-List").Range("$f$10").Offset(k, 0) = Sheets("input").Range("$o$21").Offset(k, 0)
'
'
'    'fill and rewrite flush plan
'        If Sheets("input").Range("$z$21").Offset(k, 0) = "" Or Sheets("input").Range("$z$21").Offset(k, 0) = "无" Or Sheets("input").Range("$z$21").Offset(k, 0) = "NO" Then
'
'
'
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "P") <> 0 Or InStr(Sheets("input").Range("$y$21").Offset(k, 0), "p") <> 0 Then
'                Sheets(1).Range("$e$40").Offset(k, 0) = Sheets("input").Range("$y$21").Offset(k, 0)
'            Else
'                Sheets(1).Range("$e$40").Offset(k, 0) = "Plan " & Sheets("input").Range("$y$21").Offset(k, 0)
'            End If
'
'        Else
'
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "P") <> 0 Or InStr(Sheets("input").Range("$y$21").Offset(k, 0), "p") <> 0 Then
'                Sheets(1).Range("$e$40").Offset(k, 0) = Sheets("input").Range("$y$21").Offset(k, 0) & "+" & Sheets("input").Range("$z$21").Offset(k, 0)
'            Else
'                Sheets(1).Range("$e$40").Offset(k, 0) = "Plan " & Sheets("input").Range("$y$21").Offset(k, 0) & "+" & Sheets("input").Range("$z$21").Offset(k, 0)
'            End If
'
'        End If
'
'
''/////////////////////////////////////////////
'
''calculate progress course---1
'
'    Progress = 5 * CounterProgress
'
'    CounterProgress = CounterProgress + 1
'
'
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'
''=======================================================
'        Pump_Type_all = Sheets("input").Range("$n$21").Offset(k, 0)
'
'    'filter "space" and assignment pump type to sheet1
'            j = Len(Pump_Type_all)
'            str2 = ""
'                For x = 1 To j
'
'                    str1 = VBA.Mid(Pump_Type_all, x, 1)
'                    If Asc(str1) = 32 Or Asc(str1) = 63 Then
'                    str1 = ""
'                    End If
'
'                    If str1 <> " " Then
'                          str2 = str2 & str1
'                    Else
'                         str1 = ""
'                         str2 = str2 & str1
'                    End If
'                Sheets(1).Range("$c$40").Offset(k, 0) = str2
'
'                Next x
'
''/////////////////////////////////////////////
'
''calculate progress course---2
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'    ' str2 is the right format pump type for search
'    ' delete character "O" since the dimensions are the same between "ZA" and "ZAO", as well as "ZE" and "ZEO"
'
'        If VBA.Mid(str2, 3, 1) = "O" Or VBA.Mid(str2, 3, 1) = "o" Or VBA.Mid(str2, 3, 1) = "G" Or VBA.Mid(str2, 3, 1) = "g" Then
'            str4 = VBA.Mid(str2, 1, 2) & VBA.Mid(str2, 4, Len(str2) - 3)
'            Pump_Type2 = VBA.Mid(str2, 1, 3)
'        Else
'            str4 = str2
'            Pump_Type2 = VBA.Mid(str2, 1, 2)
'        End If
'        Pump_Type = VBA.Mid(str2, 1, 2)
'    'str4 is a processing value for search which has been filtered "O" character
'
'    'transfer ZA drain block name---str6 is search key
'        If Pump_Type = "ZA" Or Pump_Type = "XA" Then
'            z = Len(str4)
'                For x = 1 To z
'
'                    If VBA.Mid(str4, x, 1) = "-" Then Exit For
'
'                Next x
'            str6 = VBA.Mid(str4, 3, x - 2) & VBA.Mid(str4, x + 2, 3)
'
'            Sheets(1).Range("$br$40").Offset(k, 0) = Application.VLookup(str6, Sheets("Data1").Range("$c$10:$d$62"), 2, False)
'        Else
'            Sheets(1).Range("$br$40").Offset(k, 0) = "N"
'        End If
'
'    'calculate LK value
'
'        z = Len(str4)
'            For x = 1 To z
'
'                If VBA.Mid(str4, x, 1) = "-" Then Exit For
'
'            Next x
'        LK = "LK" & VBA.Mid(str4, x + 1, 1)
'
'
'
'    'transfer pump weight
'        Sheets(1).Range("$f$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 24, False)
'
'
'
''/////////////////////////////////////////////
'
''calculate progress course---3
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'    'transfer discharge and suction nozzle flange information
'        If Sheets("input").Range("$i$53") <> "ASME" Then
'            'transfer discharge and suction nozzle flange information
'            Sheets(1).Range("$s$40").Offset(k, 0) = "DN" & Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False)
'            Sheets(1).Range("$t$40").Offset(k, 0) = "DN" & Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 3, False)
'            'transfer FI flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "32") <> 0 Then
'                Sheets(1).Range("$ar$40").Offset(k, 0) = "DN15"
'
'            Else
'                Sheets(1).Range("$ar$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$as$40").Offset(k, 0) = "NO"
'            End If
'
'            'transfer BI/BO flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "54") <> 0 Then
'                Sheets(1).Range("$av$40").Offset(k, 0) = "DN15"
'
'            Else
'                Sheets(1).Range("$av$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$aw$40").Offset(k, 0) = "NO"
'            End If
'
'            'transfer QI/QO flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "62") <> 0 Then
'                Sheets(1).Range("$ap$40").Offset(k, 0) = "DN15"
'
'            Else
'                Sheets(1).Range("$ap$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$aq$40").Offset(k, 0) = "NO"
'            End If
'
'            'transfer CI/CO flange information
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'                Sheets(1).Range("$al$40").Offset(k, 0) = "DN20"
'
'            Else
'                Sheets(1).Range("$al$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "NO"
'            End If
'
'        Else
'            Sheets(1).Range("$s$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)
'            Sheets(1).Range("$t$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 5, False)
'            'transfer FI flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "32") <> 0 Then
'                Sheets(1).Range("$ar$40").Offset(k, 0) = "1/2"""
'
'            Else
'                Sheets(1).Range("$ar$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$as$40").Offset(k, 0) = "NO"
'            End If
'
'            'transfer BI/BO flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "54") <> 0 Then
'                Sheets(1).Range("$av$40").Offset(k, 0) = "1/2"""
'
'            Else
'                Sheets(1).Range("$av$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$aw$40").Offset(k, 0) = "NO"
'            End If
'
'            'transfer QI/QO flange information
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "62") <> 0 Then
'                Sheets(1).Range("$ap$40").Offset(k, 0) = "1/2"""
'
'            Else
'                Sheets(1).Range("$ap$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$aq$40").Offset(k, 0) = "NO"
'            End If
'
'
'            'transfer CI/CO flange information
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'                Sheets(1).Range("$al$40").Offset(k, 0) = "3/4"""
'
'            Else
'                Sheets(1).Range("$al$40").Offset(k, 0) = "NO"
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "NO"
'            End If
'
'
'        End If
'
'  '/////////////////////////////////////////////
'
''calculate progress course---4
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'    'transfer FI, BI/BO, QI/QO position dimension
'
'    If Sheets(1).Range("$ar$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bs$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 2, False)
'        Sheets(1).Range("$bt$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 3, False)
'    Else
'        Sheets(1).Range("$bs$40").Offset(k, 0) = "N"
'        Sheets(1).Range("$bt$40").Offset(k, 0) = "N"
'    End If
'
'        '=======================
'    If Sheets(1).Range("$av$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bu$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 4, False)
'        Sheets(1).Range("$bv$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 5, False)
'    Else
'        Sheets(1).Range("$bu$40").Offset(k, 0) = "N"
'        Sheets(1).Range("$bv$40").Offset(k, 0) = "N"
'    End If
'
'        '=======================
'
'    If Sheets(1).Range("$ap$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bw$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 6, False)
'        Sheets(1).Range("$bx$40").Offset(k, 0) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 7, False)
'    Else
'        Sheets(1).Range("$bw$40").Offset(k, 0) = "N"
'        Sheets(1).Range("$bx$40").Offset(k, 0) = "N"
'    End If
'
''/////////////////////////////////////////////
'
''calculate progress course---5
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'    'transfer nozzle load information
'
'    If Pump_Type = "ZA" Or Pump_Type = "XA" Then
'
'        'discharge nozzle
'
'        Sheets(1).Range("$ax$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 4, False)
'        Sheets(1).Range("$ay$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 5, False)
'        Sheets(1).Range("$az$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 6, False)
'
'        'suction nozzle
'
'        Sheets(1).Range("$ba$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 14, False)
'        Sheets(1).Range("$bb$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 15, False)
'        Sheets(1).Range("$bc$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 16, False)
'
'        'discharge moment
'        Sheets(1).Range("$bk$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 20, False)
'        Sheets(1).Range("$bl$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 21, False)
'        Sheets(1).Range("$bm$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 22, False)
'
'        'suction moment
'        Sheets(1).Range("$bn$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 20, False)
'        Sheets(1).Range("$bo$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 21, False)
'        Sheets(1).Range("$bp$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 22, False)
'    Else
'
'        Sheets(1).Range("$ax$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 4, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$ay$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 5, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$az$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 6, False) * Sheets("input").Range("$i$56")
'
'        'suction nozzle
'
'        Sheets(1).Range("$ba$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 14, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bb$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 15, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bc$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 16, False) * Sheets("input").Range("$i$56")
'
'        'discharge moment
'        Sheets(1).Range("$bk$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 20, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bl$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 21, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bm$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$s$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 22, False) * Sheets("input").Range("$i$56")
'
'        'suction moment
'        Sheets(1).Range("$bn$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 20, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bo$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 21, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("$bp$40").Offset(k, 0) = Application.HLookup(Sheets(1).Range("$t$40").Offset(k, 0), Sheets("nozzle-load").Range("$b$5:$x$27"), 22, False) * Sheets("input").Range("$i$56")
'
'    End If
' '/////////////////////////////////////////////
'
''calculate progress course---6
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'
'     'transfer PD flange
'        If Sheets("input").Range("$ac$21").Offset(k, 0) = "" Then
'
'                If Sheets("input").Range("$i$53") <> "ASME" Then
'
'                    If Pump_Type <> "ZA" And Pump_Type <> "XA" And Pump_Type <> "XE" Then
'
'                        If Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False) < 80 Then
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "DN15"
'                        Else
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "DN20"
'                        End If
'
'                    Else
'
'                        If Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False) < 150 Then
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "DN15"
'                        Else
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "DN25"
'                        End If
'
'                    End If
'
'
'                Else
'
'                    If Pump_Type <> "ZA" And Pump_Type <> "XA" And Pump_Type <> "XE" Then
'
'                        If Len(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)) = 3 Or VBA.Mid(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False), 1, 1) >= 3 Then
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "3/4"""
'                        Else
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "1/2"""
'                        End If
'
'                    Else
'                        If Len(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)) = 3 Or VBA.Mid(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False), 1, 1) >= 6 Then
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "1"""
'                        Else
'                            Sheets(1).Range("$v$40").Offset(k, 0) = "1/2"""
'                        End If
'
'
'                    End If
'
'                End If
'
'            Else
'                Sheets(1).Range("$v$40").Offset(k, 0) = Sheets("input").Range("$ac$21").Offset(k, 0)
'            End If
' '/////////////////////////////////////////////
'
''calculate progress course---7
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'
'     'transfer nozzle list flange standard
'        If Sheets("input").Range("$i$53") = "ASME" Then
'            'pump discharge and suction nozzle flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$u$40").Offset(k, 0) = "ASME B16.5 150# RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$u$40").Offset(k, 0) = "ASME B16.5 300#  RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$u$40").Offset(k, 0) = "ASME B16.5 600#  RF"
'            'pump PD flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$w$40").Offset(k, 0) = "ASME B16.5 150#  RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$w$40").Offset(k, 0) = "ASME B16.5 300#  RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$w$40").Offset(k, 0) = "ASME B16.5 600#  RF"
'
'            'transfer CI/CO flange standard
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "ASME B16.5 150#  RF"
'
'            Else
'            End If
'
'        ElseIf Sheets("input").Range("$i$53") = "HG20615" Then
'            'pump discharge and suction nozzle flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20615 PN2.0 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20615 PN5.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20615 PN11.0 RF"
'            'pump PD flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20615 PN2.0 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20615 PN5.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20615 PN11.0 RF"
'
'            'transfer CI/CO flange standard
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "HG20615 PN2.0 RF"
'
'            Else
'            End If
'
'        ElseIf Sheets("input").Range("$i$53") = "HG20592" Then
'            'pump discharge and suction nozzle flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20592 PN2.5 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20592 PN4.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$u$40").Offset(k, 0) = "HG20615 PN11.0 RF"
'            'pump PD flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20592 PN2.5 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20592 PN4.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$w$40").Offset(k, 0) = "HG20615 PN11.0 RF"
'
'            'transfer CI/CO flange standard
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "HG20592 PN2.5 RF"
'
'            Else
'            End If
'
'        ElseIf Sheets("input").Range("$i$53") = "SH3406" Then
'            'pump discharge and suction nozzle flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$u$40").Offset(k, 0) = "SH3406 PN2.0 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$u$40").Offset(k, 0) = "SH3406 PN5.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$u$40").Offset(k, 0) = "SH3406 PN10.0 RF"
'            'pump PD flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$w$40").Offset(k, 0) = "SH3406 PN2.0 RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$w$40").Offset(k, 0) = "SH3406 PN5.0 RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$w$40").Offset(k, 0) = "SH3406 PN10.0 RF"
'
'            'transfer CI/CO flange standard
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "SH3406 PN2.0 RF"
'
'            Else
'            End If
'
'        Else
'            'pump discharge and suction nozzle flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$u$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$u$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$13") & " RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$u$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$14") & " RF"
'            'pump PD flange standard
'            If Pump_Type = "ZA" Or Pump_Type = "XA" Then Sheets(1).Range("$w$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"
'            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("$w$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$13") & " RF"
'            If Pump_Type = "ZF" Then Sheets(1).Range("$w$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$14") & " RF"
'
'            'transfer CI/CO flange standard
'            If Sheets("input").Range("$z$21").Offset(k, 0) <> "" And Sheets("input").Range("$z$21").Offset(k, 0) <> "无" Then
'
'                Sheets(1).Range("$ao$40").Offset(k, 0) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"
'
'            Else
'            End If
'
'        End If
'
'
' '/////////////////////////////////////////////
'
''calculate progress course---8
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'            'FI flange standard
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "32") <> 0 Then
'
'                Sheets(1).Range("$as$40").Offset(k, 0) = Sheets(1).Range("$u$40").Offset(k, 0)
'
'            Else
'            End If
'            'BI/BO flange standard
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "54") <> 0 Then
'
'                Sheets(1).Range("$aw$40").Offset(k, 0) = Sheets(1).Range("$u$40").Offset(k, 0)
'
'            Else
'            End If
'            'QI/QO flange standard
'            If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "62") <> 0 Then
'
'                Sheets(1).Range("$aq$40").Offset(k, 0) = Sheets(1).Range("$u$40").Offset(k, 0)
'
'            Else
'            End If
'
''/////////////////////////////////////////////
'
''calculate progress course---9
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'    'interface name list
'
'    'there are 7 kinds of interface name totally as follows,
'        'NO+NO+NO+NO
'        'FI+NO+QI/QO+CI/CO
'        'NO+NO+NO+CI/CO
'        'FI+NO+NO+CI/CO
'        'NO+BI/BO+NO+NO
'        'FI+NO+NO+NO
'        'NO+NO+QI/QO+NO
'
'    'Mr Liu Fengqing will make list block accordingly
'
'
'    If Sheets(1).Range("$ar$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bd$40").Offset(k, 0) = "F"
'    Else
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$ar$40").Offset(k, 0)
'    End If
'
'    If Sheets(1).Range("$av$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & "B"
'    Else
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & Sheets(1).Range("$av$40").Offset(k, 0)
'    End If
'
'    If Sheets(1).Range("$ap$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & "Q"
'    Else
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & Sheets(1).Range("$ap$40").Offset(k, 0)
'    End If
'
'    If Sheets(1).Range("$al$40").Offset(k, 0) <> "NO" Then
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & "C"
'    Else
'        Sheets(1).Range("$bd$40").Offset(k, 0) = Sheets(1).Range("$bd$40").Offset(k, 0) & Sheets(1).Range("$al$40").Offset(k, 0)
'    End If
'
''/////////////////////////////////////////////
'
''calculate progress course---10
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'   'judge if there's heat exchanger or reservoir existing
'
'    If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "21") <> 0 Then
'        Sheets(1).Range("$be$40").Offset(k, 0) = "Y"
'        Sheets(1).Range("$bi$40").Offset(k, 0) = "N"
'    ElseIf InStr(Sheets("input").Range("$y$21").Offset(k, 0), "23") <> 0 Then
'        Sheets(1).Range("$be$40").Offset(k, 0) = "Y"
'    'if flush plan is Plan23, write "L" for CAD programming
'        Sheets(1).Range("$bi$40").Offset(k, 0) = "L"
'    ElseIf InStr(Sheets("input").Range("$y$21").Offset(k, 0), "41") <> 0 Then
'        Sheets(1).Range("$be$40").Offset(k, 0) = "Y"
'        Sheets(1).Range("$bi$40").Offset(k, 0) = "N"
'    Else
'        Sheets(1).Range("$be$40").Offset(k, 0) = "N"
'        Sheets(1).Range("$bi$40").Offset(k, 0) = "N"
'    End If
'
'     If InStr(Sheets("input").Range("$y$21").Offset(k, 0), "52") <> 0 Then
'        Sheets(1).Range("$bf$40").Offset(k, 0) = "Y"
'    ElseIf InStr(Sheets("input").Range("$y$21").Offset(k, 0), "53") <> 0 Then
'        Sheets(1).Range("$bf$40").Offset(k, 0) = "Y"
'
'    Else
'        Sheets(1).Range("$bf$40").Offset(k, 0) = "N"
'    End If
'
'
'
''/////////////////////////////////////////////
'
''calculate progress course---11
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'
'
'    'transfer GA, PID, SEC file name
'        If Sheets(1).Range("$b$40").Offset(k, 0) <> "" Then
'            str3 = Sheets(1).Range("$b$40").Offset(k, 0) & "_" & str2 & "_" & Sheets(1).Range("$e$40").Offset(k, 0) & "_" & Sheets("input").Range("$u$21").Offset(k, 0) & "_Ver" & Sheets(1).Range("$c$3")
'        Else
'            str3 = str2 & "_" & Sheets(1).Range("$e$40").Offset(k, 0) & "_" & Sheets("input").Range("$u$21").Offset(k, 0) & "_Ver" & Sheets(1).Range("$c$3")
'        End If
'
'          'filter illegal character like "/ \ , space" etc
'
'            j = Len(str3)
'            name = ""
'                For x = 1 To j
'
'                    str1 = VBA.Mid(str3, x, 1)
'
'
'                    If str1 <> "/" And str1 <> "\" And str1 <> " " And str1 <> ")" And str1 <> "(" And str1 <> "," And str1 <> "，" Then
'                          name = name & str1
'                    Else
'                         str1 = ""
'                          name = name & str1
'                    End If
'                Sheets(1).Range("$p$40").Offset(k, 0) = "GA_" & name
'                Sheets(1).Range("$q$40").Offset(k, 0) = "PID_" & name
'                Sheets(1).Range("$r$40").Offset(k, 0) = "Sec&List_" & name
'                Next x
'
'      'calculate available rows in ZE,ZF,ZA data source
'      Counter = Sheets(Pump_Type).Range("A63655").End(xlUp).Row - 31
'
'
''/////////////////////////////////////////////
'
''calculate progress course---12
'
'    Progress = 5 * CounterProgress
'      CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'        For x = 0 To Counter - 1
'
'      Counter2 = 0
'      other_weight = 20
'            If Sheets("input").Range("$v$21").Offset(k, 0) <> "380V" Then
'
'                If Sheets(Pump_Type).Range("$a$32").Offset(x, 0) = str4 And Sheets(Pump_Type).Range("$a$32").Offset(x, 5) = Sheets("input").Range("$u$21").Offset(k, 0) & "(6KV)" Then
'                    'transfer motor weight
'                    Sheets(1).Range("$g$40").Offset(k, 0) = Sheets(Pump_Type).Range("y" & x + 32)
'                    'transfer baseplate weight
'                    Sheets(1).Range("$h$40").Offset(k, 0) = Sheets(Pump_Type).Range("z" & x + 32)
'                    'transfer others weight
'
'                    For y = 0 To 6
'
'                        str5 = Sheets("Data1").Range("$a$2").Offset(y, 0)
'                        If InStr(Sheets(1).Range("$e$40").Offset(k, 0), str5) <> 0 Then
'                            Counter2 = Counter2 + 1
'
'                        Else
'                        End If
'
'
'
'                    Next y
'                    Sheets(1).Range("$i$40").Offset(k, 0) = other_weight + 100 * Counter2
'                    'transfer total weight
'                    'total weight will be got by function
'
'                    'Sheets(1).Range("$j$40").Offset(k, 0) = Sheets(1).Range("$f$40").Offset(k, 0) + Sheets(1).Range("$g$40").Offset(k, 0) + Sheets(1).Range("$h$40").Offset(k, 0) + Sheets(1).Range("$i$40").Offset(k, 0)
'
'
'                    'transfer H value
'                    Sheets(1).Range("$x$40").Offset(k, 0) = Sheets(Pump_Type).Range("g" & x + 32)
'                    'transfer L value
'                    Sheets(1).Range("$y$40").Offset(k, 0) = Sheets(Pump_Type).Range("h" & x + 32)
'                    'transfer L1 value
'                    Sheets(1).Range("$z$40").Offset(k, 0) = Sheets(Pump_Type).Range("i" & x + 32)
'                     'transfer L2 value
'                    Sheets(1).Range("$aa$40").Offset(k, 0) = Sheets(Pump_Type).Range("j" & x + 32)
'                     'transfer L3 value
'                    Sheets(1).Range("$ab$40").Offset(k, 0) = Sheets(Pump_Type).Range("k" & x + 32)
'                      'transfer k value
'                    If Sheets("input").Range("$ab$21").Offset(k, 0) = "" Then
'
'                        Sheets(1).Range("$ac$40").Offset(k, 0) = Sheets(Pump_Type).Range("l" & x + 32)
'                    Else
'                        Sheets(1).Range("$ac$40").Offset(k, 0) = Sheets("input").Range("$ab$21").Offset(k, 0)
'                    End If
'
'                      'transfer B value
'                    Sheets(1).Range("$ad$40").Offset(k, 0) = Sheets(Pump_Type).Range("m" & x + 32)
'                      'transfer B1 value
'                    Sheets(1).Range("$ae$40").Offset(k, 0) = Sheets(Pump_Type).Range("n" & x + 32)
'                      'transfer i value
'                    Sheets(1).Range("$af$40").Offset(k, 0) = Sheets(Pump_Type).Range("o" & x + 32)
'                      'transfer h value
'                    Sheets(1).Range("$ag$40").Offset(k, 0) = Sheets(Pump_Type).Range("p" & x + 32)
'                      'transfer h1 value
'                    Sheets(1).Range("$ah$40").Offset(k, 0) = Sheets(Pump_Type).Range("q" & x + 32)
'                      'transfer h2 value
'                    Sheets(1).Range("$ai$40").Offset(k, 0) = Sheets(Pump_Type).Range("r" & x + 32)
'                      'transfer e value
'                    Sheets(1).Range("$aj$40").Offset(k, 0) = Sheets(Pump_Type).Range("s" & x + 32)
'                      'transfer y value
'                    Sheets(1).Range("$am$40").Offset(k, 0) = Sheets(Pump_Type).Range("v" & x + 32)
'                       'transfer drain outlet height value
'                    Sheets(1).Range("$an$40").Offset(k, 0) = Sheets(Pump_Type).Range("w" & x + 32)
'                        'transfer anchor bolt value
'                    Sheets(1).Range("$at$40").Offset(k, 0) = Sheets(Pump_Type).Range("ac" & x + 32)
'                        'transfer anchor bolt amount value
'                    Sheets(1).Range("$au$40").Offset(k, 0) = Sheets(Pump_Type).Range("ad" & x + 32)
'                        'transfer L4 value---only for ZA pump
'                    If Pump_Type = "ZA" Then
'                        Sheets(1).Range("$bq$40").Offset(k, 0) = Sheets(Pump_Type).Range("ae" & x + 32)
'                    Else
'                        Sheets(1).Range("$bq$40").Offset(k, 0) = "N"
'                    End If
'
'                Else
'
'                End If
'
'            Else
'
'                 If Sheets(Pump_Type).Range("$a$32").Offset(x, 0) = str4 And Sheets(Pump_Type).Range("$a$32").Offset(x, 5) = Sheets("input").Range("$u$21").Offset(k, 0) Then
'                    Sheets(1).Range("$g$40").Offset(k, 0) = Sheets(Pump_Type).Range("y" & x + 32)
'                    Sheets(1).Range("$h$40").Offset(k, 0) = Sheets(Pump_Type).Range("z" & x + 32)
'
'
'                    For y = 0 To 6
'
'                        str5 = Sheets("Data1").Range("$a$2").Offset(y, 0)
'                        If InStr(Sheets(1).Range("$e$40").Offset(k, 0), str5) <> 0 Then
'                            Counter2 = Counter2 + 1
'
'                        Else
'                        End If
'
'
'
'                    Next y
'                    Sheets(1).Range("$i$40").Offset(k, 0) = other_weight + 100 * Counter2
'
'                    'Sheets(1).Range("$j$40").Offset(k, 0) = Sheets(1).Range("$f$40").Offset(k, 0) + Sheets(1).Range("$g$40").Offset(k, 0) + Sheets(1).Range("$h$40").Offset(k, 0) + Sheets(1).Range("$i$40").Offset(k, 0)
'
'                    Sheets(1).Range("$x$40").Offset(k, 0) = Sheets(Pump_Type).Range("g" & x + 32)
'                    Sheets(1).Range("$y$40").Offset(k, 0) = Sheets(Pump_Type).Range("h" & x + 32)
'                    Sheets(1).Range("$z$40").Offset(k, 0) = Sheets(Pump_Type).Range("i" & x + 32)
'                    Sheets(1).Range("$aa$40").Offset(k, 0) = Sheets(Pump_Type).Range("j" & x + 32)
'                    Sheets(1).Range("$ab$40").Offset(k, 0) = Sheets(Pump_Type).Range("k" & x + 32)
'
'                    If Sheets("input").Range("$ab$21").Offset(k, 0) = "" Then
'
'                        Sheets(1).Range("$ac$40").Offset(k, 0) = Sheets(Pump_Type).Range("l" & x + 32)
'                    Else
'                        Sheets(1).Range("$ac$40").Offset(k, 0) = Sheets("input").Range("$ab$21").Offset(k, 0)
'                    End If
'
'
'                    Sheets(1).Range("$ad$40").Offset(k, 0) = Sheets(Pump_Type).Range("m" & x + 32)
'                    Sheets(1).Range("$ae$40").Offset(k, 0) = Sheets(Pump_Type).Range("n" & x + 32)
'                    Sheets(1).Range("$af$40").Offset(k, 0) = Sheets(Pump_Type).Range("o" & x + 32)
'                    Sheets(1).Range("$ag$40").Offset(k, 0) = Sheets(Pump_Type).Range("p" & x + 32)
'                    Sheets(1).Range("$ah$40").Offset(k, 0) = Sheets(Pump_Type).Range("q" & x + 32)
'                    Sheets(1).Range("$ai$40").Offset(k, 0) = Sheets(Pump_Type).Range("r" & x + 32)
'                    Sheets(1).Range("$aj$40").Offset(k, 0) = Sheets(Pump_Type).Range("s" & x + 32)
'                    Sheets(1).Range("$am$40").Offset(k, 0) = Sheets(Pump_Type).Range("v" & x + 32)
'                    Sheets(1).Range("$an$40").Offset(k, 0) = Sheets(Pump_Type).Range("w" & x + 32)
'                    Sheets(1).Range("$at$40").Offset(k, 0) = Sheets(Pump_Type).Range("ac" & x + 32)
'                    Sheets(1).Range("$au$40").Offset(k, 0) = Sheets(Pump_Type).Range("ad" & x + 32)
'                    If Pump_Type = "ZA" Then
'                        Sheets(1).Range("$bq$40").Offset(k, 0) = Sheets(Pump_Type).Range("ae" & x + 32)
'                    Else
'                        Sheets(1).Range("$bq$40").Offset(k, 0) = "N"
'                    End If
'
'
'                Else
'
'                End If
'
'
'
'            End If
'
'
'            'transfer f or f* value, for ZF and ZA, use f value, for ZE with fan cooling feature, use f* value
'
'            If Pump_Type <> "ZE" Then Sheets(1).Range("$ak$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 20, False)
'
'            If Pump_Type = "ZE" Then
'                If InStr(Sheets("input").Range("$e$21").Offset(k, 0), "8") <> 0 And InStr(Sheets("input").Range("$aa$21").Offset(k, 0), "F") Then
'                    Sheets(1).Range("$ak$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 21, False)
'
'
'                ElseIf InStr(Sheets("input").Range("$e$21").Offset(k, 0), "0th") <> 0 Then
'                    Sheets(1).Range("$ak$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 21, False)
'
'                Else
'                    Sheets(1).Range("$ak$40").Offset(k, 0) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 20, False)
'                End If
'            End If
'
'
'
'            'calculate heat exchanger and reservior position dimensions
'
'            If Sheets(1).Range("$be$40").Offset(k, 0) <> "N" Then
'
'                Sheets(1).Range("$bg$40").Offset(k, 0) = Sheets(1).Range("$af$40").Offset(k, 0) + Sheets(1).Range("$ak$40").Offset(k, 0) - 200
'
'
'                If VBA.Mid(Sheets(1).Range("$bg$40").Offset(k, 0), Len(Sheets(1).Range("$bg$40").Offset(k, 0)), 1) <> 5 Then Sheets(1).Range("$bg$40").Offset(k, 0) = VBA.Mid(Sheets(1).Range("$bg$40").Offset(k, 0), 1, Len(Sheets(1).Range("$bg$40").Offset(k, 0)) - 1) & "0"
'
'            Else
'                Sheets(1).Range("$bg$40").Offset(k, 0) = "N"
'            End If
'
'            If Sheets(1).Range("$bf$40").Offset(k, 0) <> "N" Then
'                Sheets(1).Range("$bh$40").Offset(k, 0) = Sheets(1).Range("$af$40").Offset(k, 0) + Sheets(1).Range("$ak$40").Offset(k, 0) - 100
'
'                If VBA.Mid(Sheets(1).Range("$bh$40").Offset(k, 0), Len(Sheets(1).Range("$bh$40").Offset(k, 0)), 1) <> 5 Then Sheets(1).Range("$bh$40").Offset(k, 0) = VBA.Mid(Sheets(1).Range("$bh$40").Offset(k, 0), 1, Len(Sheets(1).Range("$bh$40").Offset(k, 0)) - 1) & "0"
'
'
'
'            Else
'
'                Sheets(1).Range("$bh$40").Offset(k, 0) = "N"
'
'            End If
'
'            'calculate AD2 value
'            If Pump_Type = "ZA" Then
'
'                Sheets(1).Range("$bj$40").Offset(k, 0) = "1"""
'           Else
'
'                Sheets(1).Range("$bj$40").Offset(k, 0) = "2"""
'
'           End If
'
'
'
'
'
'
'
'        Next x
''/////////////////////////////////////////////
'
''calculate progress course---13
'
'    Progress = 5 * CounterProgress
'    CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'        'fill "是否为ZE+G（Y or N）" information
'          If InStr(Sheets(1).Range("$c$40").Offset(k, 0), "G") <> 0 Or InStr(Sheets(1).Range("$c$40").Offset(k, 0), "g") <> 0 Then
'            Sheets(1).Range("$by$40").Offset(k, 0) = "Y"
'          Else
'            Sheets(1).Range("$by$40").Offset(k, 0) = "N"
'          End If
'        'fill "轴承架号是否>=LK6（Y or N）" information
'           j = Len(Sheets(1).Range("$c$40").Offset(k, 0))
'           If VBA.Mid(Sheets(1).Range("$c$40").Offset(k, 0), j - 3, 1) >= 6 Then
'             Sheets(1).Range("$bz$40").Offset(k, 0) = "Y"
'           Else
'             Sheets(1).Range("$bz$40").Offset(k, 0) = "N"
'           End If
'         'fill Cooling method
'           Sheets(1).Range("$cb$40").Offset(k, 0) = Sheets("input").Range("$aa$21").Offset(k, 0)
'         'fill section block name
'            j = Len(Sheets("input").Range("$e$21").Offset(k, 0))
'
'            If Sheets(1).Range("$by$40").Offset(k, 0) = "Y" Then
'
'                If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) = 8 Then
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type & "8G"
'                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) = 0 Then
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type & "10G"
'                Else
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type & "G"
'                End If
'
'            ElseIf Sheets(1).Range("$by$40").Offset(k, 0) = "N" Then
'                 If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) = 8 Then
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type & "8"
'                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) = 0 Then
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type & "10"
'                Else
'                    Sheets(1).Range("$ca$40").Offset(k, 0) = Pump_Type
'                End If
'
'            End If
'           'fill section list name
'            j = Len(Sheets("input").Range("$e$21").Offset(k, 0))
'
'        If Pump_Type = "ZE" Then
'
'            If Sheets(1).Range("$by$40").Offset(k, 0) = "Y" Then
'
'                If Sheets(1).Range("$bz$40").Offset(k, 0) = "N" Then
'
'                    If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) <> 0 Then
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "8G" & "_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'                    Else
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "10" & "_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                    End If
'                Else
'                    If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) <> 0 Then
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "8G" & "(LK6)_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'                    Else
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "10" & "(LK6)_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                    End If
'
'                End If
'
'            Else
'                If Sheets(1).Range("$bz$40").Offset(k, 0) = "N" Then
'
'                    If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) <> 0 Then
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "8" & "_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'                    Else
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "10" & "_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                    End If
'                Else
'                    If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("$e$21").Offset(k, 0), " ", ""), 3), 1) <> 0 Then
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "8" & "(LK6)_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'                    Else
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type & "10" & "(LK6)_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                    End If
'
'                End If
'
'            End If
'
'        Else
'                If Sheets(1).Range("$bz$40").Offset(k, 0) = "N" Then
'
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type2 & "_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                Else
'                        Sheets(1).Range("$cc$40").Offset(k, 0) = Pump_Type2 & "(LK6)_" & Sheets(1).Range("$cb$40").Offset(k, 0)
'
'                End If
'
'        End If
'
'
'
''/////////////////////////////////////////////
'
''calculate progress course---14
'
'    Progress = 5 * CounterProgress
'    CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
''/////////////////////////////////////////////
'
''calculate progress course---15
'
'    Progress = 5 * CounterProgress
'    CounterProgress = CounterProgress + 1
'    PctDone = Progress / CounterSum
'        Call UpdateProgress(PctDone)
'
''//////////////////////////////////////////////
'
'    Next k
'
'  Unload UserForm2
'
''=======================================
''check value if there's space
'
'    CounterCheck = 0
'    Sheets(1).Range("$b$40:$cc$" & (39 + i)).Select
'
'    If TypeName(Selection) <> "Range" Then Exit Sub
'
'    Application.ScreenUpdating = False
'    For Each cell In Selection
'        If cell.Value = "" Then
'          cell.Interior.ColorIndex = 3
'          CounterCheck = CounterCheck + 1
'
'        Else
'          cell.Interior.ColorIndex = xlNone
'        End If
'    Next cell
'
'    If CounterCheck <> 0 Then
'        MsgBox ("计算结果中发现空值，请确认或修正，谢谢。")
'        Exit Sub
'    End If
'
'MsgBox ("计算完成，如需特殊输入，请在本页修正，谢谢。")
'
''===================
'
''Application.Visible = True
'
''====================
'
'End Sub
'
''//////////////////////////////////////////////////////////////////////////////////////////
'
''used for progress
'
'Sub UpdateProgress(Pct)
'    With UserForm2
'        .FrameProgress.Caption = Format(Pct, "0%")
'        .LabelProgress.Width = Pct * (.FrameProgress.Width - 10)
'        .Repaint
'    End With
'End Sub
'
''/////////////////////////////////////////////////////////////////////////////////////////////////
'
'
