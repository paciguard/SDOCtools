Attribute VB_Name = "Module1"
'Button
'=========================================
Option Explicit
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
() '  Alias "ShellExecuteA" (ByVal hwnd As Long, _
  ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long


Sub tranfer_data()

    Dim Pump_Type As Variant, Pump_Type2 As Variant, Pump_Type_all As Variant
    Dim i As Long, j As Long, k As Long, X As Long, z As Long, Counter As Long, Counter1 As Long, Counter2 As Long, Y As Long
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str2new As String, LK As String, name As String, NorDia As String
    Dim other_weight As Long
    Dim CounterProgress As Long, Progress As Long, CounterSum As Long
    Dim PctDone As Single
    Dim xCheck As Long, yCheck As Long
    Dim cell As Range
    Dim CounterCheck As Long
    Dim abcd As String
    Dim origpumptype, newpumptype As String

    'save information
    ThisWorkbook.Save

    'count available rows
    i = Application.CountA(Sheets("input").Range("$n$21:$n$50"))

    Sheets(1).Range("$c$15") = i

    'check mandatory filling in sheet "input"
    For k = 0 To i - 1

       '======================
        If Sheets("input").Range("$i$53") = "GB9113.1-2000" Then
            For Counter1 = 0 To 2

                If InStr(Sheets(1).Range("C" & 12 + Counter1), "P") = 0 Then

                    MsgBox ("请将GB标准的法兰信息填写完整")
                    Exit Sub
                Else
                End If

            Next Counter1
        Else

        End If
       '======================
        If Sheets("input").Range("U" & 21 + k) = "" Then
            MsgBox ("基座号列信息不得为空")
            Exit Sub
        Else

        End If
        '======================
        If Sheets("input").Range("V" & 21 + k) = "" Then
            MsgBox ("电机电压列信息不得为空")
            Exit Sub
        Else

        End If
        '======================

        If Sheets("input").Range("V" & 21 + k) = "6000V" Or Sheets("input").Range("V" & 21 + k) = "10000V" Then
            MsgBox ("请注意进行电机长度及底座长度校核")

        Else

        End If
        '======================
         If Sheets("input").Range("AA" & 21 + k) = "" Then
            MsgBox ("轴承形式列信息不得为空")
            Exit Sub
        Else

        End If

    Next k


Sheets(1).Activate
'===================

'Application.Visible = False

'====================
     'clear former contents

    Sheets(1).Range("$b$40:$i$69").ClearContents
    Sheets(1).Range("$k$40:$cc$69").ClearContents
    Sheets("Sec-List").Range("$f$10:$f$39").ClearContents

'/////////////////////////////////////////////

'calculate cells amount for progress

CounterSum = 75 * i

CounterProgress = 1


UserForm2.Show 0



'//////////////////////////////////////////////
'===================================================================fill data



    For k = 0 To i - 1

    'link item no.
    Sheets(1).Range("B" & 40 + k) = Sheets("input").Range("D" & 21 + k)
    'link service name
    Sheets(1).Range("D" & 40 + k) = Sheets("input").Range("F" & 21 + k)
    'link motor
    Sheets(1).Range("K" & 40 + k) = Sheets("input").Range("S" & 21 + k)
    'link motor manufacture
    Sheets(1).Range("N" & 40 + k) = Sheets("input").Range("T" & 21 + k)
    'link motor manufacture
    Sheets(1).Range("N" & 40 + k) = Sheets("input").Range("T" & 21 + k)
    'link motor voltage
        j = Len(Sheets("input").Range("V" & 21 + k))

    Sheets(1).Range("L" & 40 + k) = VBA.Mid(Sheets("input").Range("V" & 21 + k), 1, j - 1)
    'link motor power
    Sheets(1).Range("M" & 40 + k) = Sheets("input").Range("W" & 21 + k)
    'link speed
    Sheets(1).Range("O" & 40 + k) = Sheets("input").Range("R" & 21 + k)
    'link Material class to sheet "Sec-List"
    Sheets("Sec-List").Range("F" & 10 + k) = Sheets("input").Range("O" & 21 + k)


    'fill and rewrite flush plan
        If Sheets("input").Range("Z" & 21 + k) = "" Or Sheets("input").Range("Z" & 21 + k) = "无" Or Sheets("input").Range("Z" & 21 + k) = "NO" Then



            If InStr(Sheets("input").Range("Y" & 21 + k), "P") <> 0 Or InStr(Sheets("input").Range("Y" & 21 + k), "p") <> 0 Then
                Sheets(1).Range("E" & 40 + k) = Sheets("input").Range("Y" & 21 + k)
            Else
                Sheets(1).Range("E" & 40 + k) = "Plan " & Sheets("input").Range("Y" & 21 + k)
            End If

        Else

            If InStr(Sheets("input").Range("Y" & 21 + k), "P") <> 0 Or InStr(Sheets("input").Range("Y" & 21 + k), "p") <> 0 Then
                Sheets(1).Range("E" & 40 + k) = Sheets("input").Range("Y" & 21 + k) & "+" & Sheets("input").Range("Z" & 21 + k)
            Else
                Sheets(1).Range("E" & 40 + k) = "Plan " & Sheets("input").Range("Y" & 21 + k) & "+" & Sheets("input").Range("Z" & 21 + k)
            End If

        End If


'/////////////////////////////////////////////

'calculate progress course---1

    Progress = 5 * CounterProgress

    CounterProgress = CounterProgress + 1


    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////



'=======================================================
        Pump_Type_all = Sheets("input").Range("N" & 21 + k)

    'filter "space" and assignment pump type to sheet1
            j = Len(Pump_Type_all)
            str2 = ""
                For X = 1 To j

                    str1 = VBA.Mid(Pump_Type_all, X, 1)
                    If Asc(str1) = 32 Or Asc(str1) = 63 Then
                    str1 = ""
                    End If

                    If str1 <> " " Then
                          str2 = str2 & str1
                    Else
                         str1 = ""
                         str2 = str2 & str1
                    End If
                Sheets(1).Range("C" & 40 + k) = str2

                Next X
                
                If Sheets("input").Range("AB" & 21 + k) <> "" Then
                str2new = VBA.Replace(Sheets("input").Range("AB" & 21 + k), VBA.Left(str2, 2), Sheets("input").Range("AB" & 21 + k))
                End If

'/////////////////////////////////////////////

'calculate progress course---2

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

    ' str2 is the right format pump type for search
    ' delete character "O" since the dimensions are the same between "ZA" and "ZAO", as well as "ZE" and "ZEO"

        If VBA.Mid(str2, 3, 1) = "O" Or VBA.Mid(str2, 3, 1) = "o" Or VBA.Mid(str2, 3, 1) = "G" Or VBA.Mid(str2, 3, 1) = "g" Then
            str4 = VBA.Mid(str2, 1, 2) & VBA.Mid(str2, 4, Len(str2) - 3)
            Pump_Type2 = VBA.Mid(str2, 1, 3)
        Else
            str4 = str2
            Pump_Type2 = VBA.Mid(str2, 1, 2)
        End If
        Pump_Type = VBA.Mid(str2, 1, 2)
        NorDia = VBA.Right(str2, 3)
    'str4 is a processing value for search which has been filtered "O" character

    'transfer ZA drain block name---str6 is search key
        If Pump_Type = "ZA" Or Pump_Type = "XA" Then
            z = Len(str4)
                For X = 1 To z

                    If VBA.Mid(str4, X, 1) = "-" Then Exit For

                Next X
            str6 = VBA.Mid(str4, 3, X - 2) & VBA.Mid(str4, X + 2, 3)

            Sheets(1).Range("BR" & 40 + k) = Application.VLookup(str6, Sheets("Data1").Range("$c$10:$d$62"), 2, False)
        Else
            Sheets(1).Range("BR" & 40 + k) = "N"
        End If

    'calculate LK value
        If VBA.Left(str4, 2) = "CZ" Then
        LK = Application.VLookup(str4, Sheets("CZ-BearingSize").Range("$A$2:$B$41"), 2, False)
        Else
        z = Len(str4)
            For X = 1 To z

                If VBA.Mid(str4, X, 1) = "-" Then Exit For

            Next X
        LK = "LK" & VBA.Mid(str4, X + 1, 1)
        End If
        



    'transfer pump weight
    
If Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
      origpumptype = Pump_Type
      Pump_Type = "ZE7"
End If
    
    
        Sheets(1).Range("F" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 24, False)



'/////////////////////////////////////////////

'calculate progress course---3

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

    'transfer discharge and suction nozzle flange information
        If Sheets("input").Range("$i$53") <> "ASME" Then
            'transfer discharge and suction nozzle flange information
            Sheets(1).Range("S" & 40 + k) = "DN" & Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False)
            Sheets(1).Range("T" & 40 + k) = "DN" & Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 3, False)
            'transfer FI flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "32") <> 0 Then
                Sheets(1).Range("AR" & 40 + k) = "DN15"

            Else
                Sheets(1).Range("AR" & 40 + k) = "NO"
                Sheets(1).Range("AS" & 40 + k) = "NO"
            End If

            'transfer BI/BO flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "54") <> 0 Then
                Sheets(1).Range("AV" & 40 + k) = "DN15"

            Else
                Sheets(1).Range("AV" & 40 + k) = "NO"
                Sheets(1).Range("AW" & 40 + k) = "NO"
            End If

            'transfer QI/QO flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "62") <> 0 Then
                Sheets(1).Range("AP" & 40 + k) = "DN15"

            Else
                Sheets(1).Range("AP" & 40 + k) = "NO"
                Sheets(1).Range("AQ" & 40 + k) = "NO"
            End If

            'transfer CI/CO flange information
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then
                Sheets(1).Range("AL" & 40 + k) = "DN20"

            Else
                Sheets(1).Range("AL" & 40 + k) = "NO"
                Sheets(1).Range("AO" & 40 + k) = "NO"
            End If

        Else
            Sheets(1).Range("S" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)
            Sheets(1).Range("T" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 5, False)
            'transfer FI flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "32") <> 0 Then
                Sheets(1).Range("AR" & 40 + k) = "1/2"""

            Else
                Sheets(1).Range("AR" & 40 + k) = "NO"
                Sheets(1).Range("AS" & 40 + k) = "NO"
            End If

            'transfer BI/BO flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "54") <> 0 Then
                Sheets(1).Range("AV" & 40 + k) = "1/2"""

            Else
                Sheets(1).Range("AV" & 40 + k) = "NO"
                Sheets(1).Range("AW" & 40 + k) = "NO"
            End If

            'transfer QI/QO flange information
            If InStr(Sheets("input").Range("Y" & 21 + k), "62") <> 0 Then
                Sheets(1).Range("AP" & 40 + k) = "1/2"""

            Else
                Sheets(1).Range("AP" & 40 + k) = "NO"
                Sheets(1).Range("AQ" & 40 + k) = "NO"
            End If


            'transfer CI/CO flange information
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then
                Sheets(1).Range("AL" & 40 + k) = "3/4"""

            Else
                Sheets(1).Range("AL" & 40 + k) = "NO"
                Sheets(1).Range("AO" & 40 + k) = "NO"
            End If


        End If

  '/////////////////////////////////////////////
  If origpumptype <> "" Then Pump_Type = origpumptype
'calculate progress course---4

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

    'transfer FI, BI/BO, QI/QO position dimension

    If Sheets(1).Range("AR" & 40 + k) <> "NO" Then
        Sheets(1).Range("BS" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 2, False)
        Sheets(1).Range("BT" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 3, False)
    Else
        Sheets(1).Range("BS" & 40 + k) = "N"
        Sheets(1).Range("BT" & 40 + k) = "N"
    End If

        '=======================
    If Sheets(1).Range("AV" & 40 + k) <> "NO" Then
        Sheets(1).Range("BU" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 4, False)
        Sheets(1).Range("BV" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 5, False)
    Else
        Sheets(1).Range("BU" & 40 + k) = "N"
        Sheets(1).Range("BV" & 40 + k) = "N"
    End If

        '=======================

    If Sheets(1).Range("AP" & 40 + k) <> "NO" Then
        Sheets(1).Range("BW" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 6, False)
        Sheets(1).Range("BX" & 40 + k) = Application.HLookup(LK, Sheets("Data1").Range("$g$65:$o$71"), 7, False)
    Else
        Sheets(1).Range("BW" & 40 + k) = "N"
        Sheets(1).Range("BX" & 40 + k) = "N"
    End If

'/////////////////////////////////////////////

'calculate progress course---5

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////


    'transfer nozzle load information

    If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then

        'discharge nozzle

        Sheets(1).Range("AX" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 4, False)
        Sheets(1).Range("AY" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 5, False)
        Sheets(1).Range("AZ" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 6, False)

        'suction nozzle

        Sheets(1).Range("BA" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 14, False)
        Sheets(1).Range("BB" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 15, False)
        Sheets(1).Range("BC" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 16, False)

        'discharge moment
        Sheets(1).Range("BK" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False)
        Sheets(1).Range("BL" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False)
        Sheets(1).Range("BM" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False)

        'suction moment
        Sheets(1).Range("BN" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False)
        Sheets(1).Range("BO" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False)
        Sheets(1).Range("BP" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False)
    'ElseIf Pump_Type = "ZF" Then
Else
        Sheets(1).Range("AX" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 4, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("AY" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 5, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("AZ" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 6, False) * Sheets("input").Range("$i$56")

        'suction nozzle

        Sheets(1).Range("BA" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 14, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BB" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 15, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BC" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 16, False) * Sheets("input").Range("$i$56")

        'discharge moment
        Sheets(1).Range("BK" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BL" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BM" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False) * Sheets("input").Range("$i$56")

        'suction moment
        Sheets(1).Range("BN" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BO" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False) * Sheets("input").Range("$i$56")
        Sheets(1).Range("BP" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False) * Sheets("input").Range("$i$56")
'   ElseIf Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
'   'discharge nozzle
'
'        Sheets(1).Range("AX" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 4, False)
'        Sheets(1).Range("AY" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 5, False)
'        Sheets(1).Range("AZ" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 6, False)
'
'        'suction nozzle
'
'        Sheets(1).Range("BA" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 14, False)
'        Sheets(1).Range("BB" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 15, False)
'        Sheets(1).Range("BC" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 16, False)
'
'        'discharge moment
'        Sheets(1).Range("BK" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False)
'        Sheets(1).Range("BL" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False)
'        Sheets(1).Range("BM" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False)
'
'        'suction moment
'        Sheets(1).Range("BN" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False)
'        Sheets(1).Range("BO" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False)
'        Sheets(1).Range("BP" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False)
'  ElseIf Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) <> 7 Then
'      Sheets(1).Range("AX" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 4, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("AY" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 5, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("AZ" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 6, False) * Sheets("input").Range("$i$56")
'
'        'suction nozzle
'
'        Sheets(1).Range("BA" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 14, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BB" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 15, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BC" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 16, False) * Sheets("input").Range("$i$56")
'
'        'discharge moment
'        Sheets(1).Range("BK" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BL" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BM" & 40 + k) = Application.HLookup(Sheets(1).Range("S" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False) * Sheets("input").Range("$i$56")
'
'        'suction moment
'        Sheets(1).Range("BN" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 20, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BO" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 21, False) * Sheets("input").Range("$i$56")
'        Sheets(1).Range("BP" & 40 + k) = Application.HLookup(Sheets(1).Range("T" & 40 + k), Sheets("nozzle-load").Range("$b$5:$AD$27"), 22, False) * Sheets("input").Range("$i$56")
'
    End If
 '/////////////////////////////////////////////

'calculate progress course---6

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////
If Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
      origpumptype = Pump_Type
      Pump_Type = "ZE7"
End If


     'transfer PD flange
        If Sheets("input").Range("AC" & 21 + k) = "" Then

                If Sheets("input").Range("$i$53") <> "ASME" Then

                    If Pump_Type <> "ZA" And Pump_Type <> "XA" And Pump_Type <> "XE" And Pump_Type <> "CZ" Then
                    

                        If Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False) < 80 Then
                            Sheets(1).Range("V" & 40 + k) = "DN15"
                        Else
                            Sheets(1).Range("V" & 40 + k) = "DN20"
                        End If

                    Else

                        If Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 2, False) < 150 Then
                            Sheets(1).Range("V" & 40 + k) = "DN15"
                        Else
                            Sheets(1).Range("V" & 40 + k) = "DN25"
                        End If

                    End If


                Else

                    If Pump_Type <> "ZA" And Pump_Type <> "XA" And Pump_Type <> "XE" And Pump_Type <> "CZ" Then

                        If Len(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)) = 3 Or VBA.Mid(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False), 1, 1) >= 3 Then
                            Sheets(1).Range("V" & 40 + k) = "3/4"""
                        Else
                            Sheets(1).Range("V" & 40 + k) = "1/2"""
                        End If

                    Else
                        If Len(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False)) = 3 Or VBA.Mid(Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 4, False), 1, 1) >= 6 Then
                            Sheets(1).Range("V" & 40 + k) = "1"""
                        Else
                            Sheets(1).Range("V" & 40 + k) = "1/2"""
                        End If


                    End If

                End If

            Else
                Sheets(1).Range("V" & 40 + k) = Sheets("input").Range("AC" & 21 + k)
            End If
 '/////////////////////////////////////////////
    If origpumptype <> "" Then Pump_Type = origpumptype
'calculate progress course---7

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////



     'transfer nozzle list flange standard
        If Sheets("input").Range("$i$53") = "ASME" Then
            'pump discharge and suction nozzle flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("U" & 40 + k) = "ASME B16.5 150# RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("U" & 40 + k) = "ASME B16.5 300#  RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("U" & 40 + k) = "ASME B16.5 600#  RF"
            'pump PD flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("W" & 40 + k) = "ASME B16.5 150#  RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("W" & 40 + k) = "ASME B16.5 300#  RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("W" & 40 + k) = "ASME B16.5 600#  RF"

            'transfer CI/CO flange standard
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then

                Sheets(1).Range("AO" & 40 + k) = "ASME B16.5 150#  RF"

            Else
            End If

        ElseIf Sheets("input").Range("$i$53") = "HG20615" Then
            'pump discharge and suction nozzle flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("U" & 40 + k) = "HG20615 PN2.0 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("U" & 40 + k) = "HG20615 PN5.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("U" & 40 + k) = "HG20615 PN11.0 RF"
            'pump PD flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("W" & 40 + k) = "HG20615 PN2.0 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("W" & 40 + k) = "HG20615 PN5.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("W" & 40 + k) = "HG20615 PN11.0 RF"

            'transfer CI/CO flange standard
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then

                Sheets(1).Range("AO" & 40 + k) = "HG20615 PN2.0 RF"

            Else
            End If

        ElseIf Sheets("input").Range("$i$53") = "HG20592" Then
            'pump discharge and suction nozzle flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("U" & 40 + k) = "HG20592 PN2.5 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("U" & 40 + k) = "HG20592 PN4.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("U" & 40 + k) = "HG20615 PN11.0 RF"
            'pump PD flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("W" & 40 + k) = "HG20592 PN2.5 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("W" & 40 + k) = "HG20592 PN4.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("W" & 40 + k) = "HG20615 PN11.0 RF"

            'transfer CI/CO flange standard
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then

                Sheets(1).Range("AO" & 40 + k) = "HG20592 PN2.5 RF"

            Else
            End If

        ElseIf Sheets("input").Range("$i$53") = "SH3406" Then
            'pump discharge and suction nozzle flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("U" & 40 + k) = "SH3406 PN2.0 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("U" & 40 + k) = "SH3406 PN5.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("U" & 40 + k) = "SH3406 PN10.0 RF"
            'pump PD flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("W" & 40 + k) = "SH3406 PN2.0 RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("W" & 40 + k) = "SH3406 PN5.0 RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("W" & 40 + k) = "SH3406 PN10.0 RF"

            'transfer CI/CO flange standard
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then

                Sheets(1).Range("AO" & 40 + k) = "SH3406 PN2.0 RF"

            Else
            End If

        Else
            'pump discharge and suction nozzle flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("U" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("U" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$13") & " RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("U" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$14") & " RF"
            'pump PD flange standard
            If Pump_Type = "ZA" Or Pump_Type = "XA" Or Pump_Type = "CZ" Then Sheets(1).Range("W" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"
            If Pump_Type = "ZE" Or Pump_Type = "XE" Then Sheets(1).Range("W" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$13") & " RF"
            If Pump_Type = "ZF" Then Sheets(1).Range("W" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$14") & " RF"

            'transfer CI/CO flange standard
            If Sheets("input").Range("Z" & 21 + k) <> "" And Sheets("input").Range("Z" & 21 + k) <> "无" Then

                Sheets(1).Range("AO" & 40 + k) = "GB9113.1-2000 " & Sheets(1).Range("$c$12") & " RF"

            Else
            End If

        End If


 '/////////////////////////////////////////////

'calculate progress course---8

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

            'FI flange standard
            If InStr(Sheets("input").Range("Y" & 21 + k), "32") <> 0 Then

                Sheets(1).Range("AS" & 40 + k) = Sheets(1).Range("U" & 40 + k)

            Else
            End If
            'BI/BO flange standard
            If InStr(Sheets("input").Range("Y" & 21 + k), "54") <> 0 Then

                Sheets(1).Range("AW" & 40 + k) = Sheets(1).Range("U" & 40 + k)

            Else
            End If
            'QI/QO flange standard
            If InStr(Sheets("input").Range("Y" & 21 + k), "62") <> 0 Then

                Sheets(1).Range("AQ" & 40 + k) = Sheets(1).Range("U" & 40 + k)

            Else
            End If

'/////////////////////////////////////////////

'calculate progress course---9

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////


    'interface name list

    'there are 7 kinds of interface name totally as follows,
        'NO+NO+NO+NO
        'FI+NO+QI/QO+CI/CO
        'NO+NO+NO+CI/CO
        'FI+NO+NO+CI/CO
        'NO+BI/BO+NO+NO
        'FI+NO+NO+NO
        'NO+NO+QI/QO+NO

    'Mr Liu Fengqing will make list block accordingly


    If Sheets(1).Range("AR" & 40 + k) <> "NO" Then
        Sheets(1).Range("BD" & 40 + k) = "F"
    Else
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("AR" & 40 + k)
    End If

    If Sheets(1).Range("AV" & 40 + k) <> "NO" Then
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & "B"
    Else
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & Sheets(1).Range("AV" & 40 + k)
    End If

    If Sheets(1).Range("AP" & 40 + k) <> "NO" Then
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & "Q"
    Else
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & Sheets(1).Range("AP" & 40 + k)
    End If

    If Sheets(1).Range("AL" & 40 + k) <> "NO" Then
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & "C"
    Else
        Sheets(1).Range("BD" & 40 + k) = Sheets(1).Range("BD" & 40 + k) & Sheets(1).Range("AL" & 40 + k)
    End If

'/////////////////////////////////////////////

'calculate progress course---10

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

   'judge if there's heat exchanger or reservoir existing

    If InStr(Sheets("input").Range("Y" & 21 + k), "21") <> 0 Then
        Sheets(1).Range("BE" & 40 + k) = "Y"
        Sheets(1).Range("BI" & 40 + k) = "N"
    ElseIf InStr(Sheets("input").Range("Y" & 21 + k), "23") <> 0 Then
        Sheets(1).Range("BE" & 40 + k) = "Y"
    'if flush plan is Plan23, write "L" for CAD programming
        Sheets(1).Range("BI" & 40 + k) = "L"
    ElseIf InStr(Sheets("input").Range("Y" & 21 + k), "41") <> 0 Then
        Sheets(1).Range("BE" & 40 + k) = "Y"
        Sheets(1).Range("BI" & 40 + k) = "N"
    Else
        Sheets(1).Range("BE" & 40 + k) = "N"
        Sheets(1).Range("BI" & 40 + k) = "N"
    End If

     If InStr(Sheets("input").Range("Y" & 21 + k), "52") <> 0 Then
        Sheets(1).Range("BF" & 40 + k) = "Y"
    ElseIf InStr(Sheets("input").Range("Y" & 21 + k), "53") <> 0 Then
        Sheets(1).Range("BF" & 40 + k) = "Y"

    Else
        Sheets(1).Range("BF" & 40 + k) = "N"
    End If



'/////////////////////////////////////////////

'calculate progress course---11

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////



    'transfer GA, PID, SEC file name
        
        If Sheets("input").Range("AB" & 21 + k) <> "" Then
        If Sheets(1).Range("B" & 40 + k) <> "" Then
            str3 = Sheets(1).Range("B" & 40 + k) & "_" & str2new & "_" & Sheets(1).Range("E" & 40 + k) & "_" & Sheets("input").Range("U" & 21 + k) & "_Ver" & Sheets(1).Range("$c$3")
        Else
            str3 = str2new & "_" & Sheets(1).Range("E" & 40 + k) & "_" & Sheets("input").Range("U" & 21 + k) & "_Ver" & Sheets(1).Range("$c$3")
        End If

        
        
        
        Else
        If Sheets(1).Range("B" & 40 + k) <> "" Then
            str3 = Sheets(1).Range("B" & 40 + k) & "_" & str2 & "_" & Sheets(1).Range("E" & 40 + k) & "_" & Sheets("input").Range("U" & 21 + k) & "_Ver" & Sheets(1).Range("$c$3")
        Else
            str3 = str2 & "_" & Sheets(1).Range("E" & 40 + k) & "_" & Sheets("input").Range("U" & 21 + k) & "_Ver" & Sheets(1).Range("$c$3")
        End If

        End If
        
        
        
        
        
          'filter illegal character like "/ \ , space" etc

            j = Len(str3)
            name = ""
                For X = 1 To j

                    str1 = VBA.Mid(str3, X, 1)


                    If str1 <> "/" And str1 <> "\" And str1 <> " " And str1 <> ")" And str1 <> "(" And str1 <> "," And str1 <> "，" Then
                          name = name & str1
                    Else
                         str1 = ""
                          name = name & str1
                    End If
                    Sheets(1).Range("P" & 40 + k) = "GA_" & name
                    Sheets(1).Range("Q" & 40 + k) = "PID_" & name
                    Sheets(1).Range("R" & 40 + k) = "Sec&List_" & name

                    Next X
                                    
                                    
                

      'calculate available rows in ZE,ZF,ZA data source
      If Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
      Counter = Sheets("ZE7").Range("A" & Rows.Count).End(xlUp).Row - 31
      Else
      Counter = Sheets(Pump_Type).Range("A" & Rows.Count).End(xlUp).Row - 31
      End If

'/////////////////////////////////////////////

'calculate progress course---12

    Progress = 5 * CounterProgress
      CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////




If Pump_Type = "ZE" And VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
      origpumptype = Pump_Type
      Pump_Type = "ZE7"
End If

        For X = 0 To Counter - 1

      Counter2 = 0
      other_weight = 20
            If Sheets("input").Range("V" & 21 + k) <> "380V" Then

                If Sheets(Pump_Type).Range("A" & 32 + X) = str4 And Sheets(Pump_Type).Range("F" & 32 + X) = Sheets("input").Range("U" & 21 + k) & "(6KV)" Then
                    'transfer motor weight
                    Sheets(1).Range("G" & 40 + k) = Sheets(Pump_Type).Range("y" & X + 32)
                    'transfer baseplate weight
                    Sheets(1).Range("H" & 40 + k) = Sheets(Pump_Type).Range("z" & X + 32)
                    'transfer others weight

                    For Y = 0 To 6

                        str5 = Sheets("Data1").Range("A" & 2 + Y)
                        If InStr(Sheets(1).Range("E" & 40 + k), str5) <> 0 Then
                            Counter2 = Counter2 + 1

                        Else
                        End If



                    Next Y
                    Sheets(1).Range("I" & 40 + k) = other_weight + 100 * Counter2
                    'transfer total weight
                    'total weight will be got by function

                    'Sheets(1).Range("J" & 40 + k) = Sheets(1).Range("F" & 40 + k) + Sheets(1).Range("G" & 40 + k) + Sheets(1).Range("H" & 40 + k) + Sheets(1).Range("I" & 40 + k)


                    'transfer H value
                    Sheets(1).Range("X" & 40 + k) = Sheets(Pump_Type).Range("g" & X + 32)
                    'transfer L value
                    Sheets(1).Range("Y" & 40 + k) = Sheets(Pump_Type).Range("h" & X + 32)
                    'transfer L1 value
                    Sheets(1).Range("Z" & 40 + k) = Sheets(Pump_Type).Range("i" & X + 32)
                     'transfer L2 value
                    Sheets(1).Range("AA" & 40 + k) = Sheets(Pump_Type).Range("j" & X + 32)
                     'transfer L3 value
                    Sheets(1).Range("AB" & 40 + k) = Sheets(Pump_Type).Range("k" & X + 32)
                      'transfer k value
                      Sheets(1).Range("AC" & 40 + k) = Sheets(Pump_Type).Range("l" & X + 32)
                    If Sheets("input").Range("AB" & 21 + k) <> "" Then
                        Sheets(1).Range("CN" & 40 + k) = Sheets("input").Range("AB" & 21 + k)
                    End If

                      'transfer B value
                    Sheets(1).Range("AD" & 40 + k) = Sheets(Pump_Type).Range("m" & X + 32)
                      'transfer B1 value
                    Sheets(1).Range("AE" & 40 + k) = Sheets(Pump_Type).Range("n" & X + 32)
                      'transfer i value
                    Sheets(1).Range("AF" & 40 + k) = Sheets(Pump_Type).Range("o" & X + 32)
                      'transfer h value
                    Sheets(1).Range("AG" & 40 + k) = Sheets(Pump_Type).Range("p" & X + 32)
                      'transfer h1 value
                    Sheets(1).Range("AH" & 40 + k) = Sheets(Pump_Type).Range("q" & X + 32)
                      'transfer h2 value
                    Sheets(1).Range("AI" & 40 + k) = Sheets(Pump_Type).Range("r" & X + 32)
                      'transfer e value
                    Sheets(1).Range("AJ" & 40 + k) = Sheets(Pump_Type).Range("s" & X + 32)
                      'transfer y value
                    Sheets(1).Range("AM" & 40 + k) = Sheets(Pump_Type).Range("v" & X + 32)
                       'transfer drain outlet height value
                    Sheets(1).Range("AN" & 40 + k) = Sheets(Pump_Type).Range("w" & X + 32)
                        'transfer anchor bolt value
                    Sheets(1).Range("AT" & 40 + k) = Sheets(Pump_Type).Range("ac" & X + 32)
                        'transfer anchor bolt amount value
                    Sheets(1).Range("AU" & 40 + k) = Sheets(Pump_Type).Range("ad" & X + 32)
                        'transfer L4 value---only for ZA pump
                    If Pump_Type = "ZA" Or Pump_Type = "CZ" Then
                        Sheets(1).Range("BQ" & 40 + k) = Sheets(Pump_Type).Range("ae" & X + 32)
                    Else
                        Sheets(1).Range("BQ" & 40 + k) = "N"
                    End If

                Else

                End If

            Else

                 If Sheets(Pump_Type).Range("A" & 32 + X) = str4 And Sheets(Pump_Type).Range("F" & 32 + X) = Sheets("input").Range("U" & 21 + k) Then
                    Sheets(1).Range("G" & 40 + k) = Sheets(Pump_Type).Range("y" & X + 32)
                    Sheets(1).Range("H" & 40 + k) = Sheets(Pump_Type).Range("z" & X + 32)


                    For Y = 0 To 6

                        str5 = Sheets("Data1").Range("A" & 2 + Y)
                        If InStr(Sheets(1).Range("E" & 40 + k), str5) <> 0 Then
                            Counter2 = Counter2 + 1

                        Else
                        End If



                    Next Y
                    Sheets(1).Range("I" & 40 + k) = other_weight + 100 * Counter2

                    'Sheets(1).Range("J" & 40 + k) = Sheets(1).Range("F" & 40 + k) + Sheets(1).Range("G" & 40 + k) + Sheets(1).Range("H" & 40 + k) + Sheets(1).Range("I" & 40 + k)

                    Sheets(1).Range("X" & 40 + k) = Sheets(Pump_Type).Range("g" & X + 32)
                    Sheets(1).Range("Y" & 40 + k) = Sheets(Pump_Type).Range("h" & X + 32)
                    Sheets(1).Range("Z" & 40 + k) = Sheets(Pump_Type).Range("i" & X + 32)
                    Sheets(1).Range("AA" & 40 + k) = Sheets(Pump_Type).Range("j" & X + 32)
                    Sheets(1).Range("AB" & 40 + k) = Sheets(Pump_Type).Range("k" & X + 32)
                    Sheets(1).Range("AC" & 40 + k) = Sheets(Pump_Type).Range("l" & X + 32)

                    If Sheets("input").Range("AB" & 21 + k) <> "" Then
                        Sheets(1).Range("CN" & 40 + k) = Sheets("input").Range("AB" & 21 + k)
                    End If


                    Sheets(1).Range("AD" & 40 + k) = Sheets(Pump_Type).Range("m" & X + 32)
                    Sheets(1).Range("AE" & 40 + k) = Sheets(Pump_Type).Range("n" & X + 32)
                    Sheets(1).Range("AF" & 40 + k) = Sheets(Pump_Type).Range("o" & X + 32)
                    Sheets(1).Range("AG" & 40 + k) = Sheets(Pump_Type).Range("p" & X + 32)
                    Sheets(1).Range("AH" & 40 + k) = Sheets(Pump_Type).Range("q" & X + 32)
                    Sheets(1).Range("AI" & 40 + k) = Sheets(Pump_Type).Range("r" & X + 32)
                    Sheets(1).Range("AJ" & 40 + k) = Sheets(Pump_Type).Range("s" & X + 32)
                    Sheets(1).Range("AM" & 40 + k) = Sheets(Pump_Type).Range("v" & X + 32)
                    Sheets(1).Range("AN" & 40 + k) = Sheets(Pump_Type).Range("w" & X + 32)
                    Sheets(1).Range("AT" & 40 + k) = Sheets(Pump_Type).Range("ac" & X + 32)
                    Sheets(1).Range("AU" & 40 + k) = Sheets(Pump_Type).Range("ad" & X + 32)
                    If Pump_Type = "ZA" Or Pump_Type = "CZ" Then
                        Sheets(1).Range("BQ" & 40 + k) = Sheets(Pump_Type).Range("ae" & X + 32)
                    Else
                        Sheets(1).Range("BQ" & 40 + k) = "N"
                    End If


                Else

                End If



            End If

           

            'transfer f or f* value, for ZF and ZA, use f value, for ZE with fan cooling feature, use f* value

            If Pump_Type <> "ZE" Then Sheets(1).Range("AK" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 20, False)

            If Pump_Type = "ZE" Then
                If InStr(Sheets("input").Range("E" & 21 + k), "8") <> 0 And InStr(Sheets("input").Range("AA" & 21 + k), "F") Then
                    Sheets(1).Range("AK" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 21, False)


                ElseIf InStr(Sheets("input").Range("E" & 21 + k), "0th") <> 0 Then
                    Sheets(1).Range("AK" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 21, False)
                
                ElseIf InStr(Sheets("input").Range("E" & 21 + k), "7") <> 0 Then
                    Sheets(1).Range("AK" & 40 + k) = Application.VLookup(str4, Sheets("ZE7").Range("$a$31:$bb$2000"), 21, False)
                    
                Else
                    Sheets(1).Range("AK" & 40 + k) = Application.VLookup(str4, Sheets(Pump_Type).Range("$a$31:$bb$2000"), 20, False)
                End If
            End If



            'calculate heat exchanger and reservior position dimensions

            If Sheets(1).Range("BE" & 40 + k) <> "N" Then

                Sheets(1).Range("BG" & 40 + k) = Sheets(1).Range("AF" & 40 + k) + Sheets(1).Range("AK" & 40 + k) - 200


                If VBA.Mid(Sheets(1).Range("BG" & 40 + k), Len(Sheets(1).Range("BG" & 40 + k)), 1) <> 5 Then Sheets(1).Range("BG" & 40 + k) = VBA.Mid(Sheets(1).Range("BG" & 40 + k), 1, Len(Sheets(1).Range("BG" & 40 + k)) - 1) & "0"

            Else
                Sheets(1).Range("BG" & 40 + k) = "N"
            End If

            If Sheets(1).Range("BF" & 40 + k) <> "N" Then
                Sheets(1).Range("BH" & 40 + k) = Sheets(1).Range("AF" & 40 + k) + Sheets(1).Range("AK" & 40 + k) - 100

                If VBA.Mid(Sheets(1).Range("BH" & 40 + k), Len(Sheets(1).Range("BH" & 40 + k)), 1) <> 5 Then Sheets(1).Range("BH" & 40 + k) = VBA.Mid(Sheets(1).Range("BH" & 40 + k), 1, Len(Sheets(1).Range("BH" & 40 + k)) - 1) & "0"



            Else

                Sheets(1).Range("BH" & 40 + k) = "N"

            End If

            'calculate AD2 value
            If Pump_Type = "ZA" Or Pump_Type = "CZ" Then

                Sheets(1).Range("BJ" & 40 + k) = "1"""
           Else

                Sheets(1).Range("BJ" & 40 + k) = "2"""

           End If







        Next X
'/////////////////////////////////////////////
If origpumptype <> "" Then Pump_Type = origpumptype

'calculate progress course---13

    Progress = 5 * CounterProgress
    CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

        'fill "是否为ZE+G（Y or N）" information
          If InStr(Sheets(1).Range("C" & 40 + k), "G") <> 0 Or InStr(Sheets(1).Range("C" & 40 + k), "g") <> 0 Then
            Sheets(1).Range("BY" & 40 + k) = "Y"
          Else
            Sheets(1).Range("BY" & 40 + k) = "N"
          End If
        'fill "轴承架号是否>=LK6（Y or N）" information
           j = Len(Sheets(1).Range("C" & 40 + k))
           If Pump_Type <> "CZ" Then
           If VBA.Mid(Sheets(1).Range("C" & 40 + k), j - 3, 1) >= 6 Then
             Sheets(1).Range("BZ" & 40 + k) = "Y"
           Else
             Sheets(1).Range("BZ" & 40 + k) = "N"
           End If
           End If
         'fill Cooling method
           Sheets(1).Range("CB" & 40 + k) = Sheets("input").Range("AA" & 21 + k)
         'fill section block name
            j = Len(Sheets("input").Range("E" & 21 + k))

            If Sheets(1).Range("BY" & 40 + k) = "Y" Then

                If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 8 Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "8G"
                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 0 Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "10G"
                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 7 And Pump_Type = "ZE" Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "7G"
                Else
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "G"
                End If

            ElseIf Sheets(1).Range("BY" & 40 + k) = "N" Then
                 If VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 8 Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "8"
                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 0 Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "10"
                ElseIf VBA.Left(VBA.Right(VBA.Replace(Sheets("input").Range("E" & 21 + k), " ", ""), 3), 1) = 7 And Pump_Type = "ZE" Then
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type & "7"
                Else
                    Sheets(1).Range("CA" & 40 + k) = Pump_Type
                End If

            End If
           'fill section list name
            j = Len(Sheets("input").Range("E" & 21 + k))
            Dim ver As String
            ver = VBA.Left(Sheets("input").Range("E" & 21 + k), 1)
        If Pump_Type = "ZE" Then
            
            
            If Sheets(1).Range("BY" & 40 + k) = "Y" Then

                If Sheets(1).Range("BZ" & 40 + k) = "N" Then

                    If VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 8 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "8G" & "_" & Sheets(1).Range("CB" & 40 + k)
                    ElseIf VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "7G" & "_" & Sheets(1).Range("CB" & 40 + k)
                    Else
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "10" & "_" & Sheets(1).Range("CB" & 40 + k)

                    End If
                Else
                    If VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 8 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "8G" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)
                    ElseIf VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "7G" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)
                    Else
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "10" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)

                    End If

                End If

            Else
                If Sheets(1).Range("BZ" & 40 + k) = "N" Then

                    If VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 8 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "8" & "_" & Sheets(1).Range("CB" & 40 + k)
                    ElseIf VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "7" & "_" & Sheets(1).Range("CB" & 40 + k)
                    Else
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "10" & "_" & Sheets(1).Range("CB" & 40 + k)

                    End If
                Else
                    If VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 8 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "8" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)
                    If VBA.Val(VBA.Left(Sheets("input").Range("E" & 21 + k), 1)) = 7 Then
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "7" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)
                    Else
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type & "10" & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)

                    End If

                    End If

                End If
                
            End If
            
            
        ElseIf Pump_Type = "CZ" Then
        Dim CZBearingsize As String, Bearingtype As String
        CZBearingsize = Application.VLookup(str4, Sheets("CZ-BearingSize").Range("$A$2:$B$41"), 2, False)
        Bearingtype = Sheets("input").Range("AA" & 21 + k)
        Sheets(1).Range("CD" & 40 + k) = CZBearingsize
        
        If VBA.Right(VBA.Left(CZBearingsize, 3), 1) >= 6 Then
        Sheets(1).Range("CC" & 40 + k) = "CZWBALHL"
        Else
        If Val(NorDia) >= 315 Then
            If Bearingtype = "CZ轻轴" Then
            Sheets(1).Range("CC" & 40 + k) = "CZWBALHS"
            ElseIf Bearingtype = "CZ重轴" Then
            Sheets(1).Range("CC" & 40 + k) = "CZWBALHD"
            End If
        Else
            If Bearingtype = "CZ轻轴" Then
            Sheets(1).Range("CC" & 40 + k) = "CZWOBALHS"
            ElseIf Bearingtype = "CZ重轴" Then
            Sheets(1).Range("CC" & 40 + k) = "CZWOBALHD"
            End If
        End If
        
        
        End If
        
        
        
            

        Else
                If Sheets(1).Range("BZ" & 40 + k) = "N" Then

                        Sheets(1).Range("CC" & 40 + k) = Pump_Type2 & "_" & Sheets(1).Range("CB" & 40 + k)

                Else
                        Sheets(1).Range("CC" & 40 + k) = Pump_Type2 & "(LK6)_" & Sheets(1).Range("CB" & 40 + k)

                End If

        End If



'/////////////////////////////////////////////

'calculate progress course---14

    Progress = 5 * CounterProgress
    CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

'/////////////////////////////////////////////

'calculate progress course---15

    Progress = 5 * CounterProgress
    CounterProgress = CounterProgress + 1
    PctDone = Progress / CounterSum
        Call UpdateProgress(PctDone)

'//////////////////////////////////////////////

    Next k

  Unload UserForm2

'=======================================
'check value if there's space

    CounterCheck = 0
    Sheets(1).Range("$b$40:$cc$" & (39 + i)).Select

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False
    For Each cell In Selection
        If cell.Value = "" Then
          cell.Interior.ColorIndex = 3
          CounterCheck = CounterCheck + 1

        Else
          cell.Interior.ColorIndex = xlNone
        End If
    Next cell

    If CounterCheck <> 0 Then
        MsgBox ("计算结果中发现空值，请确认或修正，谢谢。")
        Exit Sub
    End If

MsgBox ("计算完成，如需特殊输入，请在本页修正，谢谢。")

'===================

'Application.Visible = True

'====================

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////

'used for progress

Sub UpdateProgress(Pct)
    With UserForm2
        .FrameProgress.Caption = Format(Pct, "0%")
        .LabelProgress.Width = Pct * (.FrameProgress.Width - 10)
        .Repaint
    End With
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////

