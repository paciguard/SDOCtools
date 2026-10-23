Attribute VB_Name = "Module8"

Sub test()

Dim m, n, backresult As String
m = "A"
n = Sheets(1).Range("c69")
backresult = InStr(n, m)


End Sub

Sub test2()

m = Sheets(1).Range("c74")
n = Sheets(1).Range("c75")

If InStr(m, "o") <> 0 Or InStr(m, "6") <> 0 Then
    Sheets(1).Range("d74") = "OK"
Else
    Sheets(1).Range("d74") = "NO"
End If

End Sub

Sub test3()

Sheets("nozzle-load").Range("b31") = Application.HLookup(Sheets("nozzle-load").Range("j5"), Sheets("nozzle-load").Range("$b$5:$x$27"), 5, False)

End Sub

Sub test4()

Sheets(1).Range("b73") = ActiveWorkbook.Sheets.Count

End Sub

Sub test5()

'   Makes cell background red if the value is negative
    Dim cell As Range
    Sheets(1).Range("$p$40:$bx$69").Select
    If TypeName(Selection) <> "Range" Then Exit Sub
   
    Application.ScreenUpdating = False
    For Each cell In Selection
        If cell.Value = "" Then
          cell.Interior.ColorIndex = 3
        Else
          cell.Interior.ColorIndex = xlNone
        End If
    Next cell


End Sub



'Sub Main()
'   Inserts random numbers on the active worksheet
   ' Dim Counter As Integer
   ' Dim RowMax As Integer, ColMax As Integer
   ' Dim r As Integer, c As Integer
   ' Dim PctDone As Single
    
    'If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    'Cells.Clear
    'Counter = 1
    'RowMax = 200
    'ColMax = 25
   '' For r = 1 To RowMax
        'For c = 1 To ColMax
            'Cells(r, c) = Int(Rnd * 1000)
           ' Counter = Counter + 1
     '   Next c
     '   PctDone = Counter / (RowMax * ColMax)
    '    Call UpdateProgress(PctDone)
  '  Next r
  '  Unload UserForm1
'End Sub


    Sub test7()

    Dim i As Integer
    Dim cell As Range
    CounterCheck = 0
    i = 24
    Application.Union(Sheets(1).Range("$p$40:$bx$" & (39 + i)), Sheets(1).Range("$c$40:$c$" & (39 + i)), Sheets(1).Range("$e$40:$j$" & (39 + i))).Select

    If TypeName(Selection) <> "Range" Then Exit Sub
   
    Application.ScreenUpdating = False
    For Each cell In Selection
        If cell.Value = "" Or "#N/A" Or "#VALUE" Then
          cell.Interior.ColorIndex = 3
          CounterCheck = CounterCheck + 1
        Else
          cell.Interior.ColorIndex = xlNone
        End If
    Next cell

    If CounterCheck <> 0 Then MsgBox "data error"
    
    End Sub


Sub test23123414()

Sheets(1).Range("d8:e16").Copy

Sheets(1).Range("g9").Select
    Sheets(1).Paste



End Sub



