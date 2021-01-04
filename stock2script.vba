Type log
    Sheet As String
    color As String
    Sign As String
    Amount As Long
    FloatAmount As Single
End Type
Private Sub stock2_fillalertlist(column As Integer, color As String, returnto As String)

Workbooks(ActiveWorkbook.Name).Activate
Sheets("BACKEND").Select

Cells(Cells(1, column).Value, column) = color
Cells(1, column).Value = Cells(1, column).Value + 1

Sheets(returnto).Select

End Sub
Private Sub stock2_logtofile(CurrentSheet As String, v() As log, vlength As Integer)

Workbooks(ActiveWorkbook.Name).Activate
Sheets(CurrentSheet).Select

Dim TextFile As Integer
Dim FilePath As String
Dim m As String
Dim Y As String

m = Month(Date)
Y = Year(Date)
FilePath = ActiveWorkbook.Path & "\Istoric " & m & " " & Y & ".txt"

TextFile = FreeFile

Open FilePath For Append As TextFile

Print #TextFile, "Time: "; Now
Print #TextFile, "Username: "; Application.UserName
Print #TextFile, "Modifications: "

Dim i As Integer
For i = 0 To vlength
    If (v(i).Amount = 0 And v(i).FloatAmount > 0) Then
        Print #TextFile, v(i).Sheet & " " & v(i).color & " " & v(i).Sign & v(i).FloatAmount
    Else
        Print #TextFile, v(i).Sheet & " " & v(i).color & " " & v(i).Sign & v(i).Amount
    End If
Next i

Print #TextFile, ""
  
Close TextFile

End Sub
Sub stock2_35cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("A1:B1").Value = 3
Worksheets("BACKEND").Range("A3:B" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "35CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(1, Cells(i, 1).Text, "35CM")
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(2, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_35cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

Dim CurrentSheet As String
CurrentSheet = "35CM"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_35cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(1, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(2, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_12cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("C1:D1").Value = 3
Worksheets("BACKEND").Range("C3:D" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "12CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(3, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(4, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_12cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate
Dim CurrentSheet As String
CurrentSheet = "12CM"
Sheets(CurrentSheet).Select

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

If ActiveSheet.boxCV_12cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(3, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(4, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_26cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("E1:F1").Value = 3
Worksheets("BACKEND").Range("E3:F" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "26CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(5, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(6, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_26cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate
Dim CurrentSheet As String
CurrentSheet = "26CM"
Sheets(CurrentSheet).Select

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

If ActiveSheet.boxCV_26cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(5, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(6, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_31cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("G1:H1").Value = 3
Worksheets("BACKEND").Range("G3:H" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "31CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(7, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(8, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_31cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate
Dim CurrentSheet As String
CurrentSheet = "31CM"
Sheets(CurrentSheet).Select

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

If ActiveSheet.boxCV_31cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(7, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(8, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_40cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("I1:J1").Value = 3
Worksheets("BACKEND").Range("I3:J" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "40CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(9, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(10, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_40cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

Dim CurrentSheet As String
CurrentSheet = "40CM"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_40cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(9, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(10, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_48cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 10).Value

'clear alert list
Worksheets("BACKEND").Range("K1:K1").Value = 3
Worksheets("BACKEND").Range("K3:K" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "48CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(11, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_48cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 10).Value

Dim CurrentSheet As String
CurrentSheet = "48CM"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_48cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 9).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(11, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_75cm()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

'clear alert list
Worksheets("BACKEND").Range("L1:M1").Value = 3
Worksheets("BACKEND").Range("L3:M" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "75CM"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(12, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(13, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_75cm_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 15).Value

Dim CurrentSheet As String
CurrentSheet = "75CM"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_75cm.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 14).Value
    Range("I2:I" & leng).Value = Cells(3, 14).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(12, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(13, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_hearts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

'clear alert list
Worksheets("BACKEND").Range("N1:N1").Value = 3
Worksheets("BACKEND").Range("N3:N" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "HEARTS"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(12, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_hearts_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

Dim CurrentSheet As String
CurrentSheet = "HEARTS"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_hearts.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 10).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(14, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

End Sub
Sub stock2_link()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 16).Value

'clear alert list
Worksheets("BACKEND").Range("O1:P1").Value = 3
Worksheets("BACKEND").Range("O3:P" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "LINK"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet & " PASTEL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(15, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
    If Not IsEmpty(Cells(i, 10)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 10).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value + Cells(i, 10).Value
        Cells(i, 10).Clear
    End If
    
    If Not IsEmpty(Cells(i, 11)) Then
        v(vpos).Sheet = CurrentSheet & " METAL"
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 11).Value
        vpos = vpos + 1
        Cells(i, 8).Value = Cells(i, 8).Value - Cells(i, 11).Value
        Cells(i, 11).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(16, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_link_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 16).Value

Dim CurrentSheet As String
CurrentSheet = "LINK"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_link.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Columns("I:I").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 15).Value
    Range("I2:I" & leng).Value = Cells(3, 15).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
    Columns("I:I").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(15, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
If (Cells(i, 8).Value <= Cells(i, 9).Value) And Not IsEmpty(Cells(i, 8)) Then
        Call stock2_fillalertlist(16, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 8).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 8).Value > Cells(i, 9).Value) Or IsEmpty(Cells(i, 8)) Then
        Cells(i, 8).Interior.ColorIndex = 0
    End If
Next i

End Sub
Sub stock2_mfc()
'modelaj forme chrome (mfc)

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

'clear alert list
Worksheets("BACKEND").Range("Q1:Q1").Value = 3
Worksheets("BACKEND").Range("Q3:Q" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "MFC"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(17, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_mfc_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

Dim CurrentSheet As String
CurrentSheet = "MFC"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_mfc.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 10).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(17, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

End Sub
Sub stock2_disp()
'dispozitive

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

'clear alert list
Worksheets("BACKEND").Range("R1:R1").Value = 3
Worksheets("BACKEND").Range("R3:R" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "DISPOZITIVE"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(18, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_disp_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

Dim CurrentSheet As String
CurrentSheet = "DISPOZITIVE"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_disp.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 10).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(18, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

End Sub
Sub stock2_accesorii()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

'clear alert list
Worksheets("BACKEND").Range("S1:S1").Value = 3
Worksheets("BACKEND").Range("S3:S" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "ACCESORII"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(19, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_accesorii_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 11).Value

Dim CurrentSheet As String
CurrentSheet = "ACCESORII"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_accs.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 10).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(19, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

End Sub
Sub stock2_confetti()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 12).Value

'clear alert list
Worksheets("BACKEND").Range("T1:T1").Value = 3
Worksheets("BACKEND").Range("T3:T46").Clear

Dim CurrentSheet As String
CurrentSheet = "CONFETTI"
Sheets(CurrentSheet).Select

Dim v(64) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        
        If (Not IsEmpty(Cells(i, 2))) Then
            v(vpos).color = v(vpos).color & " " & Cells(i, 2).Text
        End If
        If (Not IsEmpty(Cells(i, 7))) Then
            v(vpos).color = v(vpos).color & " " & Cells(i, 7).Text
        End If
        
        v(vpos).Sign = "+"
        v(vpos).FloatAmount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        
        If (Not IsEmpty(Cells(i, 2))) Then
            v(vpos).color = v(vpos).color + Cells(i, 2).Text
        End If
        If (Not IsEmpty(Cells(i, 7))) Then
            v(vpos).color = v(vpos).color + Cells(i, 7).Text
        End If
        
        v(vpos).Sign = "-"
        v(vpos).FloatAmount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(20, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_confetti_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(15, 12).Value

Dim CurrentSheet As String
CurrentSheet = "CONFETTI"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_conf.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 11).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(20, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

End Sub
Sub stock2_print()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(18, 13).Value

'clear alert list
Worksheets("BACKEND").Range("U1:U1").Value = 3
Worksheets("BACKEND").Range("U3:U" & (leng + 3)).Clear

Dim CurrentSheet As String
CurrentSheet = "PRINT"
Sheets(CurrentSheet).Select

Dim v(200) As log
Dim vpos As Integer
vpos = 0

Dim i As Integer
For i = 2 To leng

    'add to pastel
    If Not IsEmpty(Cells(i, 5)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        
        If (Not IsEmpty(Cells(i, 7))) Then
            v(vpos).color = v(vpos).color & " " & Cells(i, 7).Text
        End If
        
        v(vpos).Sign = "+"
        v(vpos).Amount = Cells(i, 5).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value + Cells(i, 5).Value
        Cells(i, 5).Clear
    End If
    
    'subtract from pastel
    If Not IsEmpty(Cells(i, 6)) Then
        v(vpos).Sheet = CurrentSheet
        v(vpos).color = Cells(i, 1).Text
        
        If (Not IsEmpty(Cells(i, 2))) Then
            v(vpos).color = v(vpos).color + Cells(i, 2).Text
        End If
        If (Not IsEmpty(Cells(i, 7))) Then
            v(vpos).color = v(vpos).color + Cells(i, 7).Text
        End If
        
        v(vpos).Sign = "-"
        v(vpos).Amount = Cells(i, 6).Value
        vpos = vpos + 1
        Cells(i, 3).Value = Cells(i, 3).Value - Cells(i, 6).Value
        Cells(i, 6).Clear
    End If
    
    'check if alert is needed
    If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(21, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
    If (Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3)) Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    
Next i

If vpos > 0 Then
    Call stock2_logtofile(CurrentSheet, v, vpos - 1)
End If

End Sub
Sub stock2_print_updatealerts()

Workbooks(ActiveWorkbook.Name).Activate

'file length
Dim leng As Integer
leng = Cells(18, 13).Value

Dim CurrentSheet As String
CurrentSheet = "PRINT"
Sheets(CurrentSheet).Select

If ActiveSheet.boxCV_print.Value = False Then
    Columns("D:D").EntireColumn.Hidden = True
    Range("D2:D" & leng).Value = Cells(3, 12).Value
Else
    Columns("D:D").EntireColumn.Hidden = False
End If

'recheck for new alerts
Dim i As Integer
For i = 2 To leng
If (Cells(i, 3).Value <= Cells(i, 4).Value) And Not IsEmpty(Cells(i, 3)) Then
        Call stock2_fillalertlist(21, Cells(i, 1).Text, CurrentSheet)
        Cells(i, 3).Interior.color = RGB(255, 202, 202)
    End If
    
If ((Cells(i, 3).Value > Cells(i, 4).Value) Or IsEmpty(Cells(i, 3))) Then
    If Not (Cells(i, 1).Text = "ignore") Then
        Cells(i, 3).Interior.ColorIndex = 0
    End If
    End If
    
Next i

End Sub
Sub stock2_print_scrollToTop()

Workbooks(ActiveWorkbook.Name).Activate
Dim CurrentSheet As String
CurrentSheet = "PRINT"
Sheets(CurrentSheet).Select
ActiveWindow.ScrollRow = 1

End Sub
Sub stock2_print_hideIDs()

Workbooks(ActiveWorkbook.Name).Activate

Dim CurrentSheet As String
CurrentSheet = "PRINT"
Sheets(CurrentSheet).Select

If Cells(20, 13).Value = 1 Then
    Cells(20, 13).Value = 0
    Columns("G:I").EntireColumn.Hidden = False
Else
    Cells(20, 13).Value = 1
    Columns("G:I").EntireColumn.Hidden = True
End If

End Sub
Sub stock2_installupdate()

Workbooks(ActiveWorkbook.Name).Activate
Sheets("UPDATE").Select

'set up what worksheet we copy from and what worksheet we copy to
Dim wfrom As String
Dim wto As String
Dim limit As Integer

wfrom = Cells(7, 3).Text
wto = Cells(9, 3).Text

'create update logs
Dim TextFile As Integer
Dim FilePath As String
Dim CurrentTime As Date
CurrentTime = Now()

FilePath = ActiveWorkbook.Path & "\Update log.txt"

TextFile = FreeFile

Open FilePath For Append As TextFile

Print #TextFile, "Time: "; Now
Print #TextFile, "Username: "; Application.UserName
Print #TextFile, "Beggining update from " & wfrom & " to " & wto

'update information
MsgBox "A text file called /Update log.txt/ will generate in the same folder as this worksheet. " & vbCrLf _
& "Please check it at the end to see if every sheet was copied over succesfully." & vbCrLf _
& "If you do not see the text /UPDATE PROCESS FINISHED WITH CODE 1/ at the bottom, then something has gone wrong." & vbCrLf _
& "The update process will also attempt to open the file upon completion. The file not being automatically open is another sign of a failure."

'check if update is allowed
If ActiveSheet.ALLOWUPDATE.Value = False Then
    MsgBox "Please tick the consent box before updating."
    Close TextFile
    Exit Sub
End If

'reset allowupdate checkbox
ActiveSheet.ALLOWUPDATE.Value = False
Print #TextFile, "Succesfully reset the allow update checkbox"

'last chance to cancel
Dim Confirm
Confirm = MsgBox("Do not click anything or switch windows until completion." & vbCrLf & "Press OK to start the update process.", vbOKCancel)
If Confirm = vbCancel Then
    Exit Sub
End If

'12cm
Workbooks(wfrom).Activate
limit = Worksheets("12CM").Cells(15, 15).Value
Worksheets("12CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("12CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("12CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("12CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("12CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("12CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 12CM sheet from " & wfrom & " to " & wto

'26cm
Workbooks(wfrom).Activate
limit = Worksheets("26CM").Cells(15, 15).Value
Worksheets("26CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("26CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("26CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("26CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("26CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("26CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 26CM sheet from " & wfrom & " to " & wto

'31cm
Workbooks(wfrom).Activate
limit = Worksheets("31CM").Cells(15, 15).Value
Worksheets("31CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("31CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("31CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("31CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("31CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("31CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 31CM sheet from " & wfrom & " to " & wto

'35cm
Workbooks(wfrom).Activate
limit = Worksheets("35CM").Cells(15, 15).Value
Worksheets("35CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("35CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("35CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("35CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("35CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("35CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 35CM sheet from " & wfrom & " to " & wto

'40cm
Workbooks(wfrom).Activate
limit = Worksheets("40CM").Cells(15, 15).Value
Worksheets("40CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("40CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("40CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("40CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("40CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("40CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 40CM sheet from " & wfrom & " to " & wto

'48cm
Workbooks(wfrom).Activate
limit = Worksheets("48CM").Cells(15, 10).Value
Worksheets("48CM").Range("A2:F" & limit).Copy
Workbooks(wto).Activate
Worksheets("48CM").Range("A2:F" & limit).PasteSpecial xlPasteValues
Worksheets("48CM").Range("A2:F" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("48CM").Range("H3:J15").Copy
Workbooks(wto).Activate
Worksheets("48CM").Range("H3:J15").PasteSpecial xlPasteValues
Worksheets("48CM").Range("H3:J15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 48CM sheet from " & wfrom & " to " & wto

'75cm
Workbooks(wfrom).Activate
limit = Worksheets("75CM").Cells(15, 15).Value
Worksheets("75CM").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("75CM").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("75CM").Range("A2:K" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("75CM").Range("M3:O15").Copy
Workbooks(wto).Activate
Worksheets("75CM").Range("M3:O15").PasteSpecial xlPasteValues
Worksheets("75CM").Range("M3:O15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied 75CM sheet from " & wfrom & " to " & wto

'hearts
Workbooks(wfrom).Activate
limit = Worksheets("HEARTS").Cells(15, 11).Value
Worksheets("HEARTS").Range("A2:G" & limit).Copy
Workbooks(wto).Activate
Worksheets("HEARTS").Range("A2:G" & limit).PasteSpecial xlPasteValues
Worksheets("HEARTS").Range("A2:G" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("HEARTS").Range("I3:K15").Copy
Workbooks(wto).Activate
Worksheets("HEARTS").Range("I3:K15").PasteSpecial xlPasteValues
Worksheets("HEARTS").Range("I3:K15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied HEARTS sheet from " & wfrom & " to " & wto

'jumbo
Workbooks(wfrom).Activate
limit = 50
Worksheets("JUMBO").Range("A2:K" & limit).Copy
Workbooks(wto).Activate
Worksheets("JUMBO").Range("A2:K" & limit).PasteSpecial xlPasteValues
Worksheets("JUMBO").Range("A2:K" & limit).PasteSpecial xlFormats
Print #TextFile, "Succesfully copied JUMBO sheet from " & wfrom & " to " & wto

'link
Workbooks(wfrom).Activate
limit = Worksheets("LINK").Cells(15, 16).Value
Worksheets("LINK").Range("A2:L" & limit).Copy
Workbooks(wto).Activate
Worksheets("LINK").Range("A2:L" & limit).PasteSpecial xlPasteValues
Worksheets("LINK").Range("A2:L" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("LINK").Range("N3:P15").Copy
Workbooks(wto).Activate
Worksheets("LINK").Range("N3:P15").PasteSpecial xlPasteValues
Worksheets("LINK").Range("N3:P15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied LINK sheet from " & wfrom & " to " & wto

'mfc
Workbooks(wfrom).Activate
limit = Worksheets("MFC").Cells(15, 11).Value
Worksheets("MFC").Range("A2:G" & limit).Copy
Workbooks(wto).Activate
Worksheets("MFC").Range("A2:G" & limit).PasteSpecial xlPasteValues
Worksheets("MFC").Range("A2:G" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("MFC").Range("I3:K15").Copy
Workbooks(wto).Activate
Worksheets("MFC").Range("I3:K15").PasteSpecial xlPasteValues
Worksheets("MFC").Range("I3:K15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied MODELAJ/FORME/CHROME (MFC) sheet from " & wfrom & " to " & wto

'dispozitive
Workbooks(wfrom).Activate
limit = Worksheets("DISPOZITIVE").Cells(15, 11).Value
Worksheets("DISPOZITIVE").Range("A2:G" & limit).Copy
Workbooks(wto).Activate
Worksheets("DISPOZITIVE").Range("A2:G" & limit).PasteSpecial xlPasteValues
Worksheets("DISPOZITIVE").Range("A2:G" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("DISPOZITIVE").Range("I3:K15").Copy
Workbooks(wto).Activate
Worksheets("DISPOZITIVE").Range("I3:K15").PasteSpecial xlPasteValues
Worksheets("DISPOZITIVE").Range("I3:K15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied DISPOZITIVE sheet from " & wfrom & " to " & wto

'accesorii
Workbooks(wfrom).Activate
limit = Worksheets("ACCESORII").Cells(15, 11).Value
Worksheets("ACCESORII").Range("A2:G" & limit).Copy
Workbooks(wto).Activate
Worksheets("ACCESORII").Range("A2:G" & limit).PasteSpecial xlPasteValues
Worksheets("ACCESORII").Range("A2:G" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("ACCESORII").Range("I3:K15").Copy
Workbooks(wto).Activate
Worksheets("ACCESORII").Range("I3:K15").PasteSpecial xlPasteValues
Worksheets("ACCESORII").Range("I3:K15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied ACCESORII sheet from " & wfrom & " to " & wto

'confetti
Workbooks(wfrom).Activate
limit = Worksheets("CONFETTI").Cells(15, 12).Value
Worksheets("CONFETTI").Range("A2:H" & limit).Copy
Workbooks(wto).Activate
Worksheets("CONFETTI").Range("A2:H" & limit).PasteSpecial xlPasteValues
Worksheets("CONFETTI").Range("A2:H" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("CONFETTI").Range("J3:L15").Copy
Workbooks(wto).Activate
Worksheets("CONFETTI").Range("J3:L15").PasteSpecial xlPasteValues
Worksheets("CONFETTI").Range("J3:L15").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied CONFETTI sheet from " & wfrom & " to " & wto

'print
Workbooks(wfrom).Activate
limit = Worksheets("PRINT").Cells(18, 13).Value
Worksheets("PRINT").Range("A2:I" & limit).Copy
Workbooks(wto).Activate
Worksheets("PRINT").Range("A2:I" & limit).PasteSpecial xlPasteValues
Worksheets("PRINT").Range("A2:I" & limit).PasteSpecial xlFormats
Workbooks(wfrom).Activate
Worksheets("PRINT").Range("K3:M20").Copy
Workbooks(wto).Activate
Worksheets("PRINT").Range("K3:M20").PasteSpecial xlPasteValues
Worksheets("PRINT").Range("K3:M20").PasteSpecial xlFormats
Print #TextFile, "Succesfully copied PRINT sheet from " & wfrom & " to " & wto

Workbooks(wfrom).Activate
Sheets("UPDATE").Select

MsgBox "Update process reached its end with no errors"

Print #TextFile, "UPDATE PROCESS FINISHED WITH CODE 1"
Print #TextFile, ""
Close TextFile
CreateObject("Shell.Application").Open (FilePath)

End Sub