Option Explicit

Function MakeKey(ws As Worksheet, r As Long) As String
    MakeKey = Trim(CStr(ws.Cells(r, 2).Value)) & "|" & Trim(CStr(ws.Cells(r, 3).Value))
End Function

Function Num(c As Range) As Double
    If IsNumeric(c.Value2) Then Num = CDbl(c.Value2)
End Function

Sub CompareReports()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim d1 As Object, d2 As Object, bad As Object
    Dim i As Long, last1 As Long, last2 As Long
    Dim k As String, v As Variant
    Dim cMatch As Long, cDiff As Long, cMiss As Long

    On Error Resume Next
    Set ws1 = Workbooks("Отчёт1.xlsx").Sheets(1)
    Set ws2 = Workbooks("Отчёт2.xlsx").Sheets(1)
    On Error GoTo 0
    If ws1 Is Nothing Or ws2 Is Nothing Then MsgBox "Откройте оба файла!": Exit Sub

    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = CreateObject("Scripting.Dictionary")
    Set bad = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    last1 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
    last2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row

   
    ws1.Range(ws1.Cells(2, 3), ws1.Cells(last1, 4)).Interior.Pattern = xlNone
    ws2.Range(ws2.Cells(2, 3), ws2.Cells(last2, 4)).Interior.Pattern = xlNone

  
    For i = 2 To last1
        k = MakeKey(ws1, i)
        d1(k) = d1(k) + Num(ws1.Cells(i, 4))
    Next i
    For i = 2 To last2
        k = MakeKey(ws2, i)
        d2(k) = d2(k) + Num(ws2.Cells(i, 4))
    Next i


    For Each v In d1.Keys
        If d2.Exists(v) Then
            If Abs(d1(v) - d2(v)) > 0.005 Then bad(v) = 1
        End If
    Next v


    For i = 2 To last1
        k = MakeKey(ws1, i)
        If Not d2.Exists(k) Then
            ws1.Cells(i, 3).Interior.Color = vbRed: cMiss = cMiss + 1
        ElseIf bad.Exists(k) Then
            ws1.Cells(i, 4).Interior.Color = vbYellow: cDiff = cDiff + 1
        Else
            cMatch = cMatch + 1
        End If
    Next i


    For i = 2 To last2
        If bad.Exists(MakeKey(ws2, i)) Then ws2.Cells(i, 4).Interior.Color = vbYellow
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Совпало: " & cMatch & " | Расхождений: " & cDiff & " | Не найдено: " & cMiss
End Sub
