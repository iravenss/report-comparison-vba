Option Explicit

Sub CompareReportsHybrid()
    Dim wb1 As Workbook, wb2 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long
    Dim c1 As String, c2 As String
    Dim d1 As String, d2 As String
    Dim d1Num As Double, d2Num As Double
    Dim a1 As Double, a2 As Double
    Dim isDup1 As Boolean, isDup2 As Boolean
    
    Dim countMatch As Long, countDiff As Long, countNotFound As Long
    Dim countDupProcessed As Long
    countMatch = 0: countDiff = 0: countNotFound = 0: countDupProcessed = 0
    
    ' === НОВОЕ: счётчики дубликатов на каждом отчёте ===
    Dim dupCount1 As Long, dupCount2 As Long
    dupCount1 = 0: dupCount2 = 0
    
    ' Пакетное закрашивание
    Dim yellowRange1 As Range, yellowRange2 As Range, redRange As Range
    
    ' Лог
    Dim logFile As Integer
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\log_sverka.txt"
    logFile = FreeFile
    Open logPath For Output As #logFile
    Print #logFile, "=== ГИБРИДНАЯ СВЕРКА " & Now & " ==="
    
    ' Инициализация
    On Error Resume Next
    Set wb1 = Workbooks("Отчёт1.xlsx")
    Set wb2 = Workbooks("Отчёт2.xlsx")
    On Error GoTo 0
    
    If wb1 Is Nothing Or wb2 Is Nothing Then
        MsgBox "Откройте оба файла!", vbCritical
        Close #logFile
        Exit Sub
    End If
    
    Set ws1 = wb1.Sheets(1)
    Set ws2 = wb2.Sheets(1)
    
    ' Оптимизации
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ws1.EnableFormatConditionsCalculation = False
    ws2.EnableFormatConditionsCalculation = False
    ws1.DisplayPageBreaks = False
    ws2.DisplayPageBreaks = False
    
    lastRow1 = ws1.Cells(ws1.Rows.count, 2).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.count, 2).End(xlUp).Row
    Print #logFile, "Отчёт1: " & lastRow1 - 1 & " строк"
    Print #logFile, "Отчёт2: " & lastRow2 - 1 & " строк"
    
    Dim tTotal As Single: tTotal = Timer
    Dim tStep As Single
    
    ' ============================================
    ' ЭТАП 1: Создание ключей и пометка дубликатов
    ' ============================================
    tStep = Timer
    Application.StatusBar = "Этап 1/5: Создание ключей..."
    DoEvents
    
    ' Отчёт1: колонка F = ключ, колонка G = флаг дубликата
    With ws1
        For i = 2 To lastRow1
            .Cells(i, 6).Value = Trim(CStr(.Cells(i, 2).Value)) & "|" & Trim(CStr(.Cells(i, 3).Value))
        Next i
        
        .Range("A1:G" & lastRow1).Sort key1:=.Cells(1, 6), Order1:=xlAscending, Header:=xlYes
        
        For i = 2 To lastRow1
            If i = 2 Then
                .Cells(i, 7).Value = 0
            Else
                If .Cells(i, 6).Value = .Cells(i - 1, 6).Value Then
                    .Cells(i, 7).Value = 1
                    .Cells(i - 1, 7).Value = 1
                Else
                    .Cells(i, 7).Value = 0
                End If
            End If
        Next i
        
        ' === НОВОЕ: подсчёт и логирование дубликатов в Отчёте 1 ===
        Print #logFile, "----------------------------------------"
        Print #logFile, "ДУБЛИКАТЫ В ОТЧЁТЕ 1:"
        Dim prevKey1 As String: prevKey1 = ""
        For i = 2 To lastRow1
            If .Cells(i, 7).Value = 1 Then
                dupCount1 = dupCount1 + 1
                ' Логируем только при смене ключа (чтобы не повторять один и тот же дубликат много раз)
                If .Cells(i, 6).Value <> prevKey1 Then
                    Print #logFile, "  Ключ: " & .Cells(i, 6).Value
                    prevKey1 = .Cells(i, 6).Value
                End If
            End If
        Next i
        Print #logFile, "Итого уникальных дубликат-ключей: " & CountUniqueDupKeys(ws1, lastRow1)
        Print #logFile, "Итого строк-дубликатов: " & dupCount1
    End With
    
    ' Отчёт2: аналогично
    With ws2
        For i = 2 To lastRow2
            .Cells(i, 6).Value = Trim(CStr(.Cells(i, 2).Value)) & "|" & Trim(CStr(.Cells(i, 3).Value))
        Next i
        
        .Range("A1:G" & lastRow2).Sort key1:=.Cells(1, 6), Order1:=xlAscending, Header:=xlYes
        
        For i = 2 To lastRow2
            If i = 2 Then
                .Cells(i, 7).Value = 0
            Else
                If .Cells(i, 6).Value = .Cells(i - 1, 6).Value Then
                    .Cells(i, 7).Value = 1
                    .Cells(i - 1, 7).Value = 1
                Else
                    .Cells(i, 7).Value = 0
                End If
            End If
        Next i
        
        ' === НОВОЕ: подсчёт и логирование дубликатов в Отчёте 2 ===
        Print #logFile, "----------------------------------------"
        Print #logFile, "ДУБЛИКАТЫ В ОТЧЁТЕ 2:"
        Dim prevKey2 As String: prevKey2 = ""
        For i = 2 To lastRow2
            If .Cells(i, 7).Value = 1 Then
                dupCount2 = dupCount2 + 1
                If .Cells(i, 6).Value <> prevKey2 Then
                    Print #logFile, "  Ключ: " & .Cells(i, 6).Value
                    prevKey2 = .Cells(i, 6).Value
                End If
            End If
        Next i
        Print #logFile, "Итого уникальных дубликат-ключей: " & CountUniqueDupKeys(ws2, lastRow2)
        Print #logFile, "Итого строк-дубликатов: " & dupCount2
        Print #logFile, "----------------------------------------"
    End With
    
    Print #logFile, "Ключи и дубликаты: " & Round(Timer - tStep, 2) & " сек"
    
    ' ============================================
    ' ЭТАП 2: Сортировка по (Компания, Документ) для merge
    ' ============================================
    tStep = Timer
    Application.StatusBar = "Этап 2/5: Сортировка для merge..."
    DoEvents
    
    For i = 2 To lastRow1
        ws1.Cells(i, 5).Value = i
    Next i
    For i = 2 To lastRow2
        ws2.Cells(i, 5).Value = i
    Next i
    
    ws1.Range("A1:G" & lastRow1).Sort _
        key1:=ws1.Cells(1, 2), Order1:=xlAscending, _
        key2:=ws1.Cells(1, 3), Order2:=xlAscending, _
        Header:=xlYes
    
    ws2.Range("A1:G" & lastRow2).Sort _
        key1:=ws2.Cells(1, 2), Order1:=xlAscending, _
        key2:=ws2.Cells(1, 3), Order2:=xlAscending, _
        Header:=xlYes
    
    Print #logFile, "Сортировка: " & Round(Timer - tStep, 2) & " сек"
    
    ' ============================================
    ' ЭТАП 3: Merge для уникальных записей
    ' ============================================
    tStep = Timer
    Application.StatusBar = "Этап 3/5: Merge (уникальные)..."
    DoEvents
    
    i = 2: j = 2
    
    Do While i <= lastRow1 And j <= lastRow2
        
        If i Mod 1000 = 0 Then
            Application.StatusBar = "Merge: стр. " & i & "/" & lastRow1 & _
                " | Совпало: " & countMatch & " | Расхождений: " & countDiff
            DoEvents
        End If
        
        c1 = Trim(CStr(ws1.Cells(i, 2).Value))
        d1 = Trim(CStr(ws1.Cells(i, 3).Value))
        isDup1 = (ws1.Cells(i, 7).Value = 1)
        
        c2 = Trim(CStr(ws2.Cells(j, 2).Value))
        d2 = Trim(CStr(ws2.Cells(j, 3).Value))
        isDup2 = (ws2.Cells(j, 7).Value = 1)
        
        If c1 = c2 Then
            On Error Resume Next
            d1Num = CDbl(d1): d2Num = CDbl(d2)
            
            If Err.Number = 0 Then
                On Error GoTo 0
                If d1Num = d2Num Then
                    If isDup1 Or isDup2 Then
                        i = i + 1
                        j = j + 1
                    Else
                        a1 = ws1.Cells(i, 4).Value
                        a2 = ws2.Cells(j, 4).Value
                        
                        If Abs(a1 - a2) > 0.001 Then
                            If yellowRange1 Is Nothing Then
                                Set yellowRange1 = ws1.Cells(i, 4)
                            Else
                                Set yellowRange1 = Union(yellowRange1, ws1.Cells(i, 4))
                            End If
                            
                            If yellowRange2 Is Nothing Then
                                Set yellowRange2 = ws2.Cells(j, 4)
                            Else
                                Set yellowRange2 = Union(yellowRange2, ws2.Cells(j, 4))
                            End If
                            
                            countDiff = countDiff + 1
                            Print #logFile, "РАСХОЖДЕНИЕ: " & c1 & " Док=" & d1 & " Р1=" & a1 & " Р2=" & a2
                        Else
                            countMatch = countMatch + 1
                        End If
                        
                        i = i + 1
                        j = j + 1
                    End If
                ElseIf d1Num < d2Num Then
                    If Not isDup1 Then
                        If redRange Is Nothing Then
                            Set redRange = ws1.Cells(i, 3)
                        Else
                            Set redRange = Union(redRange, ws1.Cells(i, 3))
                        End If
                        countNotFound = countNotFound + 1
                        Print #logFile, "НЕ НАЙДЕН: " & c1 & " Док=" & d1
                    End If
                    i = i + 1
                Else
                    j = j + 1
                End If
            Else
                On Error GoTo 0
                If d1 = d2 Then
                    If isDup1 Or isDup2 Then
                        i = i + 1
                        j = j + 1
                    Else
                        a1 = ws1.Cells(i, 4).Value
                        a2 = ws2.Cells(j, 4).Value
                        
                        If Abs(a1 - a2) > 0.001 Then
                            If yellowRange1 Is Nothing Then
                                Set yellowRange1 = ws1.Cells(i, 4)
                            Else
                                Set yellowRange1 = Union(yellowRange1, ws1.Cells(i, 4))
                            End If
                            
                            If yellowRange2 Is Nothing Then
                                Set yellowRange2 = ws2.Cells(j, 4)
                            Else
                                Set yellowRange2 = Union(yellowRange2, ws2.Cells(j, 4))
                            End If
                            
                            countDiff = countDiff + 1
                            Print #logFile, "РАСХОЖДЕНИЕ: " & c1 & " Док=" & d1 & " Р1=" & a1 & " Р2=" & a2
                        Else
                            countMatch = countMatch + 1
                        End If
                        
                        i = i + 1
                        j = j + 1
                    End If
                ElseIf d1 < d2 Then
                    If Not isDup1 Then
                        If redRange Is Nothing Then
                            Set redRange = ws1.Cells(i, 3)
                        Else
                            Set redRange = Union(redRange, ws1.Cells(i, 3))
                        End If
                        countNotFound = countNotFound + 1
                        Print #logFile, "НЕ НАЙДЕН: " & c1 & " Док=" & d1
                    End If
                    i = i + 1
                Else
                    j = j + 1
                End If
            End If
        ElseIf c1 < c2 Then
            If Not isDup1 Then
                If redRange Is Nothing Then
                    Set redRange = ws1.Cells(i, 3)
                Else
                    Set redRange = Union(redRange, ws1.Cells(i, 3))
                End If
                countNotFound = countNotFound + 1
                Print #logFile, "НЕ НАЙДЕН: " & c1 & " Док=" & d1
            End If
            i = i + 1
        Else
            j = j + 1
        End If
    Loop
    
    Do While i <= lastRow1
        isDup1 = (ws1.Cells(i, 7).Value = 1)
        If Not isDup1 Then
            If redRange Is Nothing Then
                Set redRange = ws1.Cells(i, 3)
            Else
                Set redRange = Union(redRange, ws1.Cells(i, 3))
            End If
            countNotFound = countNotFound + 1
        End If
        i = i + 1
    Loop
    
    Print #logFile, "Merge: " & Round(Timer - tStep, 2) & " сек"
    
    ' ============================================
    ' ЭТАП 4: Find/FindNext для дубликатов
    ' ============================================
    tStep = Timer
    Application.StatusBar = "Этап 4/5: Обработка дубликатов..."
    DoEvents
    
    ws1.Range("A1:G" & lastRow1).Sort key1:=ws1.Cells(1, 5), Order1:=xlAscending, Header:=xlYes
    ws2.Range("A1:G" & lastRow2).Sort key1:=ws2.Cells(1, 5), Order1:=xlAscending, Header:=xlYes
    
    For i = 2 To lastRow1
        If ws1.Cells(i, 7).Value = 1 Then
            countDupProcessed = countDupProcessed + 1
            
            c1 = Trim(CStr(ws1.Cells(i, 2).Value))
            d1 = ws1.Cells(i, 3).Value
            a1 = ws1.Cells(i, 4).Value
            
            Dim foundCell As Range
            Set foundCell = ws2.Columns(3).Find(What:=d1, LookIn:=xlValues, LookAt:=xlWhole)
            
            Dim isFound As Boolean: isFound = False
            Dim firstAddress As String
            
            If Not foundCell Is Nothing Then
                firstAddress = foundCell.Address
                
                Do
                    c2 = Trim(CStr(ws2.Cells(foundCell.Row, 2).Value))
                    
                    If c2 = c1 Then
                        isFound = True
                        a2 = ws2.Cells(foundCell.Row, 4).Value
                        
                        If Abs(a1 - a2) > 0.001 Then
                            If yellowRange1 Is Nothing Then
                                Set yellowRange1 = ws1.Cells(i, 4)
                            Else
                                Set yellowRange1 = Union(yellowRange1, ws1.Cells(i, 4))
                            End If
                            
                            If yellowRange2 Is Nothing Then
                                Set yellowRange2 = ws2.Cells(foundCell.Row, 4)
                            Else
                                Set yellowRange2 = Union(yellowRange2, ws2.Cells(foundCell.Row, 4))
                            End If
                            
                            countDiff = countDiff + 1
                            Print #logFile, "РАСХОЖДЕНИЕ (дубликат): " & c1 & " Док=" & d1 & " Р1=" & a1 & " Р2=" & a2
                        Else
                            countMatch = countMatch + 1
                        End If
                        
                        Exit Do
                    End If
                    
                    Set foundCell = ws2.Columns(3).FindNext(foundCell)
                Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
            End If
            
            If Not isFound Then
                If redRange Is Nothing Then
                    Set redRange = ws1.Cells(i, 3)
                Else
                    Set redRange = Union(redRange, ws1.Cells(i, 3))
                End If
                countNotFound = countNotFound + 1
                Print #logFile, "НЕ НАЙДЕН (дубликат): " & c1 & " Док=" & d1
            End If
        End If
    Next i
    
    Print #logFile, "Дубликаты (" & countDupProcessed & " шт): " & Round(Timer - tStep, 2) & " сек"
    
    ' ============================================
    ' ЭТАП 5: Пакетное закрашивание и очистка
    ' ============================================
    tStep = Timer
    Application.StatusBar = "Этап 5/5: Закрашивание..."
    DoEvents
    
    If Not yellowRange1 Is Nothing Then yellowRange1.Interior.Color = vbYellow
    If Not yellowRange2 Is Nothing Then yellowRange2.Interior.Color = vbYellow
    If Not redRange Is Nothing Then redRange.Interior.Color = vbRed
    
    ws1.Columns("E:G").Clear
    ws2.Columns("E:G").Clear
    
    Print #logFile, "Закрашивание: " & Round(Timer - tStep, 2) & " сек"
    
    ' Возврат настроек
    ws1.DisplayPageBreaks = True
    ws2.DisplayPageBreaks = True
    ws1.EnableFormatConditionsCalculation = True
    ws2.EnableFormatConditionsCalculation = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' ============================================
    ' ИТОГОВЫЙ ЛОГ (с информацией о дубликатах)
    ' ============================================
    Print #logFile, "================================"
    Print #logFile, "ИТОГО: " & Round(Timer - tTotal, 2) & " сек"
    Print #logFile, "Совпало: " & countMatch
    Print #logFile, "Расхождений: " & countDiff
    Print #logFile, "Не найдено: " & countNotFound
    Print #logFile, "--------------------------------"
    Print #logFile, "ИНФОРМАЦИЯ О ДУБЛИКАТАХ:"
    Print #logFile, "  Отчёт1 строк-дубликатов: " & dupCount1
    Print #logFile, "  Отчёт2 строк-дубликатов: " & dupCount2
    Print #logFile, "  Обработано через Find: " & countDupProcessed
    Print #logFile, "================================"
    Close #logFile
    
    MsgBox "Сверка завершена за " & Round(Timer - tTotal, 1) & " сек!" & _
        vbCrLf & vbCrLf & _
        "Совпало: " & countMatch & vbCrLf & _
        "Расхождений (жёлтым): " & countDiff & vbCrLf & _
        "Не найдено (красным): " & countNotFound & vbCrLf & vbCrLf & _
        "Дубликатов в Отчёт1: " & dupCount1 & vbCrLf & _
        "Дубликатов в Отчёт2: " & dupCount2 & vbCrLf & _
        "Обработано через Find: " & countDupProcessed, vbInformation
End Sub

' ============================================
' Вспомогательная функция: подсчёт уникальных дубликат-ключей
' ============================================
Private Function CountUniqueDupKeys(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim i As Long
    Dim count As Long
    Dim prevKey As String
    
    count = 0
    prevKey = ""
    
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value = 1 Then
            If ws.Cells(i, 6).Value <> prevKey Then
                count = count + 1
                prevKey = ws.Cells(i, 6).Value
            End If
        End If
    Next i
    
    CountUniqueDupKeys = count
End Function
