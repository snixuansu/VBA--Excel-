Option Explicit



' ---------- НАСТРОЙКИ ----------
Const YEAR_START As Integer = 2026
Const FIRST_DATA_ROW As Integer = 17      ' Первая строка сотрудника
Const LAST_DATA_ROW As Integer = 21       ' Последняя строка сотрудника
Const COL_POSITION As Integer = 4         ' Столбец D — должность
Const COL_DAY1 As Integer = 5             ' Столбец E — день 1
Const COL_DAY15 As Integer = 19           ' Столбец S — день 15
Const COL_DAY16 As Integer = 21           ' Столбец U — день 16
Const COL_DAY31 As Integer = 36           ' Столбец AJ — день 31
Const COL_TOTAL_1 As Integer = 20         ' Столбец T — итог с 1 по 15
Const COL_TOTAL_MONTH As Integer = 37     ' Столбец AK — всего за месяц
Const ROW_PERIOD As Integer = 4           ' Строка "за период с ..."
Const COL_PERIOD As Integer = 9           ' Столбец I — текст периода
Const ROW_HEADER_DAYS As Integer = 14     ' Строка с числами месяца
' --------------------------------


' ============================================================
' ГЛАВНЫЙ МАКРОС — СОЗДАНИЕ 24 ЛИСТОВ ИЗ ШАБЛОНА
' ============================================================
Sub CreateYearFromTemplate()
    Dim templateName As String
    Dim templateSheet As Worksheet
    Dim newSheet As Worksheet
    Dim m As Integer, half As Integer
    Dim sheetName As String
    Dim startDay As Integer, endDay As Integer
    Dim strMonthName As String
    Dim monthNames As Variant
    Dim createdCount As Integer

    monthNames = Array("января", "февраля", "марта", "апреля", "мая", "июня", _
                       "июля", "августа", "сентября", "октября", "ноября", "декабря")

    templateName = InputBox("Введите имя листа-шаблона (с правильным форматированием):", _
                            "Шаблон", "Шаблон")
    If templateName = "" Then Exit Sub

    Set templateSheet = Nothing
    On Error Resume Next
    Set templateSheet = ThisWorkbook.Sheets(templateName)
    On Error GoTo 0

    If templateSheet Is Nothing Then
        MsgBox "Лист '" & templateName & "' не найден.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    createdCount = 0

    For m = 1 To 12
        For half = 1 To 2
            If half = 1 Then
                startDay = 1: endDay = 15
            Else
                startDay = 16
                endDay = Day(DateSerial(YEAR_START, m + 1, 0))
            End If

            sheetName = Format(m, "00") & "_" & YEAR_START & "_" & startDay & "-" & endDay

            Set newSheet = Nothing
            On Error Resume Next
            Set newSheet = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0

            If newSheet Is Nothing Then
                templateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                newSheet.Name = sheetName

                strMonthName = monthNames(m - 1)

                Call UpdatePeriodInfo(newSheet, m, startDay, endDay, strMonthName)
                Call UpdateHeaderDays(newSheet, m)
                Call ClearAllMarks(newSheet)
                Call RecalcWeekends(newSheet, m, startDay, endDay)
                Call UpdateFormulas(newSheet)

                createdCount = createdCount + 1
            End If
            Set newSheet = Nothing
        Next half
    Next m

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Готово! Создано листов: " & createdCount, vbInformation
End Sub


' ============================================================
' ОБНОВИТЬ ТЕКСТ ПЕРИОДА В ШАПКЕ
' ============================================================
Sub UpdatePeriodInfo(ws As Worksheet, m As Integer, startDay As Integer, endDay As Integer, strMonthName As String)
    On Error Resume Next
    ws.Cells(ROW_PERIOD, COL_PERIOD).Value = _
        "за период с " & startDay & " по " & endDay & " " & strMonthName & " " & YEAR_START & " г."
    On Error GoTo 0
End Sub


' ============================================================
' ОБНОВИТЬ ЧИСЛА МЕСЯЦА В ШАПКЕ (строка 14)
' Убирает лишние дни в коротких месяцах и добавляет 31-й
' ============================================================
Sub UpdateHeaderDays(ws As Worksheet, m As Integer)
    Dim i As Integer, col As Integer
    Dim lastDay As Integer

    lastDay = Day(DateSerial(YEAR_START, m + 1, 0))

    ' Первая половина: столбцы E–S, дни 1–15 (всегда есть)
    For i = 1 To 15
        col = COL_DAY1 + (i - 1)
        ws.Cells(ROW_HEADER_DAYS, col).Value = i
        ws.Cells(ROW_HEADER_DAYS, col).HorizontalAlignment = xlCenter
    Next i

    ' Вторая половина: столбцы U–AJ, дни 16–31
    For i = 16 To 31
        col = COL_DAY16 + (i - 16)
        If i <= lastDay Then
            ws.Cells(ROW_HEADER_DAYS, col).Value = i
            ws.Cells(ROW_HEADER_DAYS, col).HorizontalAlignment = xlCenter
        Else
            ' День не существует в этом месяце — убираем из шапки
            ws.Cells(ROW_HEADER_DAYS, col).ClearContents
        End If
    Next i
End Sub


' ============================================================
' ОЧИСТИТЬ СТАРЫЕ МЕТКИ И ЧИСЛА
' ============================================================
Sub ClearAllMarks(ws As Worksheet)
    Dim row As Integer, col As Integer
    Dim v As Variant

    For row = FIRST_DATA_ROW To LAST_DATA_ROW
        For col = COL_DAY1 To COL_TOTAL_MONTH
            If col <> COL_TOTAL_1 And col <> COL_TOTAL_MONTH Then
                v = ws.Cells(row, col).Value
                If IsNumeric(v) Or v = "В" Or v = "К" Or v = "О" Then
                    ws.Cells(row, col).ClearContents
                End If
            End If
        Next col
    Next row
End Sub


' ============================================================
' РАССТАВИТЬ ВЫХОДНЫЕ И ЧАСЫ ПО СТАВКЕ
' ============================================================
Sub RecalcWeekends(ws As Worksheet, m As Integer, startDay As Integer, endDay As Integer)
    Dim i As Integer, col As Integer, row As Integer
    Dim curDate As Date
    Dim isHoli As Boolean
    Dim isWeekend As Boolean
    Dim hours As Variant

    For i = startDay To endDay
        curDate = DateSerial(YEAR_START, m, i)
        isWeekend = (Weekday(curDate, vbMonday) > 5)
        isHoli = IsRussianHoliday(curDate)

        ' Определяем столбец: дни 1-15 → E-S, дни 16-31 → U-AJ
        If i <= 15 Then
            col = COL_DAY1 + (i - 1)
        Else
            col = COL_DAY16 + (i - 16)
        End If

        If isWeekend Or isHoli Then
            ' Выходной или праздник — ставим "В"
            For row = FIRST_DATA_ROW To LAST_DATA_ROW
                If ws.Cells(row, col).Value = "" Then
                    ws.Cells(row, col).Value = "В"
                    ws.Cells(row, col).HorizontalAlignment = xlCenter
                End If
            Next row
        Else
            ' Рабочий день — ставим часы по ставке
            For row = FIRST_DATA_ROW To LAST_DATA_ROW
                If ws.Cells(row, col).Value = "" Then
                    hours = GetHoursByRate(CStr(ws.Cells(row, COL_POSITION).Value))
                    ws.Cells(row, col).Value = hours
                    ws.Cells(row, col).HorizontalAlignment = xlCenter
                End If
            Next row
        End If
    Next i
End Sub


' ============================================================
' ОПРЕДЕЛИТЬ ЧАСЫ ПО СТАВКЕ
' ============================================================
Function GetHoursByRate(position As String) As Variant
    Dim p As String
    p = Replace(position, ".", ",")   ' на случай, если в ячейке точка вместо запятой

    If InStr(p, "0,25") > 0 Then
        GetHoursByRate = 2
    ElseIf InStr(p, "0,5") > 0 Then
        GetHoursByRate = 4
    Else
        GetHoursByRate = 8
    End If
End Function


' ============================================================
' ПРАЗДНИКИ УЧРЕЖДЕНИЯ
' ============================================================
Function IsRussianHoliday(d As Date) As Boolean
    Dim mm As Integer, dd As Integer
    mm = Month(d): dd = Day(d)

    Select Case mm
        Case 1
            ' с 1 по 11 января
            If dd >= 1 And dd <= 11 Then IsRussianHoliday = True
        Case 2
            ' 23 февраля
            If dd = 23 Then IsRussianHoliday = True
        Case 3
            ' 9 марта
            If dd = 9 Then IsRussianHoliday = True
        Case 5
            ' 1 и 11 мая
            If dd = 1 Or dd = 11 Then IsRussianHoliday = True
        Case 6
            ' 4 и 12 июня
            If dd = 4 Or dd = 12 Then IsRussianHoliday = True
        Case 12
            ' 31 декабря
            If dd = 31 Then IsRussianHoliday = True
    End Select
End Function


' ============================================================
' ОБНОВИТЬ ФОРМУЛЫ ИТОГОВ
' ============================================================
Sub UpdateFormulas(ws As Worksheet)
    Dim row As Integer
    Dim addr1 As String, addr2 As String
    Dim addr3 As String, addr4 As String

    addr1 = ws.Cells(FIRST_DATA_ROW, COL_DAY1).Address(False, False)
    addr2 = ws.Cells(FIRST_DATA_ROW, COL_DAY15).Address(False, False)
    addr3 = ws.Cells(FIRST_DATA_ROW, COL_DAY16).Address(False, False)
    addr4 = ws.Cells(FIRST_DATA_ROW, COL_DAY31).Address(False, False)

    For row = FIRST_DATA_ROW To LAST_DATA_ROW
        ' Итого с 1 по 15 — количество числовых ячеек
        ws.Cells(row, COL_TOTAL_1).Formula = _
            "=COUNT(" & Replace(addr1, FIRST_DATA_ROW, row) & ":" & _
                          Replace(addr2, FIRST_DATA_ROW, row) & ")"
        ' Всего за месяц — сумма по обеим половинам
        ws.Cells(row, COL_TOTAL_MONTH).Formula = _
            "=COUNT(" & Replace(addr1, FIRST_DATA_ROW, row) & ":" & _
                          Replace(addr2, FIRST_DATA_ROW, row) & ")+" & _
            "COUNT(" & Replace(addr3, FIRST_DATA_ROW, row) & ":" & _
                          Replace(addr4, FIRST_DATA_ROW, row) & ")"
    Next row
End Sub


' ============================================================
' ПЕРЕСЧИТАТЬ ИТОГИ ВО ВСЕХ ЛИСТАХ
' ============================================================
Sub RecalcAllTotals()
    Dim ws As Worksheet
    Dim m As Integer, half As Integer
    Dim sheetName As String
    Dim startDay As Integer, endDay As Integer

    Application.ScreenUpdating = False

    For m = 1 To 12
        For half = 1 To 2
            If half = 1 Then
                startDay = 1: endDay = 15
            Else
                startDay = 16: endDay = Day(DateSerial(YEAR_START, m + 1, 0))
            End If

            sheetName = Format(m, "00") & "_" & YEAR_START & "_" & startDay & "-" & endDay

            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0

            If Not ws Is Nothing Then
                Call UpdateHeaderDays(ws, m)
                Call UpdateFormulas(ws)
            End If
            Set ws = Nothing
        Next half
    Next m

    Application.ScreenUpdating = True
    MsgBox "Итоги пересчитаны на всех листах!", vbInformation
End Sub
