Attribute VB_Name = "Module1"

' =====================================================================
' WB_nalog_USN_NDS ver3.05  |  21.04.2026
' =====================================================================
' ТИПЫ ОТЧЁТОВ WILDBERRIES:
'   "Основной"    — содержит Продажи + Возвраты + Логистику + Хранение
'                   Признак: в "Обоснование для оплаты" есть строки
'                   "Возмещение", "Хранение" или "Удержание"
'   "По выкупам"  — содержит ТОЛЬКО выкупленные товары (Продажи + Логистика)
'                   Признак: в "Обоснование для оплаты" только
'                   "Логистика" и "Продажа" (нет Возмещения/Хранения/Удержания)
'
'
' ФОРМУЛЫ — "Основной":
'   Выкупы          = SUM("Вайлдберриз реализовал Товар (Пр)")
'                     WHERE "Тип документа" = "Продажа"
'   Возвраты        = SUM("Вайлдберриз реализовал Товар (Пр)")
'                     WHERE "Тип документа" = "Возврат"
'   Компенсации     = SUM("Компенсация скидки по программе лояльности")
'                     WHERE "Тип документа" = "Продажа"  (+)
'                     WHERE "Тип документа" = "Возврат"  (?)
'                   + SUM("К перечислению Продавцу за реализованный Товар")
'                     WHERE "Обоснование" = "Добровольная компенсация при возврате" (+)
'   Доход базовый   = Выкупы - Возвраты + Компенсации
'
' ФОРМУЛЫ — "По выкупам":
'   К перечислению  = SUM("К перечислению Продавцу за реализованный Товар")
'                     (строки с Обоснование = "Продажа")
'   Услуги доставки = SUM("Услуги по доставке товара покупателю")
'                     (строки с Обоснование = "Логистика")
'   Компенсация     = SUM("Компенсация скидки по программе лояльности")
'   Доход базовый   = К перечислению ? Услуги доставки + Компенсация
'
' РАСЧЁТ НДС (УСН + НДС):
'   Ставка НДС берётся из листа "Настройки" ячейка B3 (0 или 5).
'   При ставке 0% — НДС не начисляется.
'   При ставке 5%:
'     Порог дохода (нарастающим итогом) = 20 000 000 руб.
'     Как только накопленный доход >= порога — НДС вступает в силу
'     с 1-го числа СЛЕДУЮЩЕГО месяца после превышения.
'     Сумма НДС    = Доход базовый * НДС% / (100 + НДС%)
'                    (НДС выделяется из суммы "в т.ч. НДС")
'     Доход без НДС = Доход базовый - Сумма НДС
'     Налогооблагаемый доход УСН = Доход без НДС
'
' ЛИСТ "НАСТРОЙКИ":
'   B2 — путь к папке с отчётами
'   B3 — ставка НДС до лимита (%) (число от 0 до ,05)
'   B4 — Ставка НДС после лимита (%) (по умолчанию 0,05)
'   B5 — Сумма дохода за прошлый год (вводится вручную)
'   B6 — Лимит дохода за прошлый год (по умолчанию 60 млн. р.)
'   B7 — Лимит дохода за текущий год (по умолчанию 20 млн. р.)
'
' КАК ИМПОРТИРОВАТЬ:
'   1. Откройте WB_nalog_USN_NDS.xlsx в Excel
'   2. Alt+F11 > Редактор VBA
'   3. Файл > Импорт файла > выберите Module1.bas
'   4. Удалите старый модуль (ПКМ > Remove Module)
'   5. Закройте редактор, сохраните файл как .xlsm
' =====================================================================

' =====================================================================
' Вспомогательная функция: определяет тип отчёта WB
' Возвращает: "Основной" или "По выкупам"
'
' Алгоритм:
'   Сканирует столбец "Обоснование для оплаты".
'   Если найдена хотя бы одна строка, содержащая:
'     "Возмещение" / "Хранение" / "Удержание"
'   > тип = "Основной"
'   (эти значения присутствуют только в полном отчёте)
'   Иначе > тип = "По выкупам"
'   (в "По выкупам" "Обоснование" принимает только "Логистика" и "Продажа")
' =====================================================================
Function DetectReportType(ws As Worksheet) As String

    Dim colOsn  As Long
    Dim j       As Long
    Dim hdr     As String

    colOsn = 0

    ' Ищем столбец "Обоснование для оплаты"
    For j = 1 To 100
        hdr = Trim(CStr(ws.Cells(1, j).Value))
        If InStr(1, hdr, "Обоснование для оплаты", vbTextCompare) > 0 Then
            colOsn = j
            Exit For
        End If
    Next j

    ' Если столбец не найден — считаем "Основной" (безопаснее)
    If colOsn = 0 Then
        DetectReportType = "Основной"
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i    As Long
    Dim vOsn As String

    For i = 2 To lastRow
        vOsn = Trim(CStr(ws.Cells(i, colOsn).Value))

        ' Признак "Основного": есть Возмещение / Хранение / Удержание в Обоснование для оплаты.
        ' Эти строки присутствуют только в полном отчёте ("Основной") и отсутствуют
        ' в "По выкупам", где Обоснование принимает только значения "Логистика" и "Продажа".
        If InStr(1, vOsn, "Возмещение", vbTextCompare) > 0 Then
            DetectReportType = "Основной"
            Exit Function
        End If
        If InStr(1, vOsn, "Хранение", vbTextCompare) > 0 Then
            DetectReportType = "Основной"
            Exit Function
        End If
        If InStr(1, vOsn, "Удержание", vbTextCompare) > 0 Then
            DetectReportType = "Основной"
            Exit Function
        End If
    Next i

    ' Ни одного признака "Основного" не найдено — это "По выкупам"
    DetectReportType = "По выкупам"

End Function

' =====================================================================
' Вспомогательная функция: извлекает номер отчёта из имени файла WB.
' Формат имени: "Еженедельный_детализированный_отчет__595135893_250049872_-_1.xlsx"
' Возвращает первую числовую последовательность после двойного подчёркивания "__".
' =====================================================================
Function ExtractReportNumber(FileName As String) As String
    Dim pos As Long
    pos = InStr(FileName, "№")
    If pos = 0 Then
        ExtractReportNumber = FileName
        Exit Function
    End If
    Dim rest As String
    rest = Mid(FileName, pos + 1)
    Dim result As String
    result = ""
    Dim k As Long
    For k = 1 To Len(rest)
        Dim ch As String
        ch = Mid(rest, k, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            Exit For
        End If
    Next k
    ExtractReportNumber = result
End Function

' =====================================================================
' Вспомогательная функция: поиск точного дня пересечения порога НДС
' Входные данные:
'   dailySerial — строка "YYYY-MM-DD~base|YYYY-MM-DD~base|..." (col M в Данные)
'   cumulBefore — накопленный доход ДО данного отчёта (сумма G предыдущих строк)
'   limit       — порог (Настройки!B7, по умолчанию 20 млн)
' Результат:
'   первый день (по возрастанию), в который cumulBefore + cumDaily >= limit;
'   0 (Date 30.12.1899), если даты нечитаемы или строка пуста => fallback в вызывающей.
' =====================================================================
Function FindDayOfExceed(dailySerial As String, cumulBefore As Double, LIMIT As Double) As Date
    FindDayOfExceed = 0
    If Len(dailySerial) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(dailySerial, "|")
    If UBound(parts) < 0 Then Exit Function

    Dim days()  As Date
    Dim bases() As Double
    ReDim days(0 To UBound(parts))
    ReDim bases(0 To UBound(parts))

    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To UBound(parts)
        If parts(i) = "" Then GoTo SkipDailyPart
        Dim flds() As String
        flds = Split(parts(i), "~")
        If UBound(flds) < 1 Then GoTo SkipDailyPart
        If Not IsDate(flds(0)) Then GoTo SkipDailyPart
        days(n) = CDate(flds(0))
        bases(n) = val(flds(1))
        n = n + 1
SkipDailyPart:
    Next i

    If n = 0 Then Exit Function

    ' Пузырьковая сортировка по дате возрастающе (n обычно < 50)
    Dim j As Long, tmpD As Date, tmpB As Double
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If days(j) > days(j + 1) Then
                tmpD = days(j): days(j) = days(j + 1): days(j + 1) = tmpD
                tmpB = bases(j): bases(j) = bases(j + 1): bases(j + 1) = tmpB
            End If
        Next j
    Next i

    Dim acc As Double: acc = cumulBefore
    For i = 0 To n - 1
        acc = acc + bases(i)
        If acc >= LIMIT Then
            FindDayOfExceed = days(i)
            Exit Function
        End If
    Next i
End Function

' =====================================================================
' Основной макрос обработки отчётов
' =====================================================================
Sub WB_ULTRA_FINAL()

    Dim wsData As Worksheet
    Dim wsCtrl As Worksheet
    Dim wsSet  As Worksheet
    Dim wsSvod As Worksheet
    Dim wsCal As Worksheet

    Set wsData = ThisWorkbook.Sheets("Данные")
    Set wsCtrl = ThisWorkbook.Sheets("Контроль")
    Set wsSet = ThisWorkbook.Sheets("Настройки")
    Set wsSvod = ThisWorkbook.Sheets("Свод")
    Set wsCal = ThisWorkbook.Sheets("Календарь")

    ' ========================
    ' Читаем настройки
    ' ========================
    Dim FolderPath As String
    FolderPath = Trim(CStr(wsSet.Cells(2, 2).Value))
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"

    ' Ставка НДС (0 или 5 — число, не процент)
    ' Ставка НДС до достижения порога 20 млн (B3)
    ' Вводить как число: 5 означает 5%.
    ' Если ячейка отформатирована как «%» и хранит 0,05 — конвертируем автоматически.
    Dim ndsRate1 As Double
    ndsRate1 = 0
    If IsNumeric(wsSet.Cells(3, 2).Value) Then
        ndsRate1 = CDbl(wsSet.Cells(3, 2).Value)
        If ndsRate1 > 0 And ndsRate1 < 1 Then ndsRate1 = ndsRate1 * 100
    End If
    If ndsRate1 < 0 Then ndsRate1 = 0
    If ndsRate1 > 20 Then ndsRate1 = 20

    ' Ставка НДС после достижения порога 20 млн (B4)
    ' Вводить как число: 5 означает 5%.
    ' Если ячейка отформатирована как «%» и хранит 0,05 — конвертируем автоматически.
    Dim ndsRate2 As Double
    ndsRate2 = 0
    If IsNumeric(wsSet.Cells(4, 2).Value) Then
        ndsRate2 = CDbl(wsSet.Cells(4, 2).Value)
        If ndsRate2 > 0 And ndsRate2 < 1 Then ndsRate2 = ndsRate2 * 100
    End If
    If ndsRate2 < 0 Then ndsRate2 = 0
    If ndsRate2 > 20 Then ndsRate2 = 20

    ' Доход за прошлый год (B5, вводится вручную)
    Dim prevYearIncome As Double
    prevYearIncome = 0
    If IsNumeric(wsSet.Cells(5, 2).Value) Then
        prevYearIncome = CDbl(wsSet.Cells(5, 2).Value)
    End If

    ' Лимит прошлого года (B6, константа 60 млн) — защита от изменений
    Dim LIMIT_PREV As Double
    LIMIT_PREV = 60000000
    wsSet.Cells(6, 2).Value = LIMIT_PREV

    ' Лимит текущего года (B7, константа 20 млн) — защита от изменений
    Dim LIMIT As Double
    LIMIT = 20000000
    wsSet.Cells(7, 2).Value = LIMIT

    ' Если доход прошлого года превысил лимит 60 млн —
    ' НДС 5% применяется с начала текущего года (ставка до лимита = ndsRate2)
    Dim prevYearExceeded As Boolean
    prevYearExceeded = (prevYearIncome >= LIMIT_PREV)
    If prevYearExceeded Then
        ndsRate1 = ndsRate2   ' до лимита та же ставка, что и после
        wsSet.Cells(3, 2).Value = ndsRate2 / 100  ' обновляем B3 для наглядности
    End If

    ' ========================
    ' Заголовки листа "Данные"
    ' ========================
    wsData.Cells(1, 1).Value = "Дата формирования отчета"
    wsData.Cells(1, 2).Value = "Дата начала"
    wsData.Cells(1, 3).Value = "Дата конца"
    wsData.Cells(1, 4).Value = "Выкупы / К перечислению ,р."
    wsData.Cells(1, 5).Value = "Возвраты / Услуги доставки ,р."
    wsData.Cells(1, 6).Value = "Компенсации ,р."
    wsData.Cells(1, 7).Value = "Доход с НДС ,р."
    wsData.Cells(1, 8).Value = "Сумма НДС ,р."
    wsData.Cells(1, 9).Value = "Доход без НДС ,р."
    wsData.Cells(1, 10).Value = "Тип отчёта"
    wsData.Cells(1, 11).Value = "Номер отчёта"
    wsData.Cells(1, 13).Value = "Разбивка по дням"

    ' ========================
    ' Заголовки листа "Свод"
    ' ========================
    wsSvod.Cells(1, 1).Value = "Месяц"
    wsSvod.Cells(1, 2).Value = "Выкупы / К перечислению ,р."
    wsSvod.Cells(1, 3).Value = "Возвраты / Услуги доставки ,р."
    wsSvod.Cells(1, 4).Value = "Компенсации ,р."
    wsSvod.Cells(1, 5).Value = "Доход с НДС ,р."
    wsSvod.Cells(1, 6).Value = "Сумма НДС ,р."
    wsSvod.Cells(1, 7).Value = "Доход без НДС ,р."
    wsSvod.Cells(1, 8).Value = "Ставка НДС, %"
    wsSvod.Cells(1, 9).Value = "Отчетов Основной"
    wsSvod.Cells(1, 10).Value = "Отчетов По выкупам"

    ' ========================
    ' Заголовки листа "Контроль"
    ' ========================
    wsCtrl.Cells(1, 1).Value = "Показатель"
    wsCtrl.Cells(1, 2).Value = "Значение"
    wsCtrl.Cells(2, 1).Value = "Доход с НДС накопительным итогом ,р."
    wsCtrl.Cells(3, 1).Value = "Сумма НДС накопительным итогом ,р."
    wsCtrl.Cells(4, 1).Value = "Доход без НДС накопительным итогом ,р."
    wsCtrl.Cells(5, 1).Value = "Превышен порог освобождения от уплаты НДС?"
    wsCtrl.Cells(6, 1).Value = "Дата превышения порога"
    wsCtrl.Cells(7, 1).Value = "НДС с 1-го числа след. мес. (" & Format(ndsRate2, "0") & "%)"
    wsCtrl.Cells(8, 1).Value = "Файлов обработано (всего)"
    wsCtrl.Cells(9, 1).Value = "Файлов типа Основной"
    wsCtrl.Cells(10, 1).Value = "Файлов типа По выкупам"

    ' ========================
    ' Заголовки листа "Настройки" (строки 2–7)
    ' ========================
    wsSet.Cells(2, 1).Value = "Путь к папке с отчётами"
    wsSet.Cells(3, 1).Value = "Ставка НДС до лимита ,%"
    wsSet.Cells(4, 1).Value = "Ставка НДС после лимита ,%"
    wsSet.Cells(5, 1).Value = "Сумма дохода за прошлый год ,р."
    wsSet.Cells(6, 1).Value = "Порог дохода за прошлый год (НДС) ,р."
    wsSet.Cells(7, 1).Value = "Порог дохода за текущий год (НДС) ,р."
    wsSet.Cells(8, 1).Value = "Ставка налога по УСН ,%"
    wsSet.Cells(9, 1).Value = "Кап 1% взноса ,р."
    wsSet.Cells(10, 1).Value = "Сумма дополнительного страхового взноса за 2025 ,р."
    wsSet.Cells(11, 1).Value = "Сумма фиксированных страховых взносов за 2025 ,р."
    If prevYearExceeded Then
        wsSet.Cells(3, 1).Value = "Ставка НДС до порога (%) [авто: прошл. год > 60 млн]"
    End If

    ' ========================
    ' Очистка предыдущих данных
    ' ========================
    Dim lastOld As Long
    lastOld = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastOld > 1 Then wsData.Rows("2:" & lastOld).Delete
    
        ' 3. Заголовки.
    wsCal.Cells(1, 1).Value = "Мероприятие"
    wsCal.Cells(1, 2).Value = "Плановый срок"
    wsCal.Cells(1, 3).Value = "Сумма, р."

    ' 4. Данные — 26 событий, отсортированы хронологически по дате.
    '    Каждая пара на отдельной строке (без line continuations).
    wsCal.Cells(2, 1).Value = "Оплата фиксированных страховых взносов за 2025": wsCal.Cells(2, 2).Value = DateSerial(2026, 1, 9)
    wsCal.Cells(3, 1).Value = "Подача декларации НДС за 4 кв 2025":             wsCal.Cells(3, 2).Value = DateSerial(2026, 1, 25)
    wsCal.Cells(4, 1).Value = "Оплата НДС 1/3 за 4 кв 2025":                    wsCal.Cells(4, 2).Value = DateSerial(2026, 1, 28)
    wsCal.Cells(5, 1).Value = "Оплата НДС 2/3 за 4 кв 2025":                    wsCal.Cells(5, 2).Value = DateSerial(2026, 2, 28)
    wsCal.Cells(6, 1).Value = "Оплата НДС 3/3 за 4 кв 2025":                    wsCal.Cells(6, 2).Value = DateSerial(2026, 3, 28)
    wsCal.Cells(7, 1).Value = "Подача отчетности РПН":                          wsCal.Cells(7, 2).Value = DateSerial(2026, 4, 14)
    wsCal.Cells(8, 1).Value = "Оплата экологического сбора":                    wsCal.Cells(8, 2).Value = DateSerial(2026, 4, 15)
    wsCal.Cells(9, 1).Value = "Подача декларации НДС за 1 кв 2026":             wsCal.Cells(9, 2).Value = DateSerial(2026, 4, 25)
    wsCal.Cells(10, 1).Value = "Оплата аванса УСН 1 квартал":                   wsCal.Cells(10, 2).Value = DateSerial(2026, 4, 28)
    wsCal.Cells(11, 1).Value = "Оплата НДС 1/3 за 1 кв 2026":                   wsCal.Cells(11, 2).Value = DateSerial(2026, 4, 28)
    wsCal.Cells(12, 1).Value = "Подача декларации УСН за 2025":                 wsCal.Cells(12, 2).Value = DateSerial(2026, 4, 27)
    wsCal.Cells(13, 1).Value = "Оплата налога УСН за 2025":                     wsCal.Cells(13, 2).Value = DateSerial(2026, 4, 28)
    wsCal.Cells(14, 1).Value = "Оплата НДС 2/3 за 1 кв 2026":                   wsCal.Cells(14, 2).Value = DateSerial(2026, 5, 28)
    wsCal.Cells(15, 1).Value = "Оплата НДС 3/3 за 1 кв 2026":                   wsCal.Cells(15, 2).Value = DateSerial(2026, 6, 28)
    wsCal.Cells(16, 1).Value = "Оплата дополнительного страхового взноса 1%":   wsCal.Cells(16, 2).Value = DateSerial(2026, 7, 1)
    wsCal.Cells(17, 1).Value = "Подача декларации НДС за 2 кв 2026":            wsCal.Cells(17, 2).Value = DateSerial(2026, 7, 25)
    wsCal.Cells(18, 1).Value = "Оплата аванса УСН полугодие":                   wsCal.Cells(18, 2).Value = DateSerial(2026, 7, 28)
    wsCal.Cells(19, 1).Value = "Оплата НДС 1/3 за 2 кв 2026":                   wsCal.Cells(19, 2).Value = DateSerial(2026, 7, 28)
    wsCal.Cells(20, 1).Value = "Оплата НДС 2/3 за 2 кв 2026":                   wsCal.Cells(20, 2).Value = DateSerial(2026, 8, 28)
    wsCal.Cells(21, 1).Value = "Оплата НДС 3/3 за 2 кв 2026":                   wsCal.Cells(21, 2).Value = DateSerial(2026, 9, 28)
    wsCal.Cells(22, 1).Value = "Подача декларации НДС за 3 кв 2026":            wsCal.Cells(22, 2).Value = DateSerial(2026, 10, 25)
    wsCal.Cells(23, 1).Value = "Оплата аванса УСН 9 месяцев":                   wsCal.Cells(23, 2).Value = DateSerial(2026, 10, 28)
    wsCal.Cells(24, 1).Value = "Оплата НДС 1/3 за 3 кв 2026":                   wsCal.Cells(24, 2).Value = DateSerial(2026, 10, 28)
    wsCal.Cells(25, 1).Value = "Оплата НДС 2/3 за 3 кв 2026":                   wsCal.Cells(25, 2).Value = DateSerial(2026, 11, 28)
    wsCal.Cells(26, 1).Value = "Оплата НДС 3/3 за 3 кв 2026":                   wsCal.Cells(26, 2).Value = DateSerial(2026, 12, 28)
    wsCal.Cells(27, 1).Value = "Оплата фиксированных страховых взносов за 2026": wsCal.Cells(27, 2).Value = DateSerial(2026, 12, 31)

    Application.ScreenUpdating = False

    Dim FileName  As String
    Dim wb        As Workbook
    Dim ws        As Worksheet
    Dim cntTotal  As Long
    Dim cntOsn    As Long
    Dim cntVykup  As Long

    cntTotal = 0: cntOsn = 0: cntVykup = 0

    FileName = Dir(FolderPath & "*.xlsx")

    Do While FileName <> ""

        ' Пропуск lock-файлов Excel (~$…xlsx), которые Dir() тоже возвращает.
        If Left(FileName, 2) = "~$" Then
            FileName = Dir
            GoTo NextFile
        End If

        Set wb = Workbooks.Open(FolderPath & FileName)
        Set ws = wb.Sheets(1)

        ' ========================
        ' Определяем тип отчёта
        ' ========================
        Dim reportType As String
        reportType = DetectReportType(ws)

        ' ========================
        ' Поиск нужных столбцов
        ' (общие для обоих типов отчётов)
        ' ========================
        Dim colDoc  As Long   ' "Тип документа"
        Dim colArt  As Long   ' "Артикул поставщика"
        Dim colSale As Long   ' "Вайлдберриз реализовал Товар (Пр)"
        Dim colComp As Long   ' "Компенсация скидки по программе лояльности"
        Dim colPay  As Long   ' "К перечислению Продавцу за реализованный Товар"
        Dim colDel  As Long   ' "Услуги по доставке товара покупателю"
        Dim colOsn  As Long   ' "Обоснование для оплаты"
        Dim colVozm   As Long  ' "Возмещение издержек по перевозке/по складским операциям с товаром"
        Dim colStore  As Long  ' "Хранение"
        Dim colDeduct As Long  ' "Удержания"

        colDoc = 0: colArt = 0: colSale = 0: colComp = 0: colPay = 0: colDel = 0: colOsn = 0
        colVozm = 0: colStore = 0: colDeduct = 0

        Dim j As Long
        For j = 1 To 100
            Dim hdr As String
            hdr = Trim(CStr(ws.Cells(1, j).Value))

            If InStr(1, hdr, "Тип документа", vbTextCompare) > 0 Then
                colDoc = j
            End If
            If InStr(1, hdr, "Артикул поставщика", vbTextCompare) > 0 Then
                colArt = j
            End If
            If InStr(1, hdr, "Вайлдберриз реализовал", vbTextCompare) > 0 Then
                colSale = j
            End If
            If InStr(1, hdr, "Компенсация скидки по программе лояльности", vbTextCompare) > 0 Then
                colComp = j
            End If
            If InStr(1, hdr, "К перечислению Продавцу", vbTextCompare) > 0 Then
                colPay = j
            End If
            If InStr(1, hdr, "Услуги по доставке товара покупателю", vbTextCompare) > 0 Then
                colDel = j
            End If
            If InStr(1, hdr, "Обоснование для оплаты", vbTextCompare) > 0 Then
                colOsn = j
            End If
            If InStr(1, hdr, "Возмещение издержек", vbTextCompare) > 0 Then
                colVozm = j
            End If
            ' "Хранение" и "Удержания" — ищем как отдельные заголовки колонок сумм.
            ' vbTextCompare делает поиск регистронезависимым. Осторожно: "Хранение"
            ' может попасть в более длинный заголовок, но в реальных отчётах WB
            ' этот заголовок уникален.
            If Trim(hdr) = "Хранение" Then
                colStore = j
            End If
            If Trim(hdr) = "Удержания" Then
                colDeduct = j
            End If
        Next j

        ' ========================
        ' Расчёт по строкам с разбивкой по месяцам
        ' Один отчёт может содержать строки из разных месяцев (например, февраль
        ' и март). Записываем в «Данные» отдельную строку на каждый месяц — это
        ' позволяет корректно применить разные ставки НДС к каждому периоду.
        ' ========================
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Ищем столбец «Дата продажи» заранее — нужен для группировки по месяцам
        Dim colDate As Long
        colDate = 0
        Dim jd As Long
        For jd = 1 To 20
            If InStr(1, Trim(CStr(ws.Cells(1, jd).Value)), "Дата продажи", vbTextCompare) > 0 Then
                colDate = jd
                Exit For
            End If
        Next jd

        ' Словари для накопления сумм по месяцам: ключ = "YYYY-MM"
        Dim dictVyk    As Object  ' Выкупы / К перечислению
        Dim dictVoz    As Object  ' Возвраты / Логистика
        Dim dictKmp    As Object  ' Компенсации
        Dim dictDt     As Object  ' Первая дата строки в этом месяце
        Dim dictMinDt  As Object  ' Минимальная Дата продажи в месяце
        Dim dictMaxDt  As Object  ' Максимальная Дата продажи в месяце
        Set dictVyk = CreateObject("Scripting.Dictionary")
        Set dictVoz = CreateObject("Scripting.Dictionary")
        Set dictKmp = CreateObject("Scripting.Dictionary")
        Set dictDt = CreateObject("Scripting.Dictionary")
        Set dictMinDt = CreateObject("Scripting.Dictionary")
        Set dictMaxDt = CreateObject("Scripting.Dictionary")

        ' Дневные словари (ключ "YYYY-MM-DD") — для точного дня превышения порога.
        Dim dictDailyVyk As Object
        Dim dictDailyVoz As Object
        Dim dictDailyKmp As Object
        Set dictDailyVyk = CreateObject("Scripting.Dictionary")
        Set dictDailyVoz = CreateObject("Scripting.Dictionary")
        Set dictDailyKmp = CreateObject("Scripting.Dictionary")

        Dim i As Long

        If reportType = "По выкупам" Then

            ' ---------------------------------------------------------
            ' АЛГОРИТМ "По выкупам" (с разбивкой по месяцам)
            '   Шаг 1: Артикулы текущего периода = уникальные артикулы из строк «Продажа»
            '   Шаг 2: Для каждого месяца: Доход = SUM(col34 WHERE Тип = «Продажа»)
            '   Шаг 3: Для каждого месяца: Логистика = SUM(col37 WHERE Обоснование = «Логистика»
            '          AND Артикул IN {артикулы из Шага 1})
            ' ---------------------------------------------------------
            If colPay = 0 Or colDel = 0 Or colOsn = 0 Or colArt = 0 Then
                MsgBox "Файл: " & FileName & Chr(13) & _
                       "Тип: По выкупам" & Chr(13) & _
                       "Не найден обязательный столбец." & Chr(13) & _
                       "Файл пропущен.", vbExclamation
                wb.Close False
                FileName = Dir
                GoTo NextFile
            End If

            ' Шаг 1: Собираем артикулы текущего периода
            Dim dictArts As Object
            Set dictArts = CreateObject("Scripting.Dictionary")
            For i = 2 To lastRow
                If Trim(CStr(ws.Cells(i, colDoc).Value)) = "Продажа" Then
                    Dim artKey As String
                    artKey = Trim(CStr(ws.Cells(i, colArt).Value))
                    If artKey <> "" And Not dictArts.Exists(artKey) Then
                        dictArts.Add artKey, 1
                    End If
                End If
            Next i

            ' Шаг 2+3: Группируем по месяцам
            For i = 2 To lastRow
                Dim docV  As String
                Dim osnV  As String
                Dim artV  As String
                Dim dtV   As Variant
                Dim mkV   As String
                docV = Trim(CStr(ws.Cells(i, colDoc).Value))
                osnV = Trim(CStr(ws.Cells(i, colOsn).Value))
                artV = Trim(CStr(ws.Cells(i, colArt).Value))
                If colDate > 0 Then
                    dtV = ws.Cells(i, colDate).Value
                Else
                    dtV = ws.Cells(i, 1).Value
                End If
                If IsDate(dtV) Then
                    mkV = Format(CDate(dtV), "YYYY-MM")
                Else
                    mkV = "0000-00"
                End If

                ' Инициализируем ключ если встречается впервые
                If Not dictVyk.Exists(mkV) Then
                    dictVyk.Add mkV, 0
                    dictVoz.Add mkV, 0
                    dictKmp.Add mkV, 0
                    dictDt.Add mkV, dtV
                    ' Null: IsDate(Null)=False — безопасная инициализация
                    dictMinDt.Add mkV, Null
                    dictMaxDt.Add mkV, Null
                End If
                ' Обновляем мин/макс ТОЛЬКО для Продажа/Возврат
                If (docV = "Продажа" Or docV = "Возврат") And IsDate(dtV) Then
                    If Not IsDate(dictMinDt(mkV)) Then
                        dictMinDt(mkV) = dtV
                    ElseIf CDate(dtV) < CDate(dictMinDt(mkV)) Then
                        dictMinDt(mkV) = dtV
                    End If
                    If Not IsDate(dictMaxDt(mkV)) Then
                        dictMaxDt(mkV) = dtV
                    ElseIf CDate(dtV) > CDate(dictMaxDt(mkV)) Then
                        dictMaxDt(mkV) = dtV
                    End If
                End If

                ' Шаг 2: col34 только для строк «Продажа»
                If docV = "Продажа" Then
                    If IsNumeric(ws.Cells(i, colPay).Value) Then
                        dictVyk(mkV) = dictVyk(mkV) + CDbl(ws.Cells(i, colPay).Value)
                    End If
                End If

                ' Шаг 3: col37 только для «Логистика» текущих артикулов
                If osnV = "Логистика" And dictArts.Exists(artV) Then
                    If IsNumeric(ws.Cells(i, colDel).Value) Then
                        dictVoz(mkV) = dictVoz(mkV) + CDbl(ws.Cells(i, colDel).Value)
                    End If
                End If
            Next i

        Else  ' Основной

            ' -------------------------------------------------------
            ' АЛГОРИТМ "Основной" (с разбивкой по месяцам)
            '   Выкупы  = SUM(col16 WHERE Тип = «Продажа»)
            '   Возвраты = SUM(col16 WHERE Тип = «Возврат»)
            '   Компенсации = SUM(col68) +/- + Добровольная компенсация
            ' -------------------------------------------------------
            If colSale = 0 Then
                MsgBox "Файл: " & FileName & Chr(13) & _
                       "Тип: Основной" & Chr(13) & _
                       "Не найден столбец 'Вайлдберриз реализовал'." & Chr(13) & _
                       "Файл пропущен.", vbExclamation
                wb.Close False
                FileName = Dir
                GoTo NextFile
            End If
            If colDoc = 0 Then
                MsgBox "Файл: " & FileName & Chr(13) & _
                       "Тип: Основной" & Chr(13) & _
                       "Не найден столбец 'Тип документа'." & Chr(13) & _
                       "Файл пропущен.", vbExclamation
                wb.Close False
                FileName = Dir
                GoTo NextFile
            End If

            ' Пре-скан: есть ли в файле строки Продажа/Возврат?
            ' Если нет — это отчёт без реальных продаж (например, старт продаж
            ' 510675152), и сервисные строки (Возмещение/Хранение/Удержания)
            ' являются единственным источником движения денег — учитываются в F.
            ' Если Продажа/Возврат есть, сервисные строки игнорируются
            ' (они уже отражены в "К перечислению Продавцу" по основным операциям).
            Dim hasSales As Boolean: hasSales = False
            Dim ps As Long
            For ps = 2 To lastRow
                Dim psDoc As String
                psDoc = Trim(CStr(ws.Cells(ps, colDoc).Value))
                If psDoc = "Продажа" Or psDoc = "Возврат" Then
                    hasSales = True
                    Exit For
                End If
            Next ps

            For i = 2 To lastRow
                Dim docType As String
                Dim osnType As String
                Dim saleVal As Double
                Dim dtT     As Variant
                Dim mkT     As String
                docType = Trim(CStr(ws.Cells(i, colDoc).Value))
                osnType = ""
                If colOsn > 0 Then osnType = Trim(CStr(ws.Cells(i, colOsn).Value))
                If colDate > 0 Then
                    dtT = ws.Cells(i, colDate).Value
                Else
                    dtT = ws.Cells(i, 1).Value
                End If
                If IsDate(dtT) Then
                    mkT = Format(CDate(dtT), "YYYY-MM")
                Else
                    mkT = "0000-00"
                End If

                ' Дневной ключ для определения точной даты превышения порога.
                Dim dkT As String
                If IsDate(dtT) Then
                    dkT = Format(CDate(dtT), "YYYY-MM-DD")
                Else
                    dkT = "0000-00-00"
                End If
                If Not dictDailyVyk.Exists(dkT) Then
                    dictDailyVyk.Add dkT, 0
                    dictDailyVoz.Add dkT, 0
                    dictDailyKmp.Add dkT, 0
                End If

                ' Инициализируем ключ если встречается впервые
                If Not dictVyk.Exists(mkT) Then
                    dictVyk.Add mkT, 0
                    dictVoz.Add mkT, 0
                    dictKmp.Add mkT, 0
                    dictDt.Add mkT, dtT
                    ' Null: IsDate(Null)=False — безопасная инициализация
                    dictMinDt.Add mkT, Null
                    dictMaxDt.Add mkT, Null
                End If
                ' Обновляем мин/макс для Продажа/Возврат.
                ' Для сервисных строк даты учитываются только если нет Продажа/Возврат
                ' в файле (hasSales=False) — иначе даты берутся с основных операций.
                Dim isMoneyRow As Boolean
                isMoneyRow = (docType = "Продажа" Or docType = "Возврат")
                If Not isMoneyRow And Not hasSales Then
                    If InStr(1, osnType, "Возмещение", vbTextCompare) > 0 _
                       Or InStr(1, osnType, "Хранение", vbTextCompare) > 0 _
                       Or InStr(1, osnType, "Удержания", vbTextCompare) > 0 Then
                        isMoneyRow = True
                    End If
                End If
                If isMoneyRow And IsDate(dtT) Then
                    If Not IsDate(dictMinDt(mkT)) Then
                        dictMinDt(mkT) = dtT
                    ElseIf CDate(dtT) < CDate(dictMinDt(mkT)) Then
                        dictMinDt(mkT) = dtT
                    End If
                    If Not IsDate(dictMaxDt(mkT)) Then
                        dictMaxDt(mkT) = dtT
                    ElseIf CDate(dtT) > CDate(dictMaxDt(mkT)) Then
                        dictMaxDt(mkT) = dtT
                    End If
                End If

                saleVal = 0
                If IsNumeric(ws.Cells(i, colSale).Value) Then
                    saleVal = CDbl(ws.Cells(i, colSale).Value)
                End If
                If docType = "Продажа" Then dictVyk(mkT) = dictVyk(mkT) + saleVal
                If docType = "Продажа" Then dictDailyVyk(dkT) = dictDailyVyk(dkT) + saleVal
                If docType = "Возврат" Then dictVoz(mkT) = dictVoz(mkT) + saleVal
                If docType = "Возврат" Then dictDailyVoz(dkT) = dictDailyVoz(dkT) + saleVal

                If colComp > 0 Then
                    If IsNumeric(ws.Cells(i, colComp).Value) Then
                        Dim compVal As Double
                        compVal = CDbl(ws.Cells(i, colComp).Value)
                        If docType = "Возврат" Then
                            dictKmp(mkT) = dictKmp(mkT) - compVal
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) - compVal
                        Else
                            dictKmp(mkT) = dictKmp(mkT) + compVal
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) + compVal
                        End If
                    End If
                End If

                If colPay > 0 Then
                    If InStr(1, osnType, "Добровольная компенсация", vbTextCompare) > 0 Then
                        If IsNumeric(ws.Cells(i, colPay).Value) Then
                            dictKmp(mkT) = dictKmp(mkT) + CDbl(ws.Cells(i, colPay).Value)
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) + CDbl(ws.Cells(i, colPay).Value)
                        End If
                    End If
                End If

                ' Сервисные строки учитываются только в файлах без Продажа/Возврат
                ' (гейт hasSales). В обычных еженедельных отчётах Возмещение/Хранение/
                ' Удержания уже отражены в поле "К перечислению Продавцу" и их
                ' добавление привело бы к двойному учёту.
                If Not hasSales Then
                    '   Возмещение издержек -> dictKmp += col"Возмещение"
                    '   Хранение            -> dictKmp -= col"Хранение"
                    '   Удержания           -> dictKmp -= col"Удержания"
                    If colVozm > 0 And InStr(1, osnType, "Возмещение издержек", vbTextCompare) > 0 Then
                        If IsNumeric(ws.Cells(i, colVozm).Value) Then
                            Dim vVozm As Double: vVozm = CDbl(ws.Cells(i, colVozm).Value)
                            dictKmp(mkT) = dictKmp(mkT) + vVozm
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) + vVozm
                        End If
                    End If
                    If colStore > 0 And InStr(1, osnType, "Хранение", vbTextCompare) > 0 Then
                        If IsNumeric(ws.Cells(i, colStore).Value) Then
                            Dim vStore As Double: vStore = CDbl(ws.Cells(i, colStore).Value)
                            dictKmp(mkT) = dictKmp(mkT) - vStore
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) - vStore
                        End If
                    End If
                    If colDeduct > 0 And InStr(1, osnType, "Удержания", vbTextCompare) > 0 Then
                        If IsNumeric(ws.Cells(i, colDeduct).Value) Then
                            Dim vDeduct As Double: vDeduct = CDbl(ws.Cells(i, colDeduct).Value)
                            dictKmp(mkT) = dictKmp(mkT) - vDeduct
                            dictDailyKmp(dkT) = dictDailyKmp(dkT) - vDeduct
                        End If
                    End If
                End If
            Next i

        End If  ' reportType

        ' ========================
        ' Запись в Данные - ОДНА строка на файл.
        ' Разбивка на 2 строки допустима только для файла, в котором
        ' происходит превышение порога 20 млн. Это определяется позже
        ' в блоке расчёта НДС. Здесь сохраняем помесячную разбивку
        ' в col K (служебный) для возможного разделения.
        ' ========================
        Dim mk     As Variant
        Dim repNum As String
        repNum = ExtractReportNumber(FileName)

        ' Суммируем все месяцы в одну строку, сериализуем разбивку в col M
        Dim totalVyk As Double: totalVyk = 0
        Dim totalVoz As Double: totalVoz = 0
        Dim totalKmp As Double: totalKmp = 0
        Dim firstDt   As Variant: firstDt = ""
        Dim fileMinDt As Variant: fileMinDt = ""
        Dim fileMaxDt As Variant: fileMaxDt = ""
        Dim mSerial   As String: mSerial = ""

        For Each mk In dictVyk.Keys
            totalVyk = totalVyk + dictVyk(mk)
            totalVoz = totalVoz + dictVoz(mk)
            totalKmp = totalKmp + dictKmp(mk)
            ' Запоминаем самую раннюю дату для col A (Дата формирования)
            If firstDt = "" Then
                firstDt = dictDt(mk)
            ElseIf IsDate(dictDt(mk)) And IsDate(firstDt) Then
                If CDate(dictDt(mk)) < CDate(firstDt) Then firstDt = dictDt(mk)
            End If
            ' Обновляем общий мин/макс — только если словарь содержит реальную дату
            If IsDate(dictMinDt(mk)) Then
                If Not IsDate(fileMinDt) Then
                    fileMinDt = dictMinDt(mk)
                ElseIf CDate(dictMinDt(mk)) < CDate(fileMinDt) Then
                    fileMinDt = dictMinDt(mk)
                End If
            End If
            If IsDate(dictMaxDt(mk)) Then
                If Not IsDate(fileMaxDt) Then
                    fileMaxDt = dictMaxDt(mk)
                ElseIf CDate(dictMaxDt(mk)) > CDate(fileMaxDt) Then
                    fileMaxDt = dictMaxDt(mk)
                End If
            End If
            ' Сериализуем: YYYY-MM~vyk~voz~kmp~dtNum~minDtNum~maxDtNum
            Dim dtNum    As Long: dtNum = 0
            Dim minDtNum As Long: minDtNum = 0
            Dim maxDtNum As Long: maxDtNum = 0
            If IsDate(dictDt(mk)) Then dtNum = CLng(CDbl(CDate(dictDt(mk))))
            If IsDate(dictMinDt(mk)) Then minDtNum = CLng(CDbl(CDate(dictMinDt(mk))))
            If IsDate(dictMaxDt(mk)) Then maxDtNum = CLng(CDbl(CDate(dictMaxDt(mk))))
            mSerial = mSerial & CStr(mk) & "~" & _
                      Str(dictVyk(mk)) & "~" & _
                      Str(dictVoz(mk)) & "~" & _
                      Str(dictKmp(mk)) & "~" & _
                      CStr(dtNum) & "~" & _
                      CStr(minDtNum) & "~" & _
                      CStr(maxDtNum) & "|"
        Next mk

        Dim totalBaz As Double
        totalBaz = totalVyk - totalVoz + totalKmp

        Dim newRow As Long
        newRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1

        ' Дата формирования отчета = макс. Дата продажи + 1 день
        Dim fileFormDate As Variant
        If IsDate(fileMaxDt) Then
            fileFormDate = CDate(fileMaxDt) + 1
        Else
            fileFormDate = fileMinDt
        End If

        wsData.Cells(newRow, 1).Value = fileFormDate  ' A: Дата формирования отчета
        wsData.Cells(newRow, 2).Value = fileMinDt    ' B: Дата начала
        wsData.Cells(newRow, 3).Value = fileMaxDt    ' C: Дата конца
        wsData.Cells(newRow, 4).Value = totalVyk     ' D: Выкупы
        wsData.Cells(newRow, 5).Value = totalVoz     ' E: Возвраты
        wsData.Cells(newRow, 6).Value = totalKmp     ' F: Компенсации
        wsData.Cells(newRow, 7).Value = totalBaz     ' G: Доход базовый
        wsData.Cells(newRow, 10).Value = reportType  ' J: Тип отчёта
        wsData.Cells(newRow, 11).Value = repNum      ' K: Номер отчёта
        wsData.Cells(newRow, 12).Value = mSerial     ' L: Служебный

        ' Дневная сериализация: "YYYY-MM-DD~base|..." — для точной даты порога.
        Dim dSerial As String: dSerial = ""
        Dim ddk As Variant
        For Each ddk In dictDailyVyk.Keys
            Dim dBase As Double
            dBase = dictDailyVyk(ddk) - dictDailyVoz(ddk) + dictDailyKmp(ddk)
            If dBase <> 0 Then
                dSerial = dSerial & CStr(ddk) & "~" & Str(dBase) & "|"
            End If
        Next ddk
        wsData.Cells(newRow, 13).Value = dSerial     ' M: разбивка по дням

        cntTotal = cntTotal + 1
        If reportType = "Основной" Then
            cntOsn = cntOsn + 1
        Else
            cntVykup = cntVykup + 1
        End If

        wb.Close False

NextFile:
        FileName = Dir

    Loop

    ' ========================
    ' Сортировка по дате
    ' ========================
    Dim lastDataRow As Long
    lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    If lastDataRow > 1 Then
        wsData.Sort.SortFields.Clear
        wsData.Sort.SortFields.Add Key:=wsData.Range("B2:B" & lastDataRow), _
                                    Order:=xlAscending
        With wsData.Sort
            .SetRange wsData.Range("A1:M" & lastDataRow)
            .Header = xlYes
            .Apply
        End With
    End If

    ' ========================
    ' Расчёт НДС и Дохода без НДС
    '
    ' Логика:
    '   - При ставке НДС = 0% > НДС не начисляется
    '   - При ставке НДС > 0% (напр., 5%):
    '       · Суммируем доход нарастающим итогом
    '       · Как только сумма >= LIMIT — фиксируем дату превышения
    '       · НДС начисляется с 1-го числа СЛЕДУЮЩЕГО месяца
    '       · Сумма НДС     = Доход базовый * НДС% / (100 + НДС%)
    '         (НДС выделяется из суммы "в т.ч. НДС", не начисляется сверху)
    '       · Доход без НДС = Доход базовый - Сумма НДС
    '         (это и есть налогооблагаемая база по УСН "Доходы")
    ' ========================
    Dim cumul       As Double
    Dim threshDate  As Date
    Dim ndsStart    As Date
    Dim threshFound As Boolean
    Dim ndsActive   As Boolean

    cumul = 0
    threshFound = False
    ndsActive = False

    ' Do While вместо For — нужен для вставки строки при разделении
    ' порогового файла (lastDataRow может увеличиться на 1).
    Dim r As Long
    r = 2
    Do While r <= lastDataRow

        Dim rDate   As Variant
        Dim rBase   As Double
        Dim rNDS    As Double
        Dim rBezNDS As Double

        rDate = wsData.Cells(r, 2).Value  ' Дата начала — для определения периода порога
        rBase = wsData.Cells(r, 7).Value

        ' Шаг 1: если порог ещё не пройден — проверяем, не пересечёт ли его эта строка.
        ' Только фиксируем ndsStart; сплит выполняется ниже единым блоком.
        If Not threshFound And (cumul + rBase >= LIMIT) Then
            threshFound = True
            ' Точная дата превышения: по дневной разбивке (col M).
            Dim dSerialRow As String
            dSerialRow = CStr(wsData.Cells(r, 13).Value)
            threshDate = FindDayOfExceed(dSerialRow, cumul, LIMIT)
            ' Fallback: если dailySerial пустой/нечитаемый — старое поведение.
            If threshDate = 0 Then
                If IsDate(rDate) Then threshDate = CDate(rDate)
            End If
            If threshDate > 0 Then
                ndsStart = DateSerial(Year(threshDate), Month(threshDate) + 1, 1)
            End If
        End If

        ' Шаг 2: если ndsStart известна — пробуем разбить ЛЮБУЮ строку,
        ' у которой в col K есть помесячная разбивка по обе стороны от ndsStart.
        ' Это покрывает как пороговый файл, так и соседние файлы того же периода.
        If threshFound Then
            Dim mData As String
            mData = CStr(wsData.Cells(r, 12).Value)
            If mData <> "" And InStr(mData, "|") > 0 Then
                Dim preVyk   As Double: preVyk = 0
                Dim preVoz   As Double: preVoz = 0
                Dim preKmp   As Double: preKmp = 0
                Dim preDt    As Variant: preDt = rDate
                Dim preMinDt As Variant: preMinDt = ""
                Dim preMaxDt As Variant: preMaxDt = ""
                Dim postVyk   As Double: postVyk = 0
                Dim postVoz   As Double: postVoz = 0
                Dim postKmp   As Double: postKmp = 0
                Dim postDt    As Variant: postDt = ""
                Dim postMinDt As Variant: postMinDt = ""
                Dim postMaxDt As Variant: postMaxDt = ""

                Dim mParts() As String
                mParts = Split(mData, "|")
                Dim pi As Long
                For pi = 0 To UBound(mParts)
                    If mParts(pi) = "" Then GoTo SkipPart
                    Dim flds() As String
                    flds = Split(mParts(pi), "~")
                    If UBound(flds) < 4 Then GoTo SkipPart
                    Dim pMonthStart As Date
                    pMonthStart = DateSerial(CInt(Left(flds(0), 4)), CInt(Mid(flds(0), 6, 2)), 1)
                    Dim pDt    As Variant
                    Dim pMinDt As Variant
                    Dim pMaxDt As Variant
                    If CLng(flds(4)) > 0 Then pDt = CDate(CLng(flds(4))) Else pDt = pMonthStart
                    If UBound(flds) >= 5 And CLng(flds(5)) > 0 Then
                        pMinDt = CDate(CLng(flds(5)))
                    Else
                        pMinDt = pDt
                    End If
                    If UBound(flds) >= 6 And CLng(flds(6)) > 0 Then
                        pMaxDt = CDate(CLng(flds(6)))
                    Else
                        pMaxDt = pDt
                    End If
                    Dim pVyk As Double: pVyk = val(flds(1))
                    Dim pVoz As Double: pVoz = val(flds(2))
                    Dim pKmp As Double: pKmp = val(flds(3))
                    If pMonthStart >= ndsStart Then
                        postVyk = postVyk + pVyk
                        postVoz = postVoz + pVoz
                        postKmp = postKmp + pKmp
                        If Not IsDate(postDt) Then postDt = pDt
                        If Not IsDate(postMinDt) Then
                            postMinDt = pMinDt: postMaxDt = pMaxDt
                        Else
                            If IsDate(pMinDt) And CDate(pMinDt) < CDate(postMinDt) Then postMinDt = pMinDt
                            If IsDate(pMaxDt) And CDate(pMaxDt) > CDate(postMaxDt) Then postMaxDt = pMaxDt
                        End If
                    Else
                        preVyk = preVyk + pVyk
                        preVoz = preVoz + pVoz
                        preKmp = preKmp + pKmp
                        If Not IsDate(preMinDt) Then
                            preMinDt = pMinDt: preMaxDt = pMaxDt
                        Else
                            If IsDate(pMinDt) And CDate(pMinDt) < CDate(preMinDt) Then preMinDt = pMinDt
                            If IsDate(pMaxDt) And CDate(pMaxDt) > CDate(preMaxDt) Then preMaxDt = pMaxDt
                        End If
                    End If
SkipPart:
                Next pi

                Dim preBaz  As Double: preBaz = preVyk - preVoz + preKmp
                Dim postBaz As Double: postBaz = postVyk - postVoz + postKmp

                ' Разбиваем только если обе части ненулевые
                If preBaz <> 0 And postBaz <> 0 Then
                    ' Обновляем строку r: доход ДО ndsStart
                    ' Дата формирования = макс. дата этой части + 1 день
                    If IsDate(preMaxDt) Then wsData.Cells(r, 1).Value = CDate(preMaxDt) + 1
                    wsData.Cells(r, 2).Value = preMinDt   ' B: Дата начала
                    wsData.Cells(r, 3).Value = preMaxDt   ' C: Дата конца
                    wsData.Cells(r, 4).Value = preVyk     ' D: Выкупы
                    wsData.Cells(r, 5).Value = preVoz     ' E: Возвраты
                    wsData.Cells(r, 6).Value = preKmp     ' F: Компенсации
                    wsData.Cells(r, 7).Value = preBaz     ' G: Доход базовый
                    wsData.Cells(r, 12).Value = ""         ' L: очищаем mSerial — иначе Свод посчитает все месяцы дважды
                    rBase = preBaz

                    ' Вставляем строку r+1: доход ОТ ndsStart
                    wsData.Rows(r + 1).Insert Shift:=xlDown
                    Dim postFormDate As Variant
                    If IsDate(postMaxDt) Then postFormDate = CDate(postMaxDt) + 1 Else postFormDate = postDt
                    wsData.Cells(r + 1, 1).Value = postFormDate ' A: Дата формирования
                    wsData.Cells(r + 1, 2).Value = postMinDt    ' B: Дата начала
                    wsData.Cells(r + 1, 3).Value = postMaxDt    ' C: Дата конца
                    wsData.Cells(r + 1, 4).Value = postVyk      ' D: Выкупы
                    wsData.Cells(r + 1, 5).Value = postVoz      ' E: Возвраты
                    wsData.Cells(r + 1, 6).Value = postKmp      ' F: Компенсации
                    wsData.Cells(r + 1, 7).Value = postBaz      ' G: Доход базовый
                    wsData.Cells(r + 1, 10).Value = wsData.Cells(r, 10).Value ' J: Тип
                    wsData.Cells(r + 1, 11).Value = wsData.Cells(r, 11).Value ' K: Номер
                    lastDataRow = lastDataRow + 1
                End If
            End If
        End If

        ' Накопительный итог
        cumul = cumul + rBase

        ' Определяем ставку по АКТУАЛЬНОЙ дате строки (перечитываем после сплита).
        ' Прямое сравнение с ndsStart — без накопленного флага ndsActive,
        ' чтобы строка Feb-части сплитованного файла не получила мартовскую ставку.
        Dim curRate     As Double
        Dim effectDate  As Variant: effectDate = wsData.Cells(r, 3).Value   ' Дата конца — для применения ставки НДС
        If threshFound And IsDate(effectDate) And CDate(effectDate) >= ndsStart Then
            curRate = ndsRate2
            ndsActive = True  ' для совместимости с блоком Контроль
        Else
            curRate = ndsRate1
        End If

        If curRate > 0 Then
            rNDS = rBase * curRate / (100 + curRate)
            rBezNDS = rBase - rNDS
        Else
            rNDS = 0
            rBezNDS = rBase
        End If

        wsData.Cells(r, 8).Value = rNDS       ' H: Сумма НДС
        wsData.Cells(r, 9).Value = rBezNDS    ' I: Доход без НДС

        r = r + 1
    Loop

    ' Очистка технической колонки M (разбивка по дням): использовалась только для
    ' определения точной даты превышения порога в FindDayOfExceed.
    If lastDataRow >= 1 Then
        wsData.Range("M1:M" & lastDataRow).ClearContents
    End If

    ' ========================
    ' Итоги в лист "Контроль"
    ' ========================
    Dim totBase   As Double
    Dim totNDS    As Double
    Dim totBezNDS As Double

    If lastDataRow > 1 Then
        totBase = Application.WorksheetFunction.Sum(wsData.Range("G2:G" & lastDataRow))
        totNDS = Application.WorksheetFunction.Sum(wsData.Range("H2:H" & lastDataRow))
        totBezNDS = Application.WorksheetFunction.Sum(wsData.Range("I2:I" & lastDataRow))
    End If

    wsCtrl.Cells(2, 2).Value = totBase
    wsCtrl.Cells(3, 2).Value = totNDS
    wsCtrl.Cells(4, 2).Value = totBezNDS

    If prevYearExceeded Then
        ' Прошлый год превысил 60 млн — НДС действует с 01.01 текущего года
        wsCtrl.Cells(5, 2).Value = "ДА (доход прошл. года " & Format(prevYearIncome, "# ##0") & " руб. >= 60 млн)"
        wsCtrl.Cells(6, 2).Value = "прошлый год > 60 млн"
        wsCtrl.Cells(7, 2).Value = DateSerial(Year(Now()), 1, 1)   ' 01.01 текущего года
    ElseIf threshFound Then
        ' Порог 20 млн превышен в текущем году
        wsCtrl.Cells(5, 2).Value = "ДА"
        wsCtrl.Cells(6, 2).Value = threshDate
        wsCtrl.Cells(7, 2).Value = ndsStart
    Else
        wsCtrl.Cells(5, 2).Value = "НЕТ (накоплено " & Format(totBase, "# ##0") & " из 20 000 000 руб.)"
        wsCtrl.Cells(6, 2).Value = "-"
        wsCtrl.Cells(7, 2).Value = "-"
    End If

    wsCtrl.Cells(8, 2).Value = cntTotal
    wsCtrl.Cells(9, 2).Value = cntOsn
    wsCtrl.Cells(10, 2).Value = cntVykup

    ' ========================
    ' Заполнение листа "Свод" по месяцам
    ' ========================
    Dim monthNames(1 To 12) As String
    monthNames(1) = "Январь":    monthNames(2) = "Февраль"
    monthNames(3) = "Март":      monthNames(4) = "Апрель"
    monthNames(5) = "Май":       monthNames(6) = "Июнь"
    monthNames(7) = "Июль":      monthNames(8) = "Август"
    monthNames(9) = "Сентябрь": monthNames(10) = "Октябрь"
    monthNames(11) = "Ноябрь":   monthNames(12) = "Декабрь"

    Dim mVyk(1 To 12)    As Double
    Dim mVoz(1 To 12)    As Double
    Dim mKmp(1 To 12)    As Double
    Dim mBaz(1 To 12)    As Double
    Dim mNDS(1 To 12)    As Double
    Dim mBez(1 To 12)    As Double
    Dim mHasNDS(1 To 12) As Boolean
    Dim mCntOsn(1 To 12) As Long
    Dim mCntVyk(1 To 12) As Long

    Dim m As Integer
    For m = 1 To 12
        mVyk(m) = 0: mVoz(m) = 0: mKmp(m) = 0
        mBaz(m) = 0: mNDS(m) = 0: mBez(m) = 0
        mHasNDS(m) = False
        mCntOsn(m) = 0: mCntVyk(m) = 0
    Next m

    ' Строим Свод напрямую из mSerial (col L) — помесячная разбивка по Дате продажи.
    ' Для каждой строки Данных читаем mSerial, парсим месячные бакеты,
    ' распределяем НДС пропорционально доходу каждого месяца.
    For r = 2 To lastDataRow
        Dim rTp As String
        rTp = Trim(CStr(wsData.Cells(r, 10).Value))
        Dim rSerial  As String: rSerial = CStr(wsData.Cells(r, 12).Value)
        Dim rTotBase As Double: rTotBase = wsData.Cells(r, 7).Value
        Dim rTotNDS  As Double: rTotNDS = wsData.Cells(r, 8).Value

        If rSerial <> "" And InStr(rSerial, "|") > 0 Then
            ' Строка содержит mSerial — разбираем по месяцам
            Dim sParts() As String
            sParts = Split(rSerial, "|")
            ' Шаг 1: максимальный год среди бакетов
            Dim maxYear As Integer: maxYear = 0
            Dim spi As Long
            For spi = 0 To UBound(sParts)
                If sParts(spi) <> "" Then
                    Dim spf() As String: spf = Split(sParts(spi), "~")
                    If UBound(spf) >= 0 And Len(spf(0)) >= 4 Then
                        Dim spY As Integer: spY = CInt(Left(spf(0), 4))
                        If spY > maxYear Then maxYear = spY
                    End If
                End If
            Next spi
            ' Шаг 2: парсим, год < maxYear -> январь
            Dim si As Long
            For si = 0 To UBound(sParts)
                If sParts(si) = "" Then GoTo SkipSerial
                Dim sf() As String
                sf = Split(sParts(si), "~")
                If UBound(sf) < 3 Then GoTo SkipSerial
                Dim sYear  As Integer: sYear = CInt(Left(sf(0), 4))
                Dim sMonth As Integer
                If sYear < maxYear Then
                    sMonth = 1
                Else
                    sMonth = CInt(Mid(sf(0), 6, 2))
                End If
                If sMonth < 1 Or sMonth > 12 Then GoTo SkipSerial
                Dim sVyk As Double: sVyk = val(sf(1))
                Dim sVoz As Double: sVoz = val(sf(2))
                Dim sKmp As Double: sKmp = val(sf(3))
                Dim sBaz As Double: sBaz = sVyk - sVoz + sKmp
                ' НДС пропорционально доле месяца в общем доходе строки
                Dim sNDS As Double: sNDS = 0
                If rTotBase <> 0 Then sNDS = rTotNDS * sBaz / rTotBase
                Dim sBez As Double: sBez = sBaz - sNDS
                mVyk(sMonth) = mVyk(sMonth) + sVyk
                mVoz(sMonth) = mVoz(sMonth) + sVoz
                mKmp(sMonth) = mKmp(sMonth) + sKmp
                mBaz(sMonth) = mBaz(sMonth) + sBaz
                mNDS(sMonth) = mNDS(sMonth) + sNDS
                mBez(sMonth) = mBez(sMonth) + sBez
                If sNDS <> 0 Then mHasNDS(sMonth) = True
                If rTp = "Основной" Then mCntOsn(sMonth) = mCntOsn(sMonth) + 1
                If rTp = "По выкупам" Then mCntVyk(sMonth) = mCntVyk(sMonth) + 1
SkipSerial:
            Next si
        Else
            ' Нет mSerial (сплит-строка или один месяц) — используем Дату конца
            Dim rEnd2 As Variant: rEnd2 = wsData.Cells(r, 3).Value
            Dim rD2   As Variant: rD2 = wsData.Cells(r, 2).Value
            Dim rMon  As Integer: rMon = 0
            If IsDate(rEnd2) Then
                rMon = Month(CDate(rEnd2))
            ElseIf IsDate(rD2) Then
                rMon = Month(CDate(rD2))
            End If
            If rMon >= 1 And rMon <= 12 Then
                mVyk(rMon) = mVyk(rMon) + wsData.Cells(r, 4).Value
                mVoz(rMon) = mVoz(rMon) + wsData.Cells(r, 5).Value
                mKmp(rMon) = mKmp(rMon) + wsData.Cells(r, 6).Value
                mBaz(rMon) = mBaz(rMon) + rTotBase
                mNDS(rMon) = mNDS(rMon) + rTotNDS
                mBez(rMon) = mBez(rMon) + wsData.Cells(r, 9).Value
                If rTotNDS <> 0 Then mHasNDS(rMon) = True
                If rTp = "Основной" Then mCntOsn(rMon) = mCntOsn(rMon) + 1
                If rTp = "По выкупам" Then mCntVyk(rMon) = mCntVyk(rMon) + 1
            End If
        End If
    Next r

    ' Очистка старых данных Свода
    Dim lastSvod As Long
    lastSvod = wsSvod.Cells(wsSvod.Rows.Count, 1).End(xlUp).Row
    If lastSvod > 1 Then wsSvod.Rows("2:" & lastSvod).Delete

    ' Маппинг месяца в строку Свода с учётом квартальных итогов:
    ' Q1: Янв=2 Фев=3 Мар=4 -> Q1=5
    ' Q2: Апр=6 Май=7 Июн=8 -> Q2=9
    ' Q3: Июл=10 Авг=11 Сен=12 -> Q3=13
    ' Q4: Окт=14 Ноя=15 Дек=16 -> Q4=17  Год=18
    Dim mRowMap(1 To 12) As Long
    mRowMap(1) = 2: mRowMap(2) = 3: mRowMap(3) = 4
    mRowMap(4) = 6: mRowMap(5) = 7: mRowMap(6) = 8
    mRowMap(7) = 10: mRowMap(8) = 11: mRowMap(9) = 12
    mRowMap(10) = 14: mRowMap(11) = 15: mRowMap(12) = 16

    Dim sRow As Long
    For m = 1 To 12
        sRow = mRowMap(m)

        wsSvod.Cells(sRow, 1).Value = monthNames(m)
        wsSvod.Cells(sRow, 2).Value = mVyk(m)
        wsSvod.Cells(sRow, 3).Value = mVoz(m)
        wsSvod.Cells(sRow, 4).Value = mKmp(m)
        wsSvod.Cells(sRow, 5).Value = mBaz(m)
        wsSvod.Cells(sRow, 6).Value = mNDS(m)
        wsSvod.Cells(sRow, 7).Value = mBez(m)

        If mBaz(m) = 0 Then
            wsSvod.Cells(sRow, 8).Value = ""
        ElseIf mHasNDS(m) Then
            wsSvod.Cells(sRow, 8).Value = Format(ndsRate2, "0") & "%"
        Else
            wsSvod.Cells(sRow, 8).Value = Format(ndsRate1, "0") & "%"
        End If

        wsSvod.Cells(sRow, 9).Value = IIf(mCntOsn(m) > 0, mCntOsn(m), "")
        wsSvod.Cells(sRow, 10).Value = IIf(mCntVyk(m) > 0, mCntVyk(m), "")
    Next m

    ' ========================
    ' Квартальные итоги и итог за год
    ' ========================
    Dim qNames(1 To 4) As String
    qNames(1) = "Итого за 1 квартал"
    qNames(2) = "Итого за 2 квартал"
    qNames(3) = "Итого за 3 квартал"
    qNames(4) = "Итого за 4 квартал"
    Dim qRows(1 To 4) As Long
    qRows(1) = 5: qRows(2) = 9: qRows(3) = 13: qRows(4) = 17
    ' Месяцы каждого квартала
    Dim qStart(1 To 4) As Integer: Dim qEnd(1 To 4) As Integer
    qStart(1) = 1: qEnd(1) = 3: qStart(2) = 4: qEnd(2) = 6
    qStart(3) = 7: qEnd(3) = 9: qStart(4) = 10: qEnd(4) = 12

    Dim q As Integer
    For q = 1 To 4
        Dim qVyk As Double: qVyk = 0
        Dim qVoz As Double: qVoz = 0
        Dim qKmp As Double: qKmp = 0
        Dim qBaz As Double: qBaz = 0
        Dim qNDS As Double: qNDS = 0
        Dim qBez As Double: qBez = 0
        Dim qm   As Integer
        For qm = qStart(q) To qEnd(q)
            qVyk = qVyk + mVyk(qm)
            qVoz = qVoz + mVoz(qm)
            qKmp = qKmp + mKmp(qm)
            qBaz = qBaz + mBaz(qm)
            qNDS = qNDS + mNDS(qm)
            qBez = qBez + mBez(qm)
        Next qm
        Dim qRow As Long: qRow = qRows(q)
        wsSvod.Cells(qRow, 1).Value = qNames(q)
        wsSvod.Cells(qRow, 2).Value = qVyk
        wsSvod.Cells(qRow, 3).Value = qVoz
        wsSvod.Cells(qRow, 4).Value = qKmp
        wsSvod.Cells(qRow, 5).Value = qBaz
        wsSvod.Cells(qRow, 6).Value = qNDS
        wsSvod.Cells(qRow, 7).Value = qBez
        ' Очищаем вспомогательные столбцы для итоговых строк
        wsSvod.Cells(qRow, 8).Value = ""
        wsSvod.Cells(qRow, 9).Value = ""
        wsSvod.Cells(qRow, 10).Value = ""
    Next q

    ' Итого за год (строка 18) = сумма квартальных итогов
    Dim yVyk As Double: yVyk = 0
    Dim yVoz As Double: yVoz = 0
    Dim yKmp As Double: yKmp = 0
    Dim yBaz As Double: yBaz = 0
    Dim yNDS As Double: yNDS = 0
    Dim yBez As Double: yBez = 0
    For q = 1 To 4
        Dim qqm As Integer
        For qqm = qStart(q) To qEnd(q)
            yVyk = yVyk + mVyk(qqm)
            yVoz = yVoz + mVoz(qqm)
            yKmp = yKmp + mKmp(qqm)
            yBaz = yBaz + mBaz(qqm)
            yNDS = yNDS + mNDS(qqm)
            yBez = yBez + mBez(qqm)
        Next qqm
    Next q
    wsSvod.Cells(18, 1).Value = "Итого за год"
    wsSvod.Cells(18, 2).Value = yVyk
    wsSvod.Cells(18, 3).Value = yVoz
    wsSvod.Cells(18, 4).Value = yKmp
    wsSvod.Cells(18, 5).Value = yBaz
    wsSvod.Cells(18, 6).Value = yNDS
    wsSvod.Cells(18, 7).Value = yBez
    wsSvod.Cells(18, 8).Value = ""
    wsSvod.Cells(18, 9).Value = ""
    wsSvod.Cells(18, 10).Value = ""

    ' Очищаем служебный столбец L на листе Данные
    If lastDataRow > 1 Then
        wsData.Range("L2:L" & lastDataRow).ClearContents
    End If

    Application.ScreenUpdating = True

    ' ========================
    ' Итоговое сообщение
    ' ========================
    Dim totVyk As Double
    Dim totVoz As Double
    Dim totKmp As Double
    If lastDataRow > 1 Then
        totVyk = Application.WorksheetFunction.Sum(wsData.Range("D2:D" & lastDataRow))
        totVoz = Application.WorksheetFunction.Sum(wsData.Range("E2:E" & lastDataRow))
        totKmp = Application.WorksheetFunction.Sum(wsData.Range("F2:F" & lastDataRow))
    End If

    Dim msgText As String
    msgText = "WB_nalog_USN_NDS — расчёт завершён!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
              "Файлов обработано:  " & cntTotal & _
              " (Основной: " & cntOsn & ", По выкупам: " & cntVykup & ")" & Chr(13) & Chr(10) & _
              "Доход пр. год:      " & Format(prevYearIncome, "# ##0") & " руб." & Chr(13) & Chr(10) & _
              IIf(prevYearExceeded, "  => лимит 60 млн превышен, НДС с 01.01 текущего года" & Chr(13) & Chr(10), "") & _
              "НДС до 20 млн:      " & Format(ndsRate1, "0") & "%" & Chr(13) & Chr(10) & _
              "НДС после 20 млн:   " & Format(ndsRate2, "0") & "%" & Chr(13) & Chr(10) & _
              "------------------------------------------" & Chr(13) & Chr(10) & _
              "Выкупы:             " & Format(totVyk, "# ##0.00") & " руб." & Chr(13) & Chr(10) & _
              "Возвраты:           " & Format(totVoz, "# ##0.00") & " руб." & Chr(13) & Chr(10) & _
              "Компенсации:        " & Format(totKmp, "# ##0.00") & " руб." & Chr(13) & Chr(10) & _
              "------------------------------------------" & Chr(13) & Chr(10) & _
              "Доход базовый:      " & Format(totBase, "# ##0.00") & " руб." & Chr(13) & Chr(10) & _
              "НДС:                " & Format(totNDS, "# ##0.00") & " руб." & Chr(13) & Chr(10) & _
              "Доход без НДС:      " & Format(totBezNDS, "# ##0.00") & " руб."

    


    ' 5. Заполнение сумм для 1 квартала 2026 из листа "Свод".
    '    НДС: полный за квартал = Свод!F5; оплаты 1/3, 2/3, 3/3 = Свод!F5 / 3
    '    УСН: аванс 1 кв = Настройки!B8 (ставка УСН) * Свод!G5 (доход без НДС)
    Dim ndsQ1  As Double
    Dim bezQ1  As Double
    Dim rateUSN As Double
    Dim vF5 As Variant, vG5 As Variant, vB8 As Variant

    vF5 = wsSvod.Cells(5, 6).Value    ' Итог 1 кв, колонка "в т.ч. НДС"
    vG5 = wsSvod.Cells(5, 7).Value    ' Итог 1 кв, колонка "Доход без НДС"
    vB8 = wsSet.Cells(8, 2).Value     ' Настройки B8 - ставка УСН

    If IsNumeric(vF5) Then ndsQ1 = CDbl(vF5) Else ndsQ1 = 0
    If IsNumeric(vG5) Then bezQ1 = CDbl(vG5) Else bezQ1 = 0
    If IsNumeric(vB8) Then rateUSN = CDbl(vB8) Else rateUSN = 0

    wsCal.Cells(9, 3).Value = ndsQ1            ' Подача декларации НДС за 1 кв 2026
    wsCal.Cells(10, 3).Value = rateUSN * bezQ1 ' Оплата аванса УСН 1 квартал
    wsCal.Cells(11, 3).Value = ndsQ1 / 3       ' Оплата НДС 1/3 за 1 кв 2026
    wsCal.Cells(14, 3).Value = ndsQ1 / 3       ' Оплата НДС 2/3 за 1 кв 2026
    wsCal.Cells(15, 3).Value = ndsQ1 / 3       ' Оплата НДС 3/3 за 1 кв 2026

    ' 6. Расчёты по прошлому году:
    '    B10 (авто) = доп. страховой взнос 1% = MIN(B9, MAX(0, (B5-300000)*1%))
    '    B11 (ручной ввод) = сумма фиксированных страховых взносов за 2025
    '    C2  = B11
    '    C12 = B8*B5 (декларация, без вычетов)
    '    C13 = MAX(0, B8*B5 - B10 - B11)   (к доплате минус все взносы)
    '    C16 = B10                          (зеркало доп. страх. взноса)
    Dim vB9 As Variant, vB11 As Variant
    Dim cap1Pct           As Double
    Dim addInsurance2025  As Double
    Dim fixedInsurance25  As Double
    vB9 = wsSet.Cells(9, 2).Value
    vB11 = wsSet.Cells(11, 2).Value
    ' Кап 1% взноса: берём из B9; если пусто/0 — автозаполняем дефолтом 2025 (277 571 руб. = 8*фикс.взнос)
    If IsNumeric(vB9) And CDbl(vB9) > 0 Then
        cap1Pct = CDbl(vB9)
    Else
        cap1Pct = 277571
        wsSet.Cells(9, 2).Value = cap1Pct
    End If
    If IsNumeric(vB11) Then fixedInsurance25 = CDbl(vB11) Else fixedInsurance25 = 0

    ' Автозаполнение B10: доп. страховой взнос 1% с превышения 300 000, ограниченный капом из B9
    addInsurance2025 = (prevYearIncome - 300000) * 0.01
    If addInsurance2025 < 0 Then addInsurance2025 = 0
    If addInsurance2025 > cap1Pct Then addInsurance2025 = cap1Pct
    wsSet.Cells(10, 2).Value = addInsurance2025

    ' Календарь: суммы из полей Настроек
    wsCal.Cells(2, 3).Value = fixedInsurance25        ' C2 = B11
    wsCal.Cells(16, 3).Value = addInsurance2025       ' C16 = B10

    ' C12 (декларация): полный годовой налог без вычетов.
    ' C13 (оплата):     налог минус оба взноса, не меньше 0.
    Dim usnTaxGross As Double
    Dim usnTaxDue   As Double
    usnTaxGross = rateUSN * prevYearIncome
    If usnTaxGross < 0 Then usnTaxGross = 0
    usnTaxDue = usnTaxGross - addInsurance2025 - fixedInsurance25
    If usnTaxDue < 0 Then usnTaxDue = 0
    wsCal.Cells(12, 3).Value = usnTaxGross  ' C12: "Подача декларации УСН за 2025"
    wsCal.Cells(13, 3).Value = usnTaxDue    ' C13: "Оплата налога УСН за 2025"

    ' 5. Минимальное форматирование: формат даты и денег + автоширина.
    wsCal.Columns("B").NumberFormat = "dd.mm.yyyy"
    wsCal.Columns("C").NumberFormat = "#,##0.00 $"
    wsCal.Columns("A:C").AutoFit

    MsgBox msgText, vbInformation, "WB_nalog_USN_NDS"

End Sub
