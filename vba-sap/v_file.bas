Attribute VB_Name = "m_ZFI_TX04_008_AI_File"
' ========================================================================
' Модуль: m_ZFI_TX04_008_AI_File   (версия «ФАЙЛ», v2.10)
' Запуск: Public Sub PDChange_File
' Назначение: Регрессивное тестирование АИ (инцидент 6100013259)
'             Выгрузка слепка таблицы ZFI_TX04_008_AI через SE16N
'             (экспорт ALV -> локальный файл по каталогу/имени, которые
'             предлагает диалог SAP) и проверки по ТЗ (Блоки 1-3).
'
' Чтение файла: ADODB + текстовый драйвер (ACE 16/12, Jet 4.0 или ODBC
' Access Text Driver) со schema.ini «как из базы SQL» — порциями GetRows,
' без 100+ МБ строки в памяти VBA. Если драйвера нет или ADO вернул не то
' число строк — построчное чтение (Line Input, а для UTF-8 — ADODB.Stream).
' Готовность файла — по стабильному размеру и эксклюзивному открытию.
'
' Основа: optimized v2.00 (оригинал 1.42). Логика проверок и макет отчёта
' идентичны версии v_buffer (m_ZFI_TX04_008_AI_Buffer).
'
' Замеры: по этапам пишутся на «Технический лист» (таблица этапов F:N,
' сводка по прогонам — с колонки P); история общая с версией buffer.
'
' Зависимости: m_sapGui (SAP_Connection, OpenTransaction) — не входит в архив.
' ВНИМАНИЕ: модуль не компилировался и не запускался. Имена полей диалога
' экспорта SAP (константы SAP_FILE_DLG_*, SAP_EXPORT_DLG_FLAG_FILE) и
' кодировку файла нужно сверить на реальном SAP GUI (см. отчёт).
' ========================================================================
Option Explicit
' ============================================================================
' Windows API
' ============================================================================
#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" _
    (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If
' ============================================================================
' Константы путей к элементам SAP (SE16N, таблица ZFI_TX04_008_AI)
' ============================================================================
Private Const SAP_SBAR              As String = "wnd[0]/sbar"
Private Const SAP_TABLE_NAME        As String = "ZFI_TX04_008_AI"
Private Const SAP_SE16N_TABLE       As String = "wnd[0]/usr/ctxtGD-TAB"
Private Const SAP_SE16N_VARIANT     As String = "wnd[0]/usr/ctxtGD-VARIANT"
Private Const SAP_SE16N_MAXLINES    As String = "wnd[0]/usr/txtGD-MAX_LINES"
Private Const SAP_SE16N_EXECUTE_BTN As String = "wnd[0]/tbar[1]/btn[8]"
Private Const SAP_SE16N_GRID        As String = "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell"
' ВНИМАНИЕ: индексы строк таблицы полей [2,N] зависят от состава полей
' варианта /REGRESSIVE — проверить при первом запуске
Private Const SAP_SE16N_FLD_GJAHR   As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]"
Private Const SAP_SE16N_FLD_MONAT   As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,3]"
Private Const SAP_SE16N_FLD_VERS    As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]"
Private Const SAP_SE16N_FLD_STEP    As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,8]"
Private Const SAP_SE16N_VERS_MULTI_BTN As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,4]"
Private Const SAP_SE16N_STEP_MULTI_BTN As String = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,8]"
Private Const SAP_SE16N_MULTI_LOW1  As String = "wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,0]"
Private Const SAP_SE16N_MULTI_LOW2  As String = "wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,1]"
Private Const SAP_SE16N_MULTI_HIGH1 As String = "wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-HIGH[2,0]"
Private Const SAP_SE16N_MULTI_OK    As String = "wnd[1]/tbar[0]/btn[8]"
' ============================================================================
' Тайминги и режимы
' ============================================================================
Private Const SAP_BUSY_TIMEOUT_SEC  As Double = 600     ' Таймаут «сессия SAP занята»
Private Const ERR_BUF_ROWS          As Long = 2000      ' Размер буфера записи таблицы ошибок
Private Const ERR_CUSTOM            As Long = vbObjectError + 513  ' Контролируемая ошибка (без кода в сообщении)
' ============================================================================
' Имена листов Excel
' ============================================================================
Private Const SHEET_REPORT       As String = "Итоговый отчет"
Private Const SHEET_DATA         As String = "ZFI_TX04_008_AI"
Private Const SHEET_TECH         As String = "Технический лист"
' ============================================================================
' Ячейки параметров на листе «Итоговый отчет» (столбцы B и C)
' ============================================================================
Private Const P_ROW_VARIANT      As Long = 1   ' B1 — Вар. просмотра (REGRESSIVE)
Private Const P_ROW_GJAHR        As Long = 2   ' B2 — Финансовый год GJAHR
Private Const P_ROW_MONAT        As Long = 3   ' B3 — Период MONAT
Private Const P_ROW_VERS         As Long = 4   ' B4 — Версия (обязат.), C4 — Версия 2
Private Const P_ROW_STEP         As Long = 5   ' B5 — Шаг «с» (обязат.), C5 — «по»
Private Const RES_ROW_FIRST      As Long = 7   ' Первая строка результатов проверок
' ============================================================================
' Параметры проверок (ТЗ)
' ============================================================================
Private Const ERR_TXT_SUM        As String = "Ошибка в этой версии – принимаем в НДД больше, чем было затрат (столбец M > столбец N)"
Private Const ERR_TXT_DATE_PAST  As String = "Ошибка в этой версии – платеж до режима НДД"
Private Const ERR_TXT_DATE_FUT   As String = "Ошибка в этой версии – платеж будущего периода"
Private Const ERR_TXT_VER_M      As String = "Ошибка сравнения версий – разная сумма для НДД (столбец M)"
Private Const ERR_TXT_VER_N      As String = "Ошибка сравнения версий – разная сумма затрат (столбец N)"
Private Const ERR_TXT_CONSIST    As String = "Ошибка консистентности – ошибка в шагах расчета"
Private Const ERR_TXT_DOC_O      As String = "Ошибка в этой версии – для документов никогда не найдутся документы дальше"
Private Const STEP_DOC_O         As String = "4"           ' Блок 1, п.4: проверка на 4-м шаге (ТЗ)
Private Const TOL                As Double = 0.01          ' Допуск сравнения сумм
Private Const MAX_ERR_LINES      As Long = 10000           ' Ограничение строк ошибок на блок
Private Const ERR_COLS           As Long = 10              ' Колонок в таблице ошибок (A:J)
Private Const ERR_COLS_MAX       As Long = 12              ' Очистка — по макс. ширине прошлых версий
' --- Заголовки колонок выгрузки (распознавание по началу текста) ---
Private Const HDR_SPP            As String = "СПП-элем"
Private Const HDR_SUM_NDD        As String = "Сумма расх"
Private Const HDR_SUM_BE         As String = "Сумма в ва"
Private Const HDR_DATE_PAY       As String = "Дата плат"
Private Const HDR_KEY_BUKRS      As String = "БЕ"
Private Const HDR_KEY_GJAHR      As String = "Финансовый год"
Private Const HDR_KEY_MONAT      As String = "Период проводки"
Private Const HDR_KEY_VERS       As String = "Версия"
Private Const HDR_KEY_BELNR      As String = "Документ СПП"
Private Const HDR_KEY_DOCLN      As String = "Строка проводки"
Private Const HDR_DOC_SF         As String = "Документ ф"
Private Const HDR_DOC_O          As String = "Документ О"
Private Const HDR_KEY_STEP       As String = "STEP"
' ============================================================================
' Структуры
' ============================================================================
Private Type TRunParams
    VariantName As String       ' B1
    Gjahr       As String       ' B2
    Monat       As String       ' B3
    Vers1       As String       ' B4
    Vers2       As String       ' C4 (необязательная)
    Step1       As String       ' B5
    Step2       As String       ' C5 (необязательная)
    arrVers     As Variant      ' список версий (1 или 2)
    arrSteps    As Variant      ' список шагов (1 значение или диапазон)
    VersLabel   As String
    StepsLabel  As String
End Type
Private Type TCols               ' индексы колонок выгрузки (0 — не найдена)
    spp     As Long
    sumNDD  As Long
    sumBE   As Long
    datePay As Long
    vers    As Long
    stp     As Long
    bukrs   As Long
    gjahr   As Long
    monat   As Long
    belnr   As Long
    docln   As Long
End Type
' ============================================================================
' Переменные уровня модуля
' ============================================================================
Private m_startTime  As Double   ' Начало выполнения макроса
Private m_stage      As String   ' Последний этап (показывается в сообщении об ошибке)
Private m_wsReport   As Worksheet
Private m_wsData     As Worksheet
Private m_wsTech     As Worksheet
Private m_hdrWritten As Boolean  ' Шапка таблицы ошибок выведена
Private m_hdrRow     As Long     ' Строка шапки таблицы ошибок
Private m_docSPPOn   As Boolean  ' Флажок «Документ СПП» включён
Private m_idxDocSPP  As Long
Private m_idxDocLN   As Long
Private m_idxDocSF   As Long
Private m_idxDocO    As Long
Private m_dtThreshold As Date    ' 01.01.2025 (Блок 1, п.2)
' --- Сохранённое состояние Excel ---
Private m_stateSaved      As Boolean
Private m_prevCalc        As Long
Private m_prevScreen      As Boolean
Private m_prevEvents      As Boolean
Private m_prevInteractive As Boolean
' --- Кэш строк (строится один раз, проверки читают только его) ---
Private m_nRows        As Long
Private m_arrVer()     As String   ' Нормализованная версия строки
Private m_arrStep()    As String   ' Нормализованный шаг строки
Private m_arrStepIdx() As Long     ' Индекс шага в arrSteps (-1 — вне диапазона)
Private m_arrGrp()     As Long     ' Номер группы (СПП / сцепка документа), 0 — пустой ключ
Private m_arrDt()      As Date     ' Дата оплаты
Private m_arrDtOk()    As Boolean  ' Дата распознана
Private m_arrSumM()    As Double   ' «Сумма расходов для НДД» (M)
Private m_arrSumN()    As Double   ' «Сумма в валюте БЕ» (N)
Private m_arrDocOEmpty()  As Boolean
Private m_arrDocSFEmpty() As Boolean
Private m_nGrp         As Long     ' Число групп
Private m_grpNames     As Variant  ' Имена групп: m_grpNames(g - 1)
Private m_dictGrp      As Object
' --- Буфер записи таблицы ошибок ---
Private m_eb()         As Variant  ' (1 To ERR_BUF_ROWS, 1 To ERR_COLS)
Private m_ebLink()     As Variant  ' (1 To ERR_BUF_ROWS, 1 To 1) формулы дриллдауна
Private m_ebKind()     As Byte     ' 1 — в строке даты, 0 — числа
Private m_ebN          As Long
Private m_ebFirstRow   As Long
' ============================================================================
' Версия «файл»: константы экспорта и чтения (МЕНЯТЬ ЗДЕСЬ, а не в коде)
' ============================================================================
Private Const MODULE_TAG As String = "file"
' --- Диалог SAP: экспорт списка (выбор формата) и «Сохранить список в файл» ---
' Пути элементов зависят от версии SAP GUI — проверить записью скрипта (Alt+F12 -> Script Recording)
Private Const SAP_EXPORT_DLG_FLAG_FILE As String = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]"   ' «Unconverted» (не буфер: [4,0])
Private Const SAP_EXPORT_DLG_OK        As String = "wnd[1]/tbar[0]/btn[0]"
Private Const SAP_FILE_DLG_PATH        As String = "wnd[1]/usr/ctxtDY_PATH"
Private Const SAP_FILE_DLG_NAME        As String = "wnd[1]/usr/ctxtDY_FILENAME"
Private Const SAP_FILE_DLG_ENCODING    As String = "wnd[1]/usr/ctxtDY_FILE_ENCODING"
Private Const SAP_FILE_DLG_OK          As String = "wnd[1]/tbar[0]/btn[0]"
' --- Каталог и имя файла: «» = брать то, что предлагает SAP в диалоге ---
Private Const EXPORT_DIR_OVERRIDE      As String = ""
Private Const EXPORT_FILE_OVERRIDE     As String = ""
Private Const EXPORT_DIR_FALLBACK      As String = "\Documents\SAP\SAP GUI\"   ' относительно %USERPROFILE%, если SAP ничего не предложил
Private Const EXPORT_NAME_FALLBACK     As String = "export.txt"
Private Const EXPORT_NAME_STAMP        As Boolean = False   ' True — добавлять метку времени к имени файла
Private Const EXPORT_ENCODING_SAP      As String = ""       ' «» — не менять; "4110" — UTF-8 (тогда FILE_CHARSET = "UTF-8")
' --- Чтение файла ---
Private Const FILE_CHARSET             As String = "ANSI"   ' ANSI | OEM | UTF-8 — кодировка файла выгрузки
Private Const FILE_READ_MODE           As String = "ADO"    ' ADO | LINES (только построчное чтение)
Private Const FILE_WAIT_SEC            As Double = 900      ' Таймаут ожидания файла
Private Const FILE_POLL_MS             As Long = 500        ' Период опроса размера файла
Private Const FILE_STABLE_TICKS        As Long = 3          ' Опросов подряд с неизменным размером
Private Const HEADER_SCAN_MAX_LINES    As Long = 2000       ' Сколько первых строк просматривать в поисках заголовка
Private Const ADO_WORK_DIR             As String = "ZFI_TX04_008_AI_ado"   ' Подкаталог %TEMP% с копией файла и schema.ini
Private Const ADO_WORK_NAME            As String = "import.txt"
Private Const ADO_CHUNK_ROWS           As Long = 20000      ' Строк за один GetRows
Private Const ADO_VERIFY_ROWCOUNT      As Boolean = True    ' Число строк ADO <> rowCount SAP -> перечитать построчно
Private Const KEEP_EXPORT_FILES        As Boolean = False   ' True — не удалять файл выгрузки и рабочие файлы ADO
Private m_fso         As Object
Private m_lrNum       As Integer
Private m_lrStream    As Object
Private m_lrUseStream As Boolean
Private m_lrOpen      As Boolean
Private m_exportPath  As String   ' Файл выгрузки SAP
Private m_workDir     As String   ' Рабочий каталог ADO
Private m_workFile    As String   ' Копия файла для ADO
Private m_adoProvider As String   ' Провайдер, через который открыт файл
Private m_adoErr      As String   ' Причины отказа ADO
Private m_readMode    As String   ' Как фактически прочитан файл (для примечания в замерах)
' ============================================================================
' Замеры времени по этапам: константы и переменные
' ============================================================================
Private Const ST_CONNECT As Long = 1    ' Подключение к SAP
Private Const ST_PREP    As Long = 2    ' Подготовка листов
Private Const ST_SE16N   As Long = 3    ' SE16N: условия выборки и запуск
Private Const ST_EXPORT  As Long = 4    ' Ожидание/формирование экспорта
Private Const ST_GET     As Long = 5    ' Получение данных
Private Const ST_PARSE   As Long = 6    ' Разбор в массив
Private Const ST_CACHE   As Long = 7    ' Кэш ключей строк
Private Const ST_WRITE   As Long = 8    ' Запись на лист данных
Private Const ST_B1      As Long = 9    ' Блок 1
Private Const ST_B2      As Long = 10   ' Блок 2
Private Const ST_B3      As Long = 11   ' Блок 3
Private Const ST_REPORT  As Long = 12   ' Вывод отчёта
Private Const ST_TOTAL   As Long = 13   ' Итого
Private Const ST_COUNT   As Long = 13
Private Const TM_COL     As Long = 6    ' Колонка F технического листа: таблица этапов
Private Const TM_SUM_COL As Long = 16   ' Колонка P технического листа: сводка по прогонам
Private m_qpcFreq     As Double         ' Частота QueryPerformanceCounter (0 — не опрошена, -1 — недоступна)
Private m_startClock  As Date           ' Время запуска (для таблицы)
Private m_stT0(1 To 13)    As Double    ' Старт текущего интервала этапа
Private m_stDur(1 To 13)   As Double    ' Накопленная длительность этапа, сек
Private m_stClock(1 To 13) As Date      ' Время первого старта этапа
Private m_stOpen(1 To 13)  As Boolean   ' Интервал открыт
Private m_stUsed(1 To 13)  As Boolean   ' Этап запускался
Private m_flushInBlock     As Double    ' Время записи буфера ошибок внутри текущего блока проверок
Private m_flushBlocksTotal As Double    ' То же за все блоки (переносится в «Вывод отчёта»)
Private m_runRows          As Long      ' Строк в выгрузке (для таблицы замеров)
' ============================================================================
' Сервисные процедуры
' ============================================================================
' ============================================================================
' Замеры времени: QueryPerformanceCounter (монотонный, без перехода через полночь);
' если недоступен — Timer с защитой от обнуления в ElapsedSec
' ============================================================================
Private Function NowSec() As Double
    Dim c As Currency
    If m_qpcFreq = 0# Then
        If QueryPerformanceFrequency(c) <> 0 Then m_qpcFreq = CDbl(c) * 10000#
        If m_qpcFreq <= 0# Then m_qpcFreq = -1#
    End If
    If m_qpcFreq > 0# Then
        If QueryPerformanceCounter(c) <> 0 Then
            NowSec = CDbl(c) * 10000# / m_qpcFreq
            Exit Function
        End If
    End If
    NowSec = Timer
End Function
Private Sub InitTimers()
    Dim i As Long
    For i = 1 To ST_COUNT
        m_stT0(i) = 0#
        m_stDur(i) = 0#
        m_stClock(i) = 0
        m_stOpen(i) = False
        m_stUsed(i) = False
    Next i
    m_flushInBlock = 0#
    m_flushBlocksTotal = 0#
    m_runRows = 0
End Sub
Private Sub TimerBegin(ByVal id As Long)
    m_stT0(id) = NowSec()
    m_stOpen(id) = True
    If Not m_stUsed(id) Then
        m_stUsed(id) = True
        m_stClock(id) = Now
    End If
End Sub
' Закрывает интервал и прибавляет его длительность к этапу (интервалы суммируются).
' Из блоков проверок вычитается время записи буфера ошибок: оно относится к «Выводу отчёта».
Private Sub TimerEnd(ByVal id As Long)
    Dim d As Double
    If Not m_stOpen(id) Then Exit Sub
    d = ElapsedSec(m_stT0(id))
    If id = ST_B1 Or id = ST_B2 Or id = ST_B3 Then
        d = d - m_flushInBlock
        If d < 0# Then d = 0#
        m_flushBlocksTotal = m_flushBlocksTotal + m_flushInBlock
        m_flushInBlock = 0#
    End If
    m_stDur(id) = m_stDur(id) + d
    m_stOpen(id) = False
End Sub
' Таблица этапов (F:N) и сводка по прогонам (P:...) на техническом листе;
' история накапливается между прогонами и между версиями модуля. Возврат — номер прогона.
Private Function WriteTimings(ByVal ok As Boolean, ByVal note As String) As Long
    Dim ws As Worksheet
    Dim i As Long, runNo As Long, lastRow As Long, r0 As Long
    Dim tot As Double
    Dim tbl() As Variant, sumRow() As Variant
    On Error GoTo WtErr
    Set ws = m_wsTech
    If ws Is Nothing Then Exit Function
    For i = 1 To ST_COUNT - 1
        If m_stOpen(i) Then TimerEnd i
    Next i
    If m_flushBlocksTotal > 0# Then
        m_stDur(ST_REPORT) = m_stDur(ST_REPORT) + m_flushBlocksTotal
        If Not m_stUsed(ST_REPORT) Then
            m_stUsed(ST_REPORT) = True
            m_stClock(ST_REPORT) = Now
        End If
        m_flushBlocksTotal = 0#
    End If
    tot = ElapsedSec(m_startTime)
    m_stDur(ST_TOTAL) = tot
    m_stClock(ST_TOTAL) = m_startClock
    m_stUsed(ST_TOTAL) = True
    If Len(CStr(ws.Cells(1, TM_COL).Value)) = 0 Then
        ws.Cells(1, TM_COL).Resize(1, 9).Value = Array("Прогон №", "Версия", "Запуск", "Этап", _
            "Начало этапа", "Длительность, сек", "% от общего", "Строк в выгрузке", "Примечание")
        ws.Cells(1, TM_COL).Resize(1, 9).Font.Bold = True
        ws.Columns(TM_COL + 3).ColumnWidth = 62
        ws.Columns(TM_COL + 2).ColumnWidth = 19
        ws.Columns(TM_COL + 8).ColumnWidth = 60
    End If
    lastRow = ws.Cells(ws.Rows.Count, TM_COL).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    runNo = 1
    If lastRow >= 2 Then
        runNo = CLng(Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, TM_COL), ws.Cells(lastRow, TM_COL)))) + 1
    End If
    r0 = lastRow + 1
    ReDim tbl(1 To ST_COUNT, 1 To 9)
    For i = 1 To ST_COUNT
        tbl(i, 1) = runNo
        tbl(i, 2) = MODULE_TAG
        tbl(i, 3) = m_startClock
        tbl(i, 4) = StageName(i)
        If m_stUsed(i) Then tbl(i, 5) = m_stClock(i)
        tbl(i, 6) = Round(m_stDur(i), 3)
        If tot > 0# Then tbl(i, 7) = m_stDur(i) / tot
        tbl(i, 8) = m_runRows
    Next i
    tbl(ST_TOTAL, 9) = IIf(ok, "OK", "ОШИБКА") & ". " & note
    With ws.Cells(r0, TM_COL).Resize(ST_COUNT, 9)
        .Value = tbl
    End With
    ws.Range(ws.Cells(r0, TM_COL + 2), ws.Cells(r0 + ST_COUNT - 1, TM_COL + 2)).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range(ws.Cells(r0, TM_COL + 4), ws.Cells(r0 + ST_COUNT - 1, TM_COL + 4)).NumberFormat = "hh:mm:ss"
    ws.Range(ws.Cells(r0, TM_COL + 5), ws.Cells(r0 + ST_COUNT - 1, TM_COL + 5)).NumberFormat = "0.00"
    ws.Range(ws.Cells(r0, TM_COL + 6), ws.Cells(r0 + ST_COUNT - 1, TM_COL + 6)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(r0, TM_COL + 7), ws.Cells(r0 + ST_COUNT - 1, TM_COL + 7)).NumberFormat = "#,##0"
    ws.Cells(r0 + ST_COUNT - 1, TM_COL).Resize(1, 9).Font.Bold = True
    ws.Cells(r0 + ST_COUNT - 1, TM_COL).Resize(1, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ' --- Сводка: одна строка на прогон (удобно сравнивать buffer и file) ---
    If Len(CStr(ws.Cells(1, TM_SUM_COL).Value)) = 0 Then
        ws.Cells(1, TM_SUM_COL).Resize(1, 4).Value = Array("Прогон №", "Версия", "Запуск", "Строк")
        For i = 1 To ST_COUNT
            ws.Cells(1, TM_SUM_COL + 3 + i).Value = StageName(i) & ", сек"
        Next i
        ws.Cells(1, TM_SUM_COL + 4 + ST_COUNT).Value = "Результат"
        ws.Cells(1, TM_SUM_COL).Resize(1, 5 + ST_COUNT).Font.Bold = True
        ws.Cells(1, TM_SUM_COL).Resize(1, 5 + ST_COUNT).WrapText = True
    End If
    lastRow = ws.Cells(ws.Rows.Count, TM_SUM_COL).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    ReDim sumRow(1 To 1, 1 To 5 + ST_COUNT)
    sumRow(1, 1) = runNo
    sumRow(1, 2) = MODULE_TAG
    sumRow(1, 3) = m_startClock
    sumRow(1, 4) = m_runRows
    For i = 1 To ST_COUNT
        sumRow(1, 4 + i) = Round(m_stDur(i), 3)
    Next i
    sumRow(1, 5 + ST_COUNT) = IIf(ok, "OK", "ОШИБКА")
    ws.Cells(lastRow + 1, TM_SUM_COL).Resize(1, 5 + ST_COUNT).Value = sumRow
    ws.Cells(lastRow + 1, TM_SUM_COL + 2).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Cells(lastRow + 1, TM_SUM_COL + 3).NumberFormat = "#,##0"
    ws.Cells(lastRow + 1, TM_SUM_COL + 4).Resize(1, ST_COUNT).NumberFormat = "0.00"
    WriteTimings = runNo
    Exit Function
WtErr:
    Application.StatusBar = "Замеры не записаны: " & Err.Description
End Function
' Заголовок -> число колонок и признаки колонок сумм
Private Sub PrepareHeader(ByRef headerParts As Variant, ByRef colCount As Long, _
    ByRef sumColFlag() As Boolean)
    Dim lastPartIdx As Long, c As Long
    lastPartIdx = UBound(headerParts)
    Do While lastPartIdx > 0
        If Trim$(CStr(headerParts(lastPartIdx))) <> "" Then Exit Do
        lastPartIdx = lastPartIdx - 1
    Loop
    colCount = lastPartIdx
    If colCount < 1 Then Fail "Строка заголовка выгрузки SE16N не содержит колонок."
    ReDim sumColFlag(1 To colCount)
    For c = 1 To colCount
        If Left$(Trim$(CStr(headerParts(c))), 5) = "Сумма" Then sumColFlag(c) = True
    Next c
End Sub
' Строка уже разбита по «|»: пишет ячейки в arrOut(tgt, ...); True, если есть непустые значения
Private Function StoreParts(ByRef parts() As String, ByVal colCount As Long, _
    ByRef sumColFlag() As Boolean, ByRef arrOut As Variant, ByVal tgt As Long) As Boolean
    Dim c As Long, n As Long, s As String
    n = UBound(parts)
    If n > colCount Then n = colCount
    For c = 1 To n
        s = Trim$(parts(c))
        If Len(s) > 0 Then
            StoreParts = True
            If sumColFlag(c) Then
                arrOut(tgt, c) = ParseSapNumber(s)
            Else
                arrOut(tgt, c) = s
            End If
        End If
    Next c
End Function
' ============================================================================
' Сервисные процедуры
' ============================================================================
Private Sub Fail(ByVal msg As String)
    Err.Raise ERR_CUSTOM, "PDChange", msg
End Sub

' Секунды с t0 с учётом перехода через полночь (NowSec() обнуляется в 00:00)
Private Function ElapsedSec(ByVal t0 As Double) As Double
    Dim d As Double
    d = NowSec() - t0
    If d < 0 Then d = d + 86400#
    ElapsedSec = d
End Function

Private Sub PauseMs(ByVal ms As Long)
    Sleep ms
    DoEvents
End Sub

Private Sub ShowProgress(ByVal pct As Long, ByVal stageTxt As String)
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    m_stage = stageTxt
    Application.DisplayStatusBar = True
    Application.StatusBar = Format(pct, "0") & "% | " & stageTxt
    DoEvents
End Sub

Private Sub SaveExcelState()
    m_prevCalc = Application.Calculation
    m_prevScreen = Application.ScreenUpdating
    m_prevEvents = Application.EnableEvents
    m_prevInteractive = Application.Interactive
    m_stateSaved = True
End Sub

Private Sub RestoreExcelState(ByVal keepStatusText As Boolean)
    On Error Resume Next
    Application.DisplayAlerts = True
    If m_stateSaved Then
        Application.Interactive = m_prevInteractive
        Application.EnableEvents = m_prevEvents
        Application.Calculation = m_prevCalc
        Application.ScreenUpdating = m_prevScreen
        m_stateSaved = False
    End If
    Application.DisplayStatusBar = True
    If Not keepStatusText Then Application.StatusBar = False
End Sub

' Единая точка очистки: и для успешного завершения, и для ошибки
Private Sub CleanupRun(ByVal keepStatusText As Boolean)
    On Error Resume Next
    LineReaderClose
    Set m_fso = Nothing
    Erase m_arrVer, m_arrStep, m_arrStepIdx, m_arrGrp, m_arrDt, m_arrDtOk, _
        m_arrSumM, m_arrSumN, m_arrDocOEmpty, m_arrDocSFEmpty
    Erase m_eb, m_ebLink, m_ebKind
    m_grpNames = Empty
    m_nGrp = 0: m_nRows = 0: m_ebN = 0
    Set m_dictGrp = Nothing
    Set m_wsData = Nothing
    Set m_wsTech = Nothing
    Set m_wsReport = Nothing
    RestoreExcelState keepStatusText
End Sub

Private Sub ActivateExcel()
    On Error Resume Next
    ShowWindow Application.hwnd, 5
    SetForegroundWindow Application.hwnd
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If Not SheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    Else
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    End If
End Function

' ============================================================================
' Работа с SAP GUI
' ============================================================================
Private Function SapStatusText(ByVal sess As Object) As String
    Dim sb As Object
    On Error Resume Next
    Set sb = sess.findById(SAP_SBAR, False)
    If Not sb Is Nothing Then SapStatusText = sb.MessageType & ": " & sb.Text
End Function

' Элемент или понятная ошибка (с путём и текстом статус-строки SAP)
Private Function SapEl(ByVal sess As Object, ByVal id As String) As Object
    Dim el As Object
    On Error Resume Next
    Set el = sess.findById(id, False)
    On Error GoTo 0
    If el Is Nothing Then
        Fail "Элемент SAP GUI не найден: " & id & ". Проверьте текущий экран SAP, вариант " & _
            "просмотра и индексы полей. Статус-строка SAP: " & SapStatusText(sess)
    End If
    Set SapEl = el
End Function

Private Sub SapSetText(ByVal sess As Object, ByVal id As String, ByVal value As String)
    SapEl(sess, id).Text = value
End Sub

Private Sub SapWaitReady(ByVal sess As Object)
    Dim t0 As Double
    t0 = NowSec()
    Do While sess.Busy
        If ElapsedSec(t0) > SAP_BUSY_TIMEOUT_SEC Then
            Fail "Сессия SAP занята более " & Format(SAP_BUSY_TIMEOUT_SEC, "0") & _
                " сек (этап: " & m_stage & ")."
        End If
        PauseMs 100
    Loop
End Sub

' Ошибочное сообщение SAP (тип E/A) прерывает выполнение с его текстом
Private Sub SapCheckStatus(ByVal sess As Object, ByVal ctx As String)
    Dim sb As Object, t As String
    On Error Resume Next
    Set sb = sess.findById(SAP_SBAR, False)
    If sb Is Nothing Then Exit Sub
    t = sb.MessageType
    On Error GoTo 0
    If t = "E" Or t = "A" Then Fail ctx & ": SAP вернул сообщение типа " & t & ": " & sb.Text
End Sub

' ============================================================================
' Разбор данных
' ============================================================================
' Ключ: убираем ведущие нули у чисто цифровых значений (без CLng/IsNumeric:
' не зависит от локали и не ломается на 10-значных номерах и «1,5»)
Private Function NormalizeSapKey(ByVal s As String) As String
    Dim t As String, n As Long, p As Long
    t = Trim$(s)
    n = Len(t)
    If n > 1 Then
        If Not (t Like "*[!0-9]*") Then
            p = 1
            Do While p < n
                If Mid$(t, p, 1) <> "0" Then Exit Do
                p = p + 1
            Loop
            If p > 1 Then t = Mid$(t, p)
        End If
    End If
    NormalizeSapKey = t
End Function

' SAP-формат «1.234.567,89-»; Val не зависит от региональных настроек
Private Function ParseSapNumber(ByVal s As String) As Double
    Dim t As String, neg As Boolean
    t = s
    If InStr(t, Chr$(160)) > 0 Then t = Replace(t, Chr$(160), "")
    If InStr(t, " ") > 0 Then t = Replace(t, " ", "")
    If InStr(t, ".") > 0 Then t = Replace(t, ".", "")
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "-" Then
        neg = True
        t = Left$(t, Len(t) - 1)
    End If
    If InStr(t, ",") > 0 Then t = Replace(t, ",", ".")
    ParseSapNumber = Val(t)
    If neg Then ParseSapNumber = -ParseSapNumber
End Function

' Дата ДД.ММ.ГГГГ без CDate/IsDate по умолчанию (зависят от локали)
Private Function ParseSapDate(ByVal s As String, ByRef ok As Boolean) As Date
    Dim d As Long, m As Long, y As Long, dt As Date
    ok = False
    If Len(s) = 0 Then Exit Function
    If s Like "##.##.####" Then
        d = CLng(Left$(s, 2)): m = CLng(Mid$(s, 4, 2)): y = CLng(Right$(s, 4))
        If m >= 1 And m <= 12 And d >= 1 And y >= 1900 Then
            dt = DateSerial(y, m, d)
            If Day(dt) = d Then
                ParseSapDate = dt
                ok = True
            End If
        End If
    Else
        On Error Resume Next
        If IsDate(s) Then
            ParseSapDate = CDate(s)
            ok = True
        End If
        On Error GoTo 0
    End If
End Function

Private Function SapCellNumber(ByRef v As Variant) As Double
    If VarType(v) = vbDouble Then
        SapCellNumber = v
    ElseIf Not IsEmpty(v) Then
        SapCellNumber = ParseSapNumber(CStr(v))
    End If
End Function

Private Sub ClearArray(ByRef arr As Variant)
    If IsArray(arr) Then Erase arr
    arr = Empty
End Sub

Private Function ResolveColumnIndex(ByRef headerParts As Variant, ByVal colCount As Long, _
    ByVal pattern As String, Optional ByVal exactMatch As Boolean = False) As Long
    Dim c As Long, h As String
    For c = 1 To colCount
        h = Trim$(CStr(headerParts(c)))
        If exactMatch Then
            If h = pattern Then ResolveColumnIndex = c: Exit Function
        Else
            If Left$(h, Len(pattern)) = pattern Then ResolveColumnIndex = c: Exit Function
        End If
    Next c
End Function

' Строка-разделитель ALV состоит только из «-» и «|»; данные с «---» в тексте — не разделитель
Private Function IsSeparatorLine(ByRef ln As String) As Boolean
    If InStr(ln, "---") = 0 Then Exit Function
    IsSeparatorLine = (Len(Replace(Replace(ln, "-", ""), "|", "")) = 0)
End Function

' ============================================================================
' Флажок «Документ СПП» и ключ группировки
' ============================================================================
Private Function GetOrCreateDocSPPCheckbox(ByVal wsReport As Worksheet) As Object
    Dim chk As Object, shp As Shape
    Dim act As String
    Dim chkLeft As Single, chkTop As Single
    On Error Resume Next
    Set chk = wsReport.CheckBoxes("chkDocSPP")
    On Error GoTo 0
    If chk Is Nothing Then
        chkLeft = -1
        For Each shp In wsReport.Shapes
            act = ""
            On Error Resume Next
            act = shp.OnAction
            On Error GoTo 0
            If InStr(1, act, "PDChange", vbTextCompare) > 0 Then
                chkLeft = shp.Left
                chkTop = shp.Top + shp.Height + 2
                Exit For
            End If
        Next shp
        If chkLeft < 0 Then
            chkLeft = wsReport.Range("D6").Left
            chkTop = wsReport.Range("D6").Top
        End If
        Set chk = wsReport.CheckBoxes.Add(chkLeft, chkTop, 110, 18)
        chk.Name = "chkDocSPP"
        chk.Caption = "Документ СПП"
    End If
    Set GetOrCreateDocSPPCheckbox = chk
End Function

' Ключ группировки строки: при флажке — сцепка «документ СПП-позиция-сч/ф» (ТЗ 1.33),
' иначе СПП-элемент. Пустая строка — строка пропускается.
Private Function GroupKeyOfRow(ByRef arrData As Variant, ByVal i As Long, ByVal idxSPP As Long) As String
    Dim docPart As String
    If m_docSPPOn And m_idxDocSPP > 0 Then
        docPart = NormalizeSapKey(CStr(arrData(i, m_idxDocSPP)))
        If Len(docPart) = 0 Then Exit Function
        GroupKeyOfRow = docPart & "-" & _
            NormalizeSapKey(CStr(arrData(i, m_idxDocLN))) & "-" & _
            NormalizeSapKey(CStr(arrData(i, m_idxDocSF)))
    Else
        GroupKeyOfRow = Trim$(CStr(arrData(i, idxSPP)))
    End If
End Function

Private Function GrpName(ByVal g As Long) As String
    GrpName = CStr(m_grpNames(g - 1))
End Function

' ============================================================================
' Кэш строк: версии, шаги, даты, суммы и номера групп — один проход
' ============================================================================
Private Sub BuildRowCache(ByRef arrData As Variant, ByVal totalRows As Long, _
    ByRef cols As TCols, ByVal arrSteps As Variant)
    Dim i As Long, st As Long, g As Long
    Dim k As String, stepKey As String
    Dim dtOk As Boolean
    Dim dictStepIdx As Object
    m_nRows = totalRows
    ReDim m_arrVer(1 To totalRows)
    ReDim m_arrStep(1 To totalRows)
    ReDim m_arrStepIdx(1 To totalRows)
    ReDim m_arrGrp(1 To totalRows)
    ReDim m_arrDt(1 To totalRows)
    ReDim m_arrDtOk(1 To totalRows)
    ReDim m_arrSumM(1 To totalRows)
    ReDim m_arrSumN(1 To totalRows)
    ReDim m_arrDocOEmpty(1 To totalRows)
    ReDim m_arrDocSFEmpty(1 To totalRows)
    Set dictStepIdx = CreateObject("Scripting.Dictionary")
    For st = LBound(arrSteps) To UBound(arrSteps)
        dictStepIdx(NormalizeSapKey(CStr(arrSteps(st)))) = st - LBound(arrSteps)
    Next st
    Set m_dictGrp = CreateObject("Scripting.Dictionary")
    For i = 1 To totalRows
        If i Mod 50000 = 0 Then
            ShowProgress 90, "Построение кэша ключей: " & Format(i, "#,##0") & _
                " из " & Format(totalRows, "#,##0")
        End If
        m_arrVer(i) = NormalizeSapKey(CStr(arrData(i, cols.vers)))
        k = GroupKeyOfRow(arrData, i, cols.spp)
        If Len(k) > 0 Then
            If m_dictGrp.Exists(k) Then
                g = m_dictGrp(k)
            Else
                g = m_dictGrp.Count + 1
                m_dictGrp.Add k, g
            End If
            m_arrGrp(i) = g
        End If
        m_arrSumM(i) = SapCellNumber(arrData(i, cols.sumNDD))
        m_arrSumN(i) = SapCellNumber(arrData(i, cols.sumBE))
        If m_idxDocO > 0 Then m_arrDocOEmpty(i) = (Len(Trim$(CStr(arrData(i, m_idxDocO)))) = 0)
        If m_idxDocSF > 0 Then m_arrDocSFEmpty(i) = (Len(Trim$(CStr(arrData(i, m_idxDocSF)))) = 0)
        m_arrStepIdx(i) = -1
        If cols.stp > 0 Then
            stepKey = NormalizeSapKey(CStr(arrData(i, cols.stp)))
            m_arrStep(i) = stepKey
            If dictStepIdx.Exists(stepKey) Then m_arrStepIdx(i) = dictStepIdx(stepKey)
        End If
        If cols.datePay > 0 Then
            m_arrDt(i) = ParseSapDate(CStr(arrData(i, cols.datePay)), dtOk)
            m_arrDtOk(i) = dtOk
        End If
    Next i
    m_nGrp = m_dictGrp.Count
    If m_nGrp > 0 Then m_grpNames = m_dictGrp.Keys
    Set m_dictGrp = Nothing
End Sub

' ============================================================================
' Таблица ошибок: шапка, буфер записи, ширины
' ============================================================================
Private Sub WriteErrTableHeader(ByVal wsReport As Worksheet, ByVal resRow As Long)
    wsReport.Cells(resRow, 1).Value = "Блок проверки"
    wsReport.Cells(resRow, 2).Value = "Версия"
    wsReport.Cells(resRow, 3).Value = "Шаг"
    If m_docSPPOn Then
        wsReport.Cells(resRow, 4).Value = "Документ СПП+поз+сч/ф"
    Else
        wsReport.Cells(resRow, 4).Value = "СПП"
    End If
    wsReport.Cells(resRow, 5).Value = "Строка"
    wsReport.Cells(resRow, 6).Value = "Значение"
    wsReport.Cells(resRow, 7).Value = "Знак"
    wsReport.Cells(resRow, 8).Value = "Эталон"
    wsReport.Cells(resRow, 9).Value = "Разница"
    wsReport.Cells(resRow, 10).Value = "Текст ошибки"
    With wsReport.Range(wsReport.Cells(resRow, 1), wsReport.Cells(resRow, ERR_COLS))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
    End With
End Sub

Private Sub SetErrColumnWidths(ByVal wsReport As Worksheet)
    wsReport.Columns("A").ColumnWidth = 16
    wsReport.Columns("B").ColumnWidth = 11
    wsReport.Columns("C").ColumnWidth = 6
    wsReport.Columns("D").ColumnWidth = 25
    wsReport.Columns("E").ColumnWidth = 9
    wsReport.Columns("F").ColumnWidth = 15
    wsReport.Columns("G").ColumnWidth = 7
    wsReport.Columns("H").ColumnWidth = 15
    wsReport.Columns("I").ColumnWidth = 15
    wsReport.Columns("J").ColumnWidth = 90
End Sub

Private Sub InitErrBuf()
    ReDim m_eb(1 To ERR_BUF_ROWS, 1 To ERR_COLS)
    ReDim m_ebLink(1 To ERR_BUF_ROWS, 1 To 1)
    ReDim m_ebKind(1 To ERR_BUF_ROWS)
    m_ebN = 0
    m_hdrWritten = False
End Sub

' Версия/шаг: число, если цифровой текст разумной длины (фильтры и сортировка)
Private Function NumOrText(ByVal s As String) As Variant
    If Len(s) = 0 Then
        NumOrText = Empty
    ElseIf Len(s) <= 9 And Not (s Like "*[!0-9]*") Then
        NumOrText = CLng(s)
    Else
        NumOrText = s
    End If
End Function

' Запись накопленных строк одним присваиванием; форматы — по диапазонам
Private Sub FlushErrBuf()
    Dim n As Long, r As Long, c As Long, runStart As Long
    Dim tmp() As Variant, tmpLink() As Variant
    Dim tFlush As Double
    n = m_ebN
    If n = 0 Then Exit Sub
    tFlush = NowSec()
    ReDim tmp(1 To n, 1 To ERR_COLS)
    ReDim tmpLink(1 To n, 1 To 1)
    For r = 1 To n
        For c = 1 To ERR_COLS
            tmp(r, c) = m_eb(r, c)
        Next c
        tmpLink(r, 1) = m_ebLink(r, 1)
    Next r
    With m_wsReport
        .Range(.Cells(m_ebFirstRow, 4), .Cells(m_ebFirstRow + n - 1, 4)).NumberFormat = "@"
        .Cells(m_ebFirstRow, 1).Resize(n, ERR_COLS).Value = tmp
        With .Cells(m_ebFirstRow, 5).Resize(n, 1)
            .Formula = tmpLink
            .Font.Color = RGB(0, 0, 255)
            .Font.Underline = xlUnderlineStyleSingle
        End With
        .Cells(m_ebFirstRow, 9).Resize(n, 1).NumberFormat = "#,##0.00"
        runStart = 1
        For r = 2 To n + 1
            If r > n Then
                ApplyKindFormat runStart, n
            ElseIf m_ebKind(r) <> m_ebKind(runStart) Then
                ApplyKindFormat runStart, r - 1
                runStart = r
            End If
        Next r
    End With
    m_flushInBlock = m_flushInBlock + ElapsedSec(tFlush)
    m_ebN = 0
End Sub

' Формат колонок «Значение»/«Эталон» для серии строк буфера с одним типом
Private Sub ApplyKindFormat(ByVal fromIdx As Long, ByVal toIdx As Long)
    Dim fmt As String
    If m_ebKind(fromIdx) = 1 Then fmt = "DD.MM.YYYY" Else fmt = "#,##0.00"
    With m_wsReport
        .Range(.Cells(m_ebFirstRow + fromIdx - 1, 6), .Cells(m_ebFirstRow + toIdx - 1, 6)).NumberFormat = fmt
        .Range(.Cells(m_ebFirstRow + fromIdx - 1, 8), .Cells(m_ebFirstRow + toIdx - 1, 8)).NumberFormat = fmt
    End With
End Sub

' Добавляет строку в таблицу ошибок (шапка — при первой строке) и сдвигает resRow
Private Sub AddErrRow(ByRef resRow As Long, ByVal blockName As String, ByVal verTxt As String, _
    ByVal stepTxt As String, ByVal key As String, ByVal rowNum As Long, ByVal valCur As Variant, _
    ByVal signTxt As String, ByVal valRef As Variant, ByVal errText As String)
    Dim n As Long, c As Long
    If Not m_hdrWritten Then
        WriteErrTableHeader m_wsReport, resRow
        m_hdrRow = resRow
        resRow = resRow + 1
        m_hdrWritten = True
    End If
    If m_ebN = 0 Then m_ebFirstRow = resRow
    m_ebN = m_ebN + 1
    n = m_ebN
    For c = 1 To ERR_COLS
        m_eb(n, c) = Empty
    Next c
    m_eb(n, 1) = blockName
    m_eb(n, 2) = NumOrText(verTxt)
    m_eb(n, 3) = NumOrText(stepTxt)
    m_eb(n, 4) = key
    If rowNum > 0 Then
        m_eb(n, 5) = rowNum
        m_ebLink(n, 1) = "=HYPERLINK(""#'" & m_wsData.Name & "'!A" & CStr(rowNum) & """," & CStr(rowNum) & ")"
    Else
        m_ebLink(n, 1) = ""
    End If
    m_ebKind(n) = 0
    If IsDate(valCur) Then
        m_eb(n, 6) = CDate(valCur)
        m_ebKind(n) = 1
    ElseIf IsNumeric(valCur) Then
        m_eb(n, 6) = CDbl(valCur)
    End If
    m_eb(n, 7) = signTxt
    If IsDate(valRef) Then
        m_eb(n, 8) = CDate(valRef)
    ElseIf IsNumeric(valRef) Then
        m_eb(n, 8) = CDbl(valRef)
    End If
    If IsNumeric(valRef) And IsNumeric(valCur) Then m_eb(n, 9) = CDbl(valRef) - CDbl(valCur)
    m_eb(n, 10) = errText
    resRow = resRow + 1
    If n >= ERR_BUF_ROWS Then FlushErrBuf
End Sub

' ============================================================================
' Блок 1 — качество данных внутри одной версии
'   п.1: итог M <> 0 И итог N < итога M
'   п.2: дата оплаты < 01.01.2025 И M в строке <> 0
'   п.3: дата оплаты позже последнего дня месяца выборки
'   п.4: 4-й шаг, «Документ О» (H) и «Документ ф» (J) пусты
' Группы — по номеру из кэша; порядок вывода — порядок первого появления.
' ============================================================================
Private Function Check1_Quality(ByVal verCur As String, ByVal dtPeriodEnd As Date, _
    ByVal resRow As Long, ByVal pctBase As Long, ByVal pctSpan As Long) As Long
    Dim nG As Long, i As Long, g As Long, idx As Long, cntErr As Long
    Dim sumM() As Double, sumN() As Double, firstRow() As Long
    Dim badRow() As Long, badDt() As Date, badStep() As String
    Dim futRow() As Long, futDt() As Date, futStep() As String
    Dim docRow() As Long
    Dim ordSum() As Long, ordBad() As Long, ordFut() As Long, ordDoc() As Long
    Dim nSum As Long, nBad As Long, nFut As Long, nDoc As Long
    Dim blockName As String
    Dim dtPay As Date
    Dim checkDocO As Boolean
    nG = m_nGrp
    If nG = 0 Then
        Check1_Quality = resRow
        Exit Function
    End If
    blockName = "Блок 1"
    checkDocO = (m_idxDocO > 0 And m_idxDocSF > 0)
    ReDim sumM(1 To nG): ReDim sumN(1 To nG): ReDim firstRow(1 To nG)
    ReDim badRow(1 To nG): ReDim badDt(1 To nG): ReDim badStep(1 To nG)
    ReDim futRow(1 To nG): ReDim futDt(1 To nG): ReDim futStep(1 To nG)
    ReDim docRow(1 To nG)
    ReDim ordSum(1 To nG): ReDim ordBad(1 To nG): ReDim ordFut(1 To nG): ReDim ordDoc(1 To nG)
    For i = 1 To m_nRows
        If i Mod 10000 = 0 Then
            ShowProgress pctBase + CLng((i / m_nRows) * pctSpan), _
                blockName & " (версия " & verCur & "): строка " & Format(i, "#,##0") & _
                " из " & Format(m_nRows, "#,##0")
        End If
        If m_arrVer(i) = verCur Then
            g = m_arrGrp(i)
            If g > 0 Then
                If firstRow(g) = 0 Then
                    firstRow(g) = i + 1
                    nSum = nSum + 1
                    ordSum(nSum) = g
                End If
                sumM(g) = sumM(g) + m_arrSumM(i)
                sumN(g) = sumN(g) + m_arrSumN(i)
                If m_arrDtOk(i) Then
                    dtPay = m_arrDt(i)
                    If dtPay < m_dtThreshold Then
                        If m_arrSumM(i) <> 0 And badRow(g) = 0 Then
                            badRow(g) = i + 1: badDt(g) = dtPay: badStep(g) = m_arrStep(i)
                            nBad = nBad + 1
                            ordBad(nBad) = g
                        End If
                    ElseIf dtPay > dtPeriodEnd Then
                        If futRow(g) = 0 Then
                            futRow(g) = i + 1: futDt(g) = dtPay: futStep(g) = m_arrStep(i)
                            nFut = nFut + 1
                            ordFut(nFut) = g
                        End If
                    End If
                End If
                If checkDocO Then
                    If m_arrStep(i) = STEP_DOC_O Then
                        If m_arrDocOEmpty(i) And m_arrDocSFEmpty(i) And docRow(g) = 0 Then
                            docRow(g) = i + 1
                            nDoc = nDoc + 1
                            ordDoc(nDoc) = g
                        End If
                    End If
                End If
            End If
        End If
    Next i
    For idx = 1 To nSum
        g = ordSum(idx)
        If Round(sumM(g), 2) <> 0 And Round(sumN(g), 2) < Round(sumM(g), 2) Then
            AddErrRow resRow, blockName, verCur, "", GrpName(g), firstRow(g), _
                sumM(g), "больше", sumN(g), ERR_TXT_SUM
            cntErr = cntErr + 1
            If cntErr >= MAX_ERR_LINES Then Exit For
        End If
    Next idx
    For idx = 1 To nBad
        g = ordBad(idx)
        AddErrRow resRow, blockName, verCur, badStep(g), GrpName(g), badRow(g), _
            badDt(g), "раньше", m_dtThreshold, ERR_TXT_DATE_PAST
        cntErr = cntErr + 1
        If cntErr >= MAX_ERR_LINES Then Exit For
    Next idx
    For idx = 1 To nFut
        g = ordFut(idx)
        AddErrRow resRow, blockName, verCur, futStep(g), GrpName(g), futRow(g), _
            futDt(g), "позже", dtPeriodEnd, ERR_TXT_DATE_FUT
        cntErr = cntErr + 1
        If cntErr >= MAX_ERR_LINES Then Exit For
    Next idx
    For idx = 1 To nDoc
        g = ordDoc(idx)
        AddErrRow resRow, blockName, verCur, STEP_DOC_O, GrpName(g), docRow(g), _
            "", "пусто", "", ERR_TXT_DOC_O
        cntErr = cntErr + 1
        If cntErr >= MAX_ERR_LINES Then Exit For
    Next idx
    Check1_Quality = resRow
End Function

' ============================================================================
' Блок 2 — сверка итогов M и N между двумя версиями
' ============================================================================
Private Function Check2_VersionsCompare(ByVal verCur As String, ByVal verPrev As String, _
    ByVal resRow As Long, ByVal pctBase As Long, ByVal pctSpan As Long) As Long
    Dim nG As Long, i As Long, g As Long, idx As Long, cntErr As Long
    Dim curM() As Double, curN() As Double, curRow() As Long
    Dim prevM() As Double, prevN() As Double, prevRow() As Long
    Dim ordCur() As Long, ordPrev() As Long
    Dim nCur As Long, nPrev As Long
    Dim ver As String
    Dim dM2 As Double, dN2 As Double
    Dim blockName As String, verLegend As String
    nG = m_nGrp
    If nG = 0 Then
        Check2_VersionsCompare = resRow
        Exit Function
    End If
    blockName = "Блок 2 (" & verCur & " vs " & verPrev & ")"
    verLegend = " – Значение: версия " & verCur & ", Эталон: версия " & verPrev
    ReDim curM(1 To nG): ReDim curN(1 To nG): ReDim curRow(1 To nG)
    ReDim prevM(1 To nG): ReDim prevN(1 To nG): ReDim prevRow(1 To nG)
    ReDim ordCur(1 To nG): ReDim ordPrev(1 To nG)
    For i = 1 To m_nRows
        If i Mod 10000 = 0 Then
            ShowProgress pctBase + CLng((i / m_nRows) * pctSpan), _
                blockName & ": строка " & Format(i, "#,##0") & " из " & Format(m_nRows, "#,##0")
        End If
        ver = m_arrVer(i)
        If ver = verCur Then
            g = m_arrGrp(i)
            If g > 0 Then
                If curRow(g) = 0 Then
                    curRow(g) = i + 1
                    nCur = nCur + 1
                    ordCur(nCur) = g
                End If
                curM(g) = curM(g) + m_arrSumM(i)
                curN(g) = curN(g) + m_arrSumN(i)
            End If
        ElseIf ver = verPrev Then
            g = m_arrGrp(i)
            If g > 0 Then
                If prevRow(g) = 0 Then
                    prevRow(g) = i + 1
                    nPrev = nPrev + 1
                    ordPrev(nPrev) = g
                End If
                prevM(g) = prevM(g) + m_arrSumM(i)
                prevN(g) = prevN(g) + m_arrSumN(i)
            End If
        End If
    Next i
    If nCur = 0 Or nPrev = 0 Then
        Check2_VersionsCompare = resRow
        Exit Function
    End If
    For idx = 1 To nCur
        g = ordCur(idx)
        dM2 = prevM(g): dN2 = prevN(g)
        If Abs(Round(curM(g) - dM2, 2)) > TOL Or Abs(Round(curN(g) - dN2, 2)) > TOL Then
            If Abs(Round(curM(g) - dM2, 2)) > TOL Then
                AddErrRow resRow, blockName, verCur, "", GrpName(g), curRow(g), _
                    curM(g), "<>", dM2, ERR_TXT_VER_M & verLegend
                cntErr = cntErr + 1
            End If
            If Abs(Round(curN(g) - dN2, 2)) > TOL And cntErr < MAX_ERR_LINES Then
                AddErrRow resRow, blockName, verCur, "", GrpName(g), curRow(g), _
                    curN(g), "<>", dN2, ERR_TXT_VER_N & verLegend
                cntErr = cntErr + 1
            End If
            If cntErr >= MAX_ERR_LINES Then Exit For
        End If
    Next idx
    If cntErr < MAX_ERR_LINES Then
        For idx = 1 To nPrev
            g = ordPrev(idx)
            If curRow(g) = 0 Then
                AddErrRow resRow, blockName, verPrev, "", GrpName(g), prevRow(g), _
                    "", "только в версии " & verPrev, "", _
                    IIf(m_docSPPOn, "Документ СПП", "СПП") & " отсутствует в версии " & verCur
                cntErr = cntErr + 1
                If cntErr >= MAX_ERR_LINES Then Exit For
            End If
        Next idx
    End If
    Check2_VersionsCompare = resRow
End Function

' ============================================================================
' Блок 3 — итоги N по группе должны совпадать на всех шагах диапазона
' ============================================================================
Private Function Check3_Consistency(ByVal verCur As String, ByVal arrSteps As Variant, _
    ByVal resRow As Long, ByVal pctBase As Long, ByVal pctSpan As Long) As Long
    Dim nG As Long, nSteps As Long, i As Long, j As Long, j0 As Long, g As Long
    Dim idx As Long, nOrd As Long, cntErr As Long, jMax As Long
    Dim sums() As Double, firstRows() As Long
    Dim seen() As Boolean, ordG() As Long
    Dim dMin As Double, dMax As Double
    Dim blockName As String
    nG = m_nGrp
    If UBound(arrSteps) = LBound(arrSteps) Or nG = 0 Then
        Check3_Consistency = resRow
        Exit Function
    End If
    nSteps = UBound(arrSteps) - LBound(arrSteps) + 1
    blockName = "Блок 3"
    ReDim sums(1 To nG, 0 To nSteps - 1)
    ReDim firstRows(1 To nG, 0 To nSteps - 1)
    ReDim seen(1 To nG)
    ReDim ordG(1 To nG)
    For i = 1 To m_nRows
        If i Mod 10000 = 0 Then
            ShowProgress pctBase + CLng((i / m_nRows) * pctSpan), _
                blockName & " (версия " & verCur & "): строка " & Format(i, "#,##0") & _
                " из " & Format(m_nRows, "#,##0")
        End If
        If m_arrVer(i) = verCur Then
            j0 = m_arrStepIdx(i)
            If j0 >= 0 Then
                g = m_arrGrp(i)
                If g > 0 Then
                    If Not seen(g) Then
                        seen(g) = True
                        nOrd = nOrd + 1
                        ordG(nOrd) = g
                    End If
                    sums(g, j0) = sums(g, j0) + m_arrSumN(i)
                    If firstRows(g, j0) = 0 Then firstRows(g, j0) = i + 1
                End If
            End If
        End If
    Next i
    For idx = 1 To nOrd
        g = ordG(idx)
        dMin = sums(g, 0): dMax = sums(g, 0)
        jMax = 0
        For j = 1 To nSteps - 1
            If sums(g, j) < dMin Then dMin = sums(g, j)
            If sums(g, j) > dMax Then dMax = sums(g, j): jMax = j
        Next j
        If Abs(Round(dMax - dMin, 2)) > TOL Then
            For j = 0 To nSteps - 1
                If j = jMax Then
                    AddErrRow resRow, blockName, verCur, CStr(arrSteps(LBound(arrSteps) + j)), _
                        GrpName(g), firstRows(g, j), sums(g, j), "макс", dMax, ERR_TXT_CONSIST
                Else
                    AddErrRow resRow, blockName, verCur, CStr(arrSteps(LBound(arrSteps) + j)), _
                        GrpName(g), firstRows(g, j), sums(g, j), "<>", dMax, ERR_TXT_CONSIST
                End If
                cntErr = cntErr + 1
                If cntErr >= MAX_ERR_LINES Then Exit For
            Next j
            If cntErr >= MAX_ERR_LINES Then Exit For
        End If
    Next idx
    Check3_Consistency = resRow
End Function

' ============================================================================
' Параметры запуска
' ============================================================================
Private Function RequireParam(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
    ByVal title As String) As String
    Dim s As String
    s = Trim$(CStr(ws.Cells(r, c).Value))
    If s = "" Then
        Fail "Не заполнено поле «" & title & "» в ячейке " & ws.Cells(r, c).Address(False, False) & _
            " листа '" & ws.Name & "'!"
    End If
    RequireParam = s
End Function

Private Sub ReadParams(ByVal wsReport As Worksheet, ByRef p As TRunParams)
    Dim stA As Long, stB As Long, stTmp As Long, st As Long
    Dim arr() As Variant
    p.VariantName = RequireParam(wsReport, P_ROW_VARIANT, 2, "Вариант просмотра")
    p.Gjahr = RequireParam(wsReport, P_ROW_GJAHR, 2, "Финансовый год")
    p.Monat = RequireParam(wsReport, P_ROW_MONAT, 2, "Период")
    p.Vers1 = RequireParam(wsReport, P_ROW_VERS, 2, "Версия")
    p.Vers2 = Trim$(CStr(wsReport.Cells(P_ROW_VERS, 3).Value))
    p.Step1 = RequireParam(wsReport, P_ROW_STEP, 2, "Шаг")
    p.Step2 = Trim$(CStr(wsReport.Cells(P_ROW_STEP, 3).Value))
    If Len(p.Gjahr) <> 4 Or p.Gjahr Like "*[!0-9]*" Then
        Fail "Финансовый год (B2) должен состоять из 4 цифр, указано: '" & p.Gjahr & "'."
    End If
    If Len(p.Monat) > 3 Or p.Monat Like "*[!0-9]*" Then
        Fail "Период (B3) должен быть числом 1-16, указано: '" & p.Monat & "'."
    End If
    If CLng(p.Monat) < 1 Or CLng(p.Monat) > 16 Then
        Fail "Период (B3) должен быть в диапазоне 1-16, указано: '" & p.Monat & "'."
    End If
    If p.Vers2 <> "" Then
        p.arrVers = Array(p.Vers1, p.Vers2)
        p.VersLabel = p.Vers1 & ", " & p.Vers2
    Else
        p.arrVers = Array(p.Vers1)
        p.VersLabel = p.Vers1
    End If
    If p.Step2 <> "" And p.Step2 <> p.Step1 Then
        If Not (p.Step1 Like "*[!0-9]*") And Not (p.Step2 Like "*[!0-9]*") And _
            Len(p.Step1) <= 3 And Len(p.Step2) <= 3 Then
            stA = CLng(p.Step1): stB = CLng(p.Step2)
            If stA > stB Then
                stTmp = stA: stA = stB: stB = stTmp
            End If
            ReDim arr(0 To stB - stA)
            For st = stA To stB
                arr(st - stA) = CStr(st)
            Next st
            p.arrSteps = arr
            p.StepsLabel = CStr(stA) & "-" & CStr(stB)
        Else
            p.arrSteps = Array(p.Step1, p.Step2)
            p.StepsLabel = p.Step1 & ", " & p.Step2
        End If
    Else
        p.arrSteps = Array(p.Step1)
        p.StepsLabel = p.Step1
    End If
End Sub

' ============================================================================
' Подготовка листов (вызывается ПОСЛЕ успешного подключения к SAP)
' ============================================================================
Private Sub PrepareSheets(ByVal wbData As Workbook, ByVal wsReport As Worksheet, _
    ByRef wsData As Worksheet)
    Dim rowLast As Long, dataSheetIdx As Long
    With wsReport
        If .AutoFilterMode Then .AutoFilterMode = False
        rowLast = .UsedRange.Row + .UsedRange.Rows.Count - 1
        If rowLast >= RES_ROW_FIRST Then
            .Range(.Cells(RES_ROW_FIRST, 1), .Cells(rowLast, ERR_COLS_MAX)).Clear
        End If
    End With
    ' Лист данных пересоздаём (а не Cells.Clear): память листа освобождается полностью
    dataSheetIdx = wbData.Sheets.Count + 1
    If SheetExists(wbData, SHEET_DATA) Then
        dataSheetIdx = wbData.Sheets(SHEET_DATA).Index
        Application.DisplayAlerts = False
        wbData.Sheets(SHEET_DATA).Delete
        Application.DisplayAlerts = True
    End If
    If dataSheetIdx > wbData.Sheets.Count Then
        Set wsData = wbData.Worksheets.Add(After:=wbData.Sheets(wbData.Sheets.Count))
    Else
        Set wsData = wbData.Worksheets.Add(Before:=wbData.Sheets(dataSheetIdx))
    End If
    wsData.Name = SHEET_DATA
    Set m_wsData = wsData
    wsReport.Activate
    m_wsTech.Range("A:D").ClearContents
End Sub

' ============================================================================
' Распознавание колонок
' ============================================================================
Private Sub ResolveColumns(ByRef headerParts As Variant, ByVal colCount As Long, _
    ByRef cols As TCols, ByVal variantName As String)
    Dim missing As String
    cols.spp = ResolveColumnIndex(headerParts, colCount, HDR_SPP)
    cols.sumNDD = ResolveColumnIndex(headerParts, colCount, HDR_SUM_NDD)
    cols.sumBE = ResolveColumnIndex(headerParts, colCount, HDR_SUM_BE)
    cols.datePay = ResolveColumnIndex(headerParts, colCount, HDR_DATE_PAY)
    If cols.datePay = 0 Then cols.datePay = ResolveColumnIndex(headerParts, colCount, "Дата проводки плат")
    If cols.datePay = 0 Then cols.datePay = ResolveColumnIndex(headerParts, colCount, "Дата оплаты")
    cols.vers = ResolveColumnIndex(headerParts, colCount, HDR_KEY_VERS)
    cols.stp = ResolveColumnIndex(headerParts, colCount, HDR_KEY_STEP, True)
    cols.bukrs = ResolveColumnIndex(headerParts, colCount, HDR_KEY_BUKRS, True)
    cols.gjahr = ResolveColumnIndex(headerParts, colCount, HDR_KEY_GJAHR)
    If cols.gjahr = 0 Then cols.gjahr = ResolveColumnIndex(headerParts, colCount, "Год", True)
    cols.monat = ResolveColumnIndex(headerParts, colCount, HDR_KEY_MONAT)
    If cols.monat = 0 Then cols.monat = ResolveColumnIndex(headerParts, colCount, "Период", True)
    cols.belnr = ResolveColumnIndex(headerParts, colCount, HDR_KEY_BELNR)
    If cols.belnr = 0 Then cols.belnr = ResolveColumnIndex(headerParts, colCount, "Доку СПП")
    cols.docln = ResolveColumnIndex(headerParts, colCount, HDR_KEY_DOCLN)
    If cols.docln = 0 Then cols.docln = ResolveColumnIndex(headerParts, colCount, "СтрПр", True)
    m_idxDocO = ResolveColumnIndex(headerParts, colCount, HDR_DOC_O)
    m_idxDocSF = ResolveColumnIndex(headerParts, colCount, HDR_DOC_SF)
    If m_docSPPOn Then
        m_idxDocSPP = cols.belnr
        m_idxDocLN = cols.docln
        If m_idxDocSPP = 0 Then missing = missing & "Документ СПП (в SE16N — «Доку СПП»); "
        If m_idxDocLN = 0 Then missing = missing & "Строка проводки («СтрПр»); "
        If m_idxDocSF = 0 Then missing = missing & "Документ ф (номер сч/ф); "
        If missing <> "" Then
            Fail "Включён флажок «Документ СПП» (анализ по сцепке документ+позиция+сч/ф), " & _
                "но в выгрузке нет колонок: " & missing & _
                "Добавьте их в вариант просмотра /" & variantName & " или снимите флажок."
        End If
    Else
        m_idxDocSPP = 0
        m_idxDocLN = 0
    End If
    missing = ""
    If cols.spp = 0 Then missing = missing & "СПП-элемент; "
    If cols.sumNDD = 0 Then missing = missing & "Сумма расходов для НДД; "
    If cols.sumBE = 0 Then missing = missing & "Сумма в валюте балансовой единицы; "
    If cols.datePay = 0 Then missing = missing & "Дата проводки платежа; "
    If cols.vers = 0 Then missing = missing & "Версия; "
    If missing <> "" Then
        Fail "Не распознаны колонки выгрузки: " & missing & "Заголовки SE16N: " & Join(headerParts, " | ")
    End If
End Sub

' ============================================================================
' Лист данных
' ============================================================================
Private Sub WriteDataSheet(ByVal wsData As Worksheet, ByRef arrData As Variant, _
    ByRef headerParts As Variant, ByVal colCount As Long, ByVal totalRows As Long, _
    ByRef sumColFlag() As Boolean, ByRef cols As TCols)
    Dim c As Long, wCol As Long
    Dim keyCols As Variant, kc As Variant
    Dim hdr() As Variant
    With wsData
        ReDim hdr(1 To 1, 1 To colCount)
        For c = 1 To colCount
            hdr(1, c) = Trim$(CStr(headerParts(c)))
        Next c
        .Cells(1, 1).Resize(1, colCount).Value = hdr
        With .Range(.Cells(1, 1), .Cells(1, colCount))
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
        End With
        keyCols = Array(cols.bukrs, cols.gjahr, cols.monat, cols.vers, cols.belnr, cols.docln, cols.stp)
        For Each kc In keyCols
            If kc > 0 Then .Cells(1, kc).Interior.Color = RGB(189, 215, 238)
        Next kc
        ' Единая запись массива (порционная запись и AutoFit по листу — медленнее; см. историю 1.36/1.41)
        .Cells(2, 1).Resize(totalRows, colCount).Value = arrData
        For c = 1 To colCount
            If sumColFlag(c) Then .Columns(c).NumberFormat = "#,##0.00"
        Next c
        .Range(.Cells(1, 1), .Cells(totalRows + 1, colCount)).AutoFilter
        For c = 1 To colCount
            wCol = Len(Trim$(CStr(headerParts(c)))) + 2
            If wCol < 10 Then wCol = 10
            If wCol > 28 Then wCol = 28
            If sumColFlag(c) And wCol < 14 Then wCol = 14
            .Columns(c).ColumnWidth = wCol
        Next c
    End With
End Sub

Private Function StageName(ByVal i As Long) As String
    StageName = Choose(i, "Подключение к SAP", "Подготовка листов (очистка отчёта, пересоздание листа данных)", _
        "SE16N: условия выборки и запуск", "Ожидание/формирование экспорта (SAP -> файл, стабильный размер)", _
        "Получение данных (заголовок, копия файла, открытие ADO/файла)", "Разбор в массив (чтение строк файла)", _
        "Построение кэша ключей строк", "Запись на лист данных", "Блок 1 (качество данных)", _
        "Блок 2 (сверка версий)", "Блок 3 (консистентность шагов)", _
        "Вывод отчёта (таблица ошибок, фильтры, ширины)", "ИТОГО")
End Function
' ============================================================================
' Файловая система
' ============================================================================
Private Function GetFso() As Object
    If m_fso Is Nothing Then Set m_fso = CreateObject("Scripting.FileSystemObject")
    Set GetFso = m_fso
End Function
Private Function WithSlash(ByVal d As String) As String
    If Len(d) > 0 And Right$(d, 1) <> "\" Then d = d & "\"
    WithSlash = d
End Function
' Файл можно открыть «эксклюзивно» — SAP закончил запись и закрыл файл
Private Function CanOpenExclusive(ByVal path As String) As Boolean
    Dim f As Integer
    f = FreeFile
    On Error Resume Next
    Open path For Binary Access Read Lock Read Write As #f
    If Err.Number = 0 Then
        Close #f
        CanOpenExclusive = True
    End If
    Err.Clear
    On Error GoTo 0
End Function
' Заголовок открытого модального окна SAP (wnd[1]) или "" — если окна нет
Private Function SapPopupText(ByVal sess As Object) As String
    Dim w As Object, t As String
    On Error Resume Next
    Set w = sess.findById("wnd[1]", False)
    If Not w Is Nothing Then
        t = w.Text
        If Len(t) = 0 Then t = "(окно без заголовка)"
        SapPopupText = t
    End If
End Function
Private Function SapReadText(ByVal sess As Object, ByVal id As String) As String
    SapReadText = CStr(SapEl(sess, id).Text)
End Function
' Каталог: переопределение -> предложенный SAP -> USERPROFILE\Documents\SAP\SAP GUI\
Private Function ChooseExportDir(ByVal proposed As String) As String
    Dim d As String
    d = Trim$(EXPORT_DIR_OVERRIDE)
    If Len(d) = 0 Then d = Trim$(proposed)
    If Len(d) = 0 Then d = Environ$("USERPROFILE") & EXPORT_DIR_FALLBACK
    ChooseExportDir = WithSlash(d)
End Function
' Имя: переопределение -> предложенное SAP -> export.txt (+ метка времени по флагу)
Private Function ChooseExportName(ByVal proposed As String) As String
    Dim n As String, dotPos As Long
    n = Trim$(EXPORT_FILE_OVERRIDE)
    If Len(n) = 0 Then n = Trim$(proposed)
    If Len(n) = 0 Then n = EXPORT_NAME_FALLBACK
    If EXPORT_NAME_STAMP Then
        dotPos = InStrRev(n, ".")
        If dotPos > 1 Then
            n = Left$(n, dotPos - 1) & "_" & Format(Now, "yyyymmdd_hhnnss") & Mid$(n, dotPos)
        Else
            n = n & "_" & Format(Now, "yyyymmdd_hhnnss")
        End If
    End If
    ChooseExportName = n
End Function
' Ждёт готовности файла: размер > 0 и неизменен FILE_STABLE_TICKS опросов подряд,
' затем файл должен открыться эксклюзивно (SAP закрыл запись)
Private Sub WaitExportFile(ByVal sess As Object, ByVal path As String, ByVal timeoutSec As Double, _
    ByVal pct As Long, ByVal stageTxt As String)
    Dim t0 As Double, sz As Double, prevSz As Double, stable As Long
    Dim fso As Object, popup As String
    Set fso = GetFso()
    t0 = NowSec()
    prevSz = -1#
    Do
        sz = -1#
        If fso.FileExists(path) Then sz = CDbl(fso.GetFile(path).Size)
        If sz > 0# Then
            If sz = prevSz Then
                stable = stable + 1
            Else
                stable = 0
                prevSz = sz
            End If
            If stable >= FILE_STABLE_TICKS Then
                If CanOpenExclusive(path) Then Exit Sub
            End If
        ElseIf ElapsedSec(t0) > 10# Then
            popup = SapPopupText(sess)
            If Len(popup) > 0 Then
                Fail "Файл выгрузки не появился, в SAP открыто окно: «" & popup & "». Путь: " & path
            End If
        End If
        ShowProgress pct, stageTxt & " (ожидание файла: " & Format(ElapsedSec(t0), "0") & _
            " сек, размер " & Format(IIf(sz > 0#, sz, 0#) / 1048576#, "0.0") & " МБ)"
        PauseMs FILE_POLL_MS
        If ElapsedSec(t0) > timeoutSec Then
            Fail "Файл выгрузки не получен за " & Format(timeoutSec, "0") & " сек: " & path & _
                ". Проверьте, что в SAP GUI не ждёт ответа окно «SAP GUI Security» " & _
                "(доступ к файлам), и что каталог доступен на запись."
        End If
    Loop
End Sub
' ============================================================================
' Построчное чтение файла (Line Input для ANSI, ADODB.Stream для других кодировок)
' ============================================================================
Private Sub LineReaderOpen(ByVal path As String)
    LineReaderClose
    If UCase$(FILE_CHARSET) = "ANSI" Then
        m_lrUseStream = False
        m_lrNum = FreeFile
        Open path For Input Access Read Shared As #m_lrNum
    Else
        m_lrUseStream = True
        Set m_lrStream = CreateObject("ADODB.Stream")
        m_lrStream.Type = 2
        m_lrStream.Charset = FILE_CHARSET
        m_lrStream.LineSeparator = -1
        m_lrStream.Open
        m_lrStream.LoadFromFile path
    End If
    m_lrOpen = True
End Sub
Private Function LineReaderRead(ByRef ln As String) As Boolean
    If Not m_lrOpen Then Exit Function
    If m_lrUseStream Then
        If m_lrStream.EOS Then Exit Function
        ln = CStr(m_lrStream.ReadText(-2))
    Else
        If EOF(m_lrNum) Then Exit Function
        Line Input #m_lrNum, ln
    End If
    LineReaderRead = True
End Function
Private Sub LineReaderClose()
    On Error Resume Next
    If m_lrOpen Then
        If m_lrUseStream Then
            m_lrStream.Close
            Set m_lrStream = Nothing
        Else
            Close #m_lrNum
        End If
    End If
    m_lrOpen = False
    On Error GoTo 0
End Sub
' Первая строка заголовка в начале файла; возврат — её Trim ("" — не найдена)
Private Function FindHeaderInFile(ByVal path As String, ByRef headerParts As Variant) As String
    Dim ln As String, n As Long
    LineReaderOpen path
    Do While LineReaderRead(ln)
        n = n + 1
        If n > HEADER_SCAN_MAX_LINES Then Exit Do
        If InStr(ln, "|") > 0 Then
            If Not IsSeparatorLine(ln) Then
                If InStr(ln, "СПП") > 0 Or InStr(ln, "Версия") > 0 Then
                    headerParts = Split(ln, "|")
                    FindHeaderInFile = Trim$(ln)
                    Exit Do
                End If
            End If
        End If
    Loop
    LineReaderClose
End Function
' Построчное чтение (резервный путь и эталон для проверки ADO)
Private Function ReadRowsLines(ByVal path As String, ByVal colCount As Long, _
    ByRef sumColFlag() As Boolean, ByVal hdrLine As String, ByVal maxRows As Long, _
    ByRef arrOut As Variant) As Long
    Dim ln As String
    Dim parts() As String
    Dim tgt As Long, rowN As Long, lineNo As Long
    Dim headerSeen As Boolean
    ReDim arrOut(1 To maxRows, 1 To colCount)
    LineReaderOpen path
    Do While LineReaderRead(ln)
        lineNo = lineNo + 1
        If lineNo Mod 20000 = 0 Then
            ShowProgress 40 + CLng(45 * rowN / maxRows), "Чтение файла построчно: " & _
                Format(rowN, "#,##0") & " из " & Format(maxRows, "#,##0") & " строк"
        End If
        If InStr(ln, "|") > 0 Then
            ln = Trim$(ln)
            If Not headerSeen Then
                If ln = hdrLine Then headerSeen = True
            ElseIf ln <> hdrLine Then
                If Not IsSeparatorLine(ln) Then
                    tgt = rowN + 1
                    If tgt > maxRows Then Exit Do
                    parts = Split(ln, "|")
                    If StoreParts(parts, colCount, sumColFlag, arrOut, tgt) Then rowN = tgt
                End If
            End If
        End If
    Loop
    LineReaderClose
    ReadRowsLines = rowN
End Function
' ============================================================================
' Чтение как из базы данных: ADODB + текстовый драйвер (ACE/Jet/ODBC) + schema.ini
' ============================================================================
Private Sub WriteSchemaIni(ByVal dirPath As String, ByVal colCount As Long)
    Dim f As Integer, c As Long, cs As String
    Select Case UCase$(FILE_CHARSET)
        Case "ANSI": cs = "ANSI"
        Case "OEM": cs = "OEM"
        Case "UTF-8", "UTF8": cs = "65001"
        Case Else: cs = FILE_CHARSET
    End Select
    f = FreeFile
    Open dirPath & "schema.ini" For Output As #f
    Print #f, "[" & ADO_WORK_NAME & "]"
    Print #f, "Format=Delimited(|)"
    Print #f, "ColNameHeader=False"
    Print #f, "MaxScanRows=0"
    Print #f, "CharacterSet=" & cs
    For c = 0 To colCount
        Print #f, "Col" & CStr(c + 1) & "=C" & CStr(c) & " Text Width 255"
    Next c
    Close #f
End Sub
' Копия файла в собственный каталог (там schema.ini не затрагивает чужие настройки)
' и открытие Recordset. False — ADO недоступен (причина в m_adoErr), будет чтение построчно.
Private Function AdoOpenExport(ByVal exportPath As String, ByVal colCount As Long, _
    ByRef conn As Object, ByRef rs As Object) As Boolean
    Dim fso As Object, workDir As String, i As Long
    Dim cs(0 To 3) As String, lbl(0 To 3) As String
    On Error GoTo EH
    m_adoProvider = ""
    m_adoErr = ""
    Set fso = GetFso()
    workDir = Environ$("TEMP")
    If Len(workDir) = 0 Then workDir = CurDir$
    workDir = WithSlash(workDir) & ADO_WORK_DIR & "\"
    If Not fso.FolderExists(workDir) Then fso.CreateFolder workDir
    m_workDir = workDir
    m_workFile = workDir & ADO_WORK_NAME
    If fso.FileExists(m_workFile) Then fso.DeleteFile m_workFile, True
    fso.CopyFile exportPath, m_workFile, True
    WriteSchemaIni workDir, colCount
    lbl(0) = "Microsoft.ACE.OLEDB.16.0"
    lbl(1) = "Microsoft.ACE.OLEDB.12.0"
    lbl(2) = "Microsoft.Jet.OLEDB.4.0"
    lbl(3) = "ODBC Access Text Driver"
    For i = 0 To 2
        cs(i) = "Provider=" & lbl(i) & ";Data Source=" & workDir & _
            ";Extended Properties=""text;HDR=No;FMT=Delimited"";"
    Next i
    cs(3) = "Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=" & workDir & _
        ";Extensions=asc,csv,tab,txt;"
    For i = 0 To 3
        Set conn = CreateObject("ADODB.Connection")
        On Error Resume Next
        conn.Open cs(i)
        If Err.Number = 0 Then
            If conn.State = 1 Then m_adoProvider = lbl(i)
        Else
            m_adoErr = m_adoErr & lbl(i) & ": " & Err.Description & " | "
        End If
        Err.Clear
        On Error GoTo EH
        If Len(m_adoProvider) > 0 Then Exit For
        Set conn = Nothing
    Next i
    If Len(m_adoProvider) = 0 Then Exit Function
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & ADO_WORK_NAME & "]", conn, 0, 1, 1
    If rs.Fields.Count < colCount + 1 Then
        m_adoErr = "драйвер вернул " & rs.Fields.Count & " полей вместо " & (colCount + 1) & " (schema.ini не применён?)"
        Exit Function
    End If
    AdoOpenExport = True
    Exit Function
EH:
    m_adoErr = m_adoErr & Err.Description
End Function
Private Sub AdoCloseAll(ByRef conn As Object, ByRef rs As Object)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
End Sub
' Строка Recordset (GetRows) совпадает с заголовком выгрузки
Private Function IsHeaderRowAdo(ByRef rowsArr As Variant, ByVal r As Long, _
    ByRef headerParts As Variant, ByVal colCount As Long) As Boolean
    Dim c As Long, v As Variant
    For c = 1 To colCount
        v = rowsArr(c, r)
        If IsNull(v) Then v = ""
        If Trim$(CStr(v)) <> Trim$(CStr(headerParts(c))) Then Exit Function
    Next c
    IsHeaderRowAdo = True
End Function
' Чтение записей порциями (GetRows) в массив (1 To maxRows, 1 To colCount).
' Поле 0 — пусто (до первого «|»), поля 1..colCount — колонки выгрузки.
Private Function ReadRowsAdo(ByVal rs As Object, ByVal colCount As Long, _
    ByRef sumColFlag() As Boolean, ByRef headerParts As Variant, ByVal maxRows As Long, _
    ByRef arrOut As Variant) As Long
    Dim rowsArr As Variant, v As Variant
    Dim r As Long, c As Long, nR As Long, tgt As Long, rowN As Long, maxLen As Long
    Dim s As String, hdr1 As String
    Dim rowOk As Boolean, isFirstMatch As Boolean, headerSeen As Boolean
    Dim nonEmpty As Boolean, dashOnly As Boolean, done As Boolean
    ReDim arrOut(1 To maxRows, 1 To colCount)
    hdr1 = Trim$(CStr(headerParts(1)))
    Do While Not rs.EOF
        rowsArr = rs.GetRows(ADO_CHUNK_ROWS)
        nR = UBound(rowsArr, 2)
        ShowProgress 40 + CLng(45 * rowN / maxRows), "Чтение файла через ADO: " & _
            Format(rowN, "#,##0") & " из " & Format(maxRows, "#,##0") & " строк"
        For r = 0 To nR
            rowOk = True
            v = rowsArr(0, r)
            If Not IsNull(v) Then
                If Len(Trim$(CStr(v))) > 0 Then rowOk = False
            End If
            If rowOk Then
                isFirstMatch = False
                v = rowsArr(1, r)
                If Not IsNull(v) Then
                    If Trim$(CStr(v)) = hdr1 Then isFirstMatch = True
                End If
                If isFirstMatch Then
                    If IsHeaderRowAdo(rowsArr, r, headerParts, colCount) Then
                        headerSeen = True
                        rowOk = False
                    End If
                End If
                If Not headerSeen Then rowOk = False
            End If
            If rowOk Then
                tgt = rowN + 1
                If tgt > maxRows Then
                    done = True
                    Exit For
                End If
                nonEmpty = False
                dashOnly = True
                maxLen = 0
                For c = 1 To colCount
                    v = rowsArr(c, r)
                    If Not IsNull(v) Then
                        s = Trim$(CStr(v))
                        If Len(s) > 0 Then
                            nonEmpty = True
                            If dashOnly Then
                                If s Like "*[!-]*" Then
                                    dashOnly = False
                                ElseIf Len(s) > maxLen Then
                                    maxLen = Len(s)
                                End If
                            End If
                            If sumColFlag(c) Then
                                arrOut(tgt, c) = ParseSapNumber(s)
                            Else
                                arrOut(tgt, c) = s
                            End If
                        End If
                    End If
                Next c
                If nonEmpty Then
                    If dashOnly And maxLen >= 3 Then
                        For c = 1 To colCount
                            arrOut(tgt, c) = Empty
                        Next c
                    Else
                        rowN = tgt
                    End If
                End If
            End If
        Next r
        If done Then Exit Do
    Loop
    ReadRowsAdo = rowN
End Function
' ============================================================================
' Файл -> массив данных + заголовок
' ============================================================================
Private Sub ReadExportFile(ByVal exportPath As String, ByVal rowCountSAP As Long, _
    ByRef arrData As Variant, ByRef headerParts As Variant, ByRef colCount As Long, _
    ByRef sumColFlag() As Boolean, ByRef totalRows As Long, ByVal variantName As String)
    Dim hdrLine As String
    Dim conn As Object, rs As Object
    Dim useAdo As Boolean
    Dim n1 As Long
    TimerBegin ST_GET
    ShowProgress 33, "Файл получен. Поиск строки заголовка..."
    hdrLine = FindHeaderInFile(exportPath, headerParts)
    If Len(hdrLine) = 0 Then
        Fail "Не удалось найти строку заголовка в файле выгрузки " & exportPath & _
            ". Проверьте вариант просмотра /" & variantName & " и кодировку файла (FILE_CHARSET = " & _
            FILE_CHARSET & ", EXPORT_ENCODING_SAP = '" & EXPORT_ENCODING_SAP & "')."
    End If
    PrepareHeader headerParts, colCount, sumColFlag
    useAdo = (UCase$(FILE_READ_MODE) = "ADO")
    If useAdo Then
        ShowProgress 36, "Открытие файла через ADODB (schema.ini, " & colCount & " колонок)..."
        useAdo = AdoOpenExport(exportPath, colCount, conn, rs)
        If Not useAdo Then AdoCloseAll conn, rs
    End If
    TimerEnd ST_GET
    TimerBegin ST_PARSE
    ShowProgress 40, "Заголовок найден (" & colCount & " колонок). Чтение " & rowCountSAP & " строк в массив..."
    If useAdo Then
        n1 = ReadRowsAdo(rs, colCount, sumColFlag, headerParts, rowCountSAP, arrData)
        AdoCloseAll conn, rs
        totalRows = n1
        m_readMode = "ADO (" & m_adoProvider & ")"
        If ADO_VERIFY_ROWCOUNT And n1 <> rowCountSAP Then
            ShowProgress 40, "ADO вернул " & Format(n1, "#,##0") & " строк вместо " & _
                Format(rowCountSAP, "#,##0") & ". Повторное чтение построчно..."
            ClearArray arrData
            totalRows = ReadRowsLines(exportPath, colCount, sumColFlag, hdrLine, rowCountSAP, arrData)
            m_readMode = "ADO -> Line Input (ADO: " & n1 & " строк, ожидалось " & rowCountSAP & _
                ", построчно: " & totalRows & ")"
        End If
    Else
        totalRows = ReadRowsLines(exportPath, colCount, sumColFlag, hdrLine, rowCountSAP, arrData)
        If UCase$(FILE_READ_MODE) = "ADO" Then
            m_readMode = "Line Input (ADO недоступен: " & m_adoErr & ")"
        Else
            m_readMode = "Line Input"
        End If
    End If
    TimerEnd ST_PARSE
    If totalRows = 0 Then
        Fail "Файл выгрузки прочитан, но строки данных не распознаны (" & m_readMode & "). " & _
            "Проверьте вариант просмотра /" & variantName & " и кодировку файла."
    End If
End Sub
' Удаление файла выгрузки и рабочих файлов ADO (финансовые данные не должны оставаться на диске)
Private Sub FinishExportFiles()
    Dim fso As Object
    If KEEP_EXPORT_FILES Then Exit Sub
    On Error Resume Next
    Set fso = GetFso()
    If Len(m_exportPath) > 0 Then fso.DeleteFile m_exportPath, True
    If Len(m_workFile) > 0 Then fso.DeleteFile m_workFile, True
    If Len(m_workDir) > 0 Then fso.DeleteFile m_workDir & "schema.ini", True
    m_exportPath = ""
    m_workFile = ""
    On Error GoTo 0
End Sub
' ============================================================================
' Экспорт SE16N -> файл (каталог и имя — из диалога SAP)
' ============================================================================
Private Sub ExportSe16nToFile(ByVal sess As Object, ByRef p As TRunParams, _
    ByRef exportPath As String, ByRef rowCountSAP As Long)
    Dim grid As Object, fso As Object
    Dim waitStage As String
    Dim propDir As String, propName As String, dirPath As String, fileName As String
    TimerBegin ST_SE16N
    ShowProgress 10, "SE16N: открытие транзакции (версия " & p.VersLabel & ", шаг " & p.StepsLabel & ")..."
    OpenTransaction "/nse16n"
    SapWaitReady sess
    ShowProgress 11, "SE16N: загрузка таблицы " & SAP_TABLE_NAME & "..."
    On Error Resume Next
    sess.findById(SAP_SE16N_TABLE).Text = SAP_TABLE_NAME
    On Error GoTo 0
    SapEl(sess, "wnd[0]").sendVKey 0
    SapWaitReady sess
    ShowProgress 12, "SE16N: загрузка варианта /" & p.VariantName & "..."
    SapSetText sess, SAP_SE16N_VARIANT, "/" & p.VariantName
    SapEl(sess, "wnd[0]").sendVKey 0
    SapWaitReady sess
    SapCheckStatus sess, "Загрузка варианта /" & p.VariantName
    On Error Resume Next
    sess.findById(SAP_SE16N_MAXLINES).Text = ""
    On Error GoTo 0
    ShowProgress 15, "SE16N: заполнение полей выбора (год " & p.Gjahr & ", период " & p.Monat & _
        ", версия " & p.VersLabel & ")..."
    SapSetText sess, SAP_SE16N_FLD_GJAHR, p.Gjahr
    SapSetText sess, SAP_SE16N_FLD_MONAT, p.Monat
    If p.Vers2 <> "" Then
        ShowProgress 16, "SE16N: множественный выбор версий " & p.Vers1 & " и " & p.Vers2 & "..."
        SapEl(sess, SAP_SE16N_VERS_MULTI_BTN).SetFocus
        SapEl(sess, SAP_SE16N_VERS_MULTI_BTN).press
        SapSetText sess, SAP_SE16N_MULTI_LOW1, p.Vers1
        SapSetText sess, SAP_SE16N_MULTI_LOW2, p.Vers2
        SapEl(sess, SAP_SE16N_MULTI_OK).press
    Else
        SapSetText sess, SAP_SE16N_FLD_VERS, p.Vers1
    End If
    If p.Step2 <> "" And p.Step2 <> p.Step1 Then
        ShowProgress 17, "SE16N: множественный выбор шагов " & p.StepsLabel & "..."
        SapEl(sess, SAP_SE16N_STEP_MULTI_BTN).SetFocus
        SapEl(sess, SAP_SE16N_STEP_MULTI_BTN).press
        SapSetText sess, SAP_SE16N_MULTI_LOW1, p.Step1
        SapSetText sess, SAP_SE16N_MULTI_HIGH1, p.Step2
        SapEl(sess, SAP_SE16N_MULTI_OK).press
    Else
        SapSetText sess, SAP_SE16N_FLD_STEP, p.Step1
    End If
    SapWaitReady sess
    ShowProgress 20, "SE16N: выполнение выборки..."
    SapEl(sess, SAP_SE16N_EXECUTE_BTN).press
    SapWaitReady sess
    SapCheckStatus sess, "Выполнение выборки"
    ShowProgress 22, "SE16N: чтение результата выборки..."
    Set grid = SapEl(sess, SAP_SE16N_GRID)
    rowCountSAP = grid.rowCount
    If rowCountSAP <= 0 Then Fail "Отчёт SE16N пуст по указанным условиям выборки. Нет данных для обработки."
    If rowCountSAP + 1 > m_wsData.Rows.Count Then
        Fail "В выгрузке " & Format(rowCountSAP, "#,##0") & " строк — больше лимита листа Excel (" & _
            Format(m_wsData.Rows.Count - 1, "#,##0") & "). Сузьте выборку (шаги/версии/период)."
    End If
    TimerEnd ST_SE16N
    TimerBegin ST_EXPORT
    waitStage = "SE16N: выборка готова (" & rowCountSAP & " строк). Экспорт в файл..."
    ShowProgress 23, waitStage
    grid.pressToolbarContextButton "&MB_EXPORT"
    grid.selectContextMenuItem "&PC"
    SapEl(sess, SAP_EXPORT_DLG_FLAG_FILE).Select
    SapEl(sess, SAP_EXPORT_DLG_FLAG_FILE).SetFocus
    SapEl(sess, SAP_EXPORT_DLG_OK).press
    SapWaitReady sess
    ' --- Диалог «Сохранить список в файл»: каталог и имя, предложенные SAP ---
    propDir = SapReadText(sess, SAP_FILE_DLG_PATH)
    propName = SapReadText(sess, SAP_FILE_DLG_NAME)
    dirPath = ChooseExportDir(propDir)
    fileName = ChooseExportName(propName)
    Set fso = GetFso()
    If Not fso.FolderExists(dirPath) Then
        Fail "Каталог экспорта не существует: " & dirPath & " (предложено SAP: '" & propDir & _
            "'). Задайте EXPORT_DIR_OVERRIDE."
    End If
    exportPath = dirPath & fileName
    If fso.FileExists(exportPath) Then
        On Error Resume Next
        fso.DeleteFile exportPath, True
        On Error GoTo 0
        If fso.FileExists(exportPath) Then
            Fail "Не удалось удалить старый файл выгрузки (открыт в другой программе?): " & exportPath
        End If
    End If
    m_exportPath = exportPath
    SapSetText sess, SAP_FILE_DLG_PATH, dirPath
    SapSetText sess, SAP_FILE_DLG_NAME, fileName
    If Len(EXPORT_ENCODING_SAP) > 0 Then
        On Error Resume Next
        sess.findById(SAP_FILE_DLG_ENCODING).Text = EXPORT_ENCODING_SAP
        On Error GoTo 0
    End If
    SapEl(sess, SAP_FILE_DLG_OK).press
    SapWaitReady sess
    SapCheckStatus sess, "Экспорт в файл"
    WaitExportFile sess, exportPath, FILE_WAIT_SEC, 23, waitStage
    TimerEnd ST_EXPORT
End Sub
' ============================================================================
' ОСНОВНАЯ ПРОЦЕДУРА (версия «файл»)
' Алгоритм: параметры -> подключение к SAP -> очистка листов -> SE16N + экспорт в файл
'   -> ожидание файла -> чтение (ADO / Line Input) -> колонки -> кэш -> лист данных
'   -> проверки (Блоки 1-3) -> отчёт -> замеры
' ============================================================================
Public Sub PDChange_File()
    Dim p As TRunParams
    Dim cols As TCols
    Dim wbData As Workbook, wsReport As Worksheet, wsData As Worksheet
    Dim sess As Object, chkDocSPP As Object
    Dim exportPath As String
    Dim arrData As Variant, headerParts As Variant
    Dim sumColFlag() As Boolean
    Dim rowCountSAP As Long, totalRows As Long, colCount As Long
    Dim totalSec As Double
    Dim resRow As Long, dtPeriodEnd As Date
    Dim nVers As Long, nChecks As Long, kCheck As Long, vIdx As Long, runNo As Long
    Dim errNum As Long, errDesc As String, errMsgFull As String
    On Error GoTo ErrorHandler
    m_startTime = NowSec()
    m_startClock = Now
    InitTimers
    m_stage = ""
    m_exportPath = ""
    m_workFile = ""
    m_workDir = ""
    m_readMode = ""
    m_dtThreshold = DateSerial(2025, 1, 1)
    SaveExcelState
    ShowProgress 0, "Подготовка: чтение параметров с листа '" & SHEET_REPORT & "'..."
    Set wbData = ActiveWorkbook
    If wbData Is Nothing Then Fail "Не открыта рабочая книга!"
    If Not SheetExists(wbData, SHEET_REPORT) Then
        Fail "В активной книге '" & wbData.Name & "' нет листа '" & SHEET_REPORT & "'. " & _
            "Активируйте книгу с отчётом и запустите макрос снова."
    End If
    Set wsReport = wbData.Worksheets(SHEET_REPORT)
    Set m_wsReport = wsReport
    Set m_wsTech = GetOrCreateSheet(wbData, SHEET_TECH)
    Set chkDocSPP = GetOrCreateDocSPPCheckbox(wsReport)
    m_docSPPOn = (chkDocSPP.Value = xlOn)
    ReadParams wsReport, p
    dtPeriodEnd = DateSerial(CInt(p.Gjahr), IIf(CLng(p.Monat) > 12, 12, CLng(p.Monat)) + 1, 0)
    ShowProgress 2, "Параметры приняты (версии " & p.VersLabel & ", шаги " & p.StepsLabel & "). Подключение к SAP..."
    TimerBegin ST_CONNECT
    Set sess = SAP_Connection()
    TimerEnd ST_CONNECT
    If sess Is Nothing Then
        Fail "Не удалось подключиться к SAP (SAP_Connection вернула Nothing). " & _
            "Проверьте, что SAP GUI запущен, выполнен вход и включён Scripting."
    End If
    TimerBegin ST_PREP
    PrepareSheets wbData, wsReport, wsData
    TimerEnd ST_PREP
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.Interactive = False
    ShowProgress 5, "SAP подключён. Работа с SE16N..."
    ExportSe16nToFile sess, p, exportPath, rowCountSAP
    ReadExportFile exportPath, rowCountSAP, arrData, headerParts, colCount, sumColFlag, totalRows, p.VariantName
    FinishExportFiles
    m_runRows = totalRows
    ShowProgress 89, "Данные разобраны (" & Format(totalRows, "#,##0") & " строк). Распознавание колонок..."
    ResolveColumns headerParts, colCount, cols, p.VariantName
    ShowProgress 90, "Колонки распознаны: СПП=" & cols.spp & ", M=" & cols.sumNDD & _
        ", N=" & cols.sumBE & ", Дата=" & cols.datePay & ", Версия=" & cols.vers
    TimerBegin ST_CACHE
    BuildRowCache arrData, totalRows, cols, p.arrSteps
    TimerEnd ST_CACHE
    ShowProgress 91, "Выгрузка данных на лист (" & Format(totalRows, "#,##0") & " строк x " & colCount & " колонок)..."
    TimerBegin ST_WRITE
    WriteDataSheet wsData, arrData, headerParts, colCount, totalRows, sumColFlag, cols
    ClearArray arrData
    TimerEnd ST_WRITE
    ShowProgress 92, "Выполнение проверок (Блоки 1-3)..."
    InitErrBuf
    TimerBegin ST_REPORT
    resRow = RES_ROW_FIRST
    wsReport.Cells(resRow, 1).Value = "РЕЗУЛЬТАТЫ ПРОВЕРОК"
    wsReport.Cells(resRow, 1).Font.Bold = True
    resRow = resRow + 1
    wsReport.Cells(resRow, 1).Value = "Выгружено строк: " & totalRows & _
        " | Версии: " & p.VersLabel & " | Шаги: " & p.StepsLabel & _
        " | Год/период: " & p.Gjahr & "/" & p.Monat & _
        IIf(m_docSPPOn, " | Документ СПП: вкл", "")
    resRow = resRow + 2
    TimerEnd ST_REPORT
    nVers = UBound(p.arrVers) - LBound(p.arrVers) + 1
    nChecks = 2 * nVers
    If nVers > 1 Then nChecks = nChecks + nVers - 1
    kCheck = 0
    For vIdx = LBound(p.arrVers) To UBound(p.arrVers)
        TimerBegin ST_B1
        resRow = Check1_Quality(CStr(p.arrVers(vIdx)), dtPeriodEnd, resRow, _
            92 + (6 * kCheck) \ nChecks, 6 \ nChecks + 1)
        TimerEnd ST_B1
        kCheck = kCheck + 1
    Next vIdx
    If nVers > 1 Then
        For vIdx = LBound(p.arrVers) To UBound(p.arrVers) - 1
            TimerBegin ST_B2
            resRow = Check2_VersionsCompare(CStr(p.arrVers(vIdx + 1)), CStr(p.arrVers(vIdx)), resRow, _
                92 + (6 * kCheck) \ nChecks, 6 \ nChecks + 1)
            TimerEnd ST_B2
            kCheck = kCheck + 1
        Next vIdx
    End If
    For vIdx = LBound(p.arrVers) To UBound(p.arrVers)
        TimerBegin ST_B3
        resRow = Check3_Consistency(CStr(p.arrVers(vIdx)), p.arrSteps, resRow, _
            92 + (6 * kCheck) \ nChecks, 6 \ nChecks + 1)
        TimerEnd ST_B3
        kCheck = kCheck + 1
    Next vIdx
    TimerBegin ST_REPORT
    FlushErrBuf
    If m_hdrWritten Then
        wsReport.Range(wsReport.Cells(m_hdrRow, 1), wsReport.Cells(resRow - 1, ERR_COLS)).AutoFilter
    Else
        wsReport.Cells(resRow, 1).Value = "Ошибок не выявлено"
        wsReport.Cells(resRow, 1).Font.Bold = True
        resRow = resRow + 1
    End If
    ShowProgress 99, "Финальное оформление отчёта (ширины колонок, фильтры)..."
    SetErrColumnWidths wsReport
    TimerEnd ST_REPORT
    runNo = WriteTimings(True, "Версии " & p.VersLabel & "; шаги " & p.StepsLabel & "; период " & _
        p.Gjahr & "/" & p.Monat & "; Документ СПП: " & IIf(m_docSPPOn, "вкл", "выкл") & _
        "; экспорт в файл; чтение: " & m_readMode)
    totalSec = ElapsedSec(m_startTime)
    Set sess = Nothing
    CleanupRun True
    wsReport.Calculate
    ActivateExcel
    wsReport.Activate
    ShowProgress 100, "Готово (file, прогон " & runNo & "): " & Round(totalSec, 1) & " сек"
    MsgBox "Выполнено успешно (версия file, прогон № " & runNo & ", " & Format(totalSec, "0.0") & _
        " сек). Результаты проверок — на листе '" & SHEET_REPORT & "', замеры — на листе '" & SHEET_TECH & "'.", vbInformation
    Exit Sub
ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    errMsgFull = "Описание: " & errDesc
    If errNum <> ERR_CUSTOM And errNum <> 0 Then errMsgFull = errMsgFull & vbCrLf & "Код ошибки: " & errNum
    If Len(m_stage) > 0 Then errMsgFull = errMsgFull & vbCrLf & "Этап: " & m_stage
    If Len(m_exportPath) > 0 Then errMsgFull = errMsgFull & vbCrLf & "Файл выгрузки (не удалён для диагностики): " & m_exportPath
    runNo = WriteTimings(False, errDesc & " | этап: " & m_stage)
    Set sess = Nothing
    CleanupRun False
    MsgBox "Произошла ошибка при выполнении макроса:" & vbCrLf & errMsgFull, vbCritical, "Ошибка выполнения"
    ActivateExcel
End Sub
