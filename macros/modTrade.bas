Option Explicit
' ===== CryptoJournal. Макрос ЗАПИСАТЬ для раздела «Торговля» (шаг 1) =====
' Лист «Ввод»: Дата B5, Актив 1 C5, Актив 2 D5, Пара E5, Протокол F5,
' Торговля: Действие C8, Кол-во C9, Курс C10, ID позиции C12, Кол-во закрыть C13,
' Комиссия E17. Счётчик ID: Справочники!G5.

Private Const SH_IN As String = "Ввод"
Private Const SH_T As String = "Торговля"
Private Const SH_C As String = "Закрытые"
Private Const SH_R As String = "Справочники"
Private Const FIRST_ROW As Long = 5
Private Const LAST_ROW As Long = 300
Private Const EPS As Double = 0.000000001

' ---------- служебные функции ----------
Private Function IsNum(ByVal v As Variant) As Boolean
    IsNum = (VarType(v) = vbDouble)
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsEmpty(v) Then
        Txt = ""
    ElseIf IsError(v) Then
        Txt = ""
    Else
        Txt = Trim(CStr(v))
    End If
End Function

Private Function R8(ByVal x As Double) As Double
    R8 = Application.WorksheetFunction.Round(x, 8)
End Function

Private Function NormId(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Then
        NormId = ""
    ElseIf VarType(v) = vbDouble Then
        NormId = "T-" & Format(v, "0000")
    Else
        s = UCase(Trim(CStr(v)))
        If Len(s) > 0 Then
            If IsNumeric(s) Then s = "T-" & Format(CDbl(s), "0000")
        End If
        NormId = s
    End If
End Function

Private Function AnyFilled(ByVal rng As Range) As Boolean
    Dim c As Range
    For Each c In rng.Cells
        If Len(Txt(c.Value2)) > 0 Then AnyFilled = True: Exit Function
    Next c
End Function

Private Function FindFreeRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = FIRST_ROW To LAST_ROW
        If Len(Txt(ws.Cells(r, 2).Value2)) = 0 And Len(Txt(ws.Cells(r, 6).Value2)) = 0 _
           And Len(Txt(ws.Cells(r, 9).Value2)) = 0 Then
            FindFreeRow = r: Exit Function
        End If
    Next r
    FindFreeRow = 0
End Function

Private Function FindId(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim r As Long
    For r = FIRST_ROW To LAST_ROW
        If UCase(Txt(ws.Cells(r, 2).Value2)) = id Then FindId = r: Exit Function
    Next r
    FindId = 0
End Function

Private Sub Fail(ByVal msg As String)
    MsgBox msg, vbExclamation, "ЗАПИСАТЬ — Торговля"
End Sub

' ---------- кнопка ЗАПИСАТЬ ----------
Public Sub WriteTrade()
    Dim wsIn As Worksheet, wsT As Worksheet, wsC As Worksheet, wsR As Worksheet
    Dim d As Variant, pair As String, proto As String, act As String
    Dim qty As Double, rate As Double, fee As Double, eff As Double
    Dim idRaw As Variant, id As String, closeQ As Double
    Dim pr As Long, cr As Long, nr As Long
    Dim posQty As Double, posRate As Double, posBuy As Boolean, actBuy As Boolean
    Dim newQty As Double, newRate As Double, remQ As Double
    Dim cnt As Long, resultMsg As String

    On Error GoTo EH
    Set wsIn = ThisWorkbook.Worksheets(SH_IN)
    Set wsT = ThisWorkbook.Worksheets(SH_T)
    Set wsC = ThisWorkbook.Worksheets(SH_C)
    Set wsR = ThisWorkbook.Worksheets(SH_R)

    ' --- 1. Заполнен только раздел «Торговля» ---
    If AnyFilled(wsIn.Range("E8:E10")) Or AnyFilled(wsIn.Range("E12:E13")) Or AnyFilled(wsIn.Range("G8:G13")) Then
        Fail "Заполнено больше одного раздела. Очистите Лэндинг и Пулы, оставьте только Торговлю."
        Exit Sub
    End If

    ' --- 2. Чтение и проверка полей ---
    d = wsIn.Range("B5").Value2
    If Not IsNum(d) Then Fail "Не указана Дата (B5).": Exit Sub
    pair = UCase(Txt(wsIn.Range("E5").Value2))
    If Len(pair) = 0 Then Fail "Не заполнены Актив 1 и Актив 2 (Пара).": Exit Sub
    proto = Txt(wsIn.Range("F5").Value2)
    act = Txt(wsIn.Range("C8").Value2)
    If StrComp(act, "Buy", vbTextCompare) = 0 Then
        actBuy = True
    ElseIf StrComp(act, "Sell", vbTextCompare) = 0 Then
        actBuy = False
    Else
        Fail "Выберите Действие: Buy или Sell.": Exit Sub
    End If
    If Not IsNum(wsIn.Range("C9").Value2) Then Fail "Не указано Кол-во.": Exit Sub
    qty = wsIn.Range("C9").Value2
    If qty <= 0 Then Fail "Кол-во должно быть больше нуля.": Exit Sub
    If Not IsNum(wsIn.Range("C10").Value2) Then Fail "Не указан Курс.": Exit Sub
    rate = wsIn.Range("C10").Value2
    If rate <= 0 Then Fail "Курс должен быть больше нуля.": Exit Sub
    If Not IsNum(wsIn.Range("E17").Value2) Then Fail "Не указана Комиссия (E17).": Exit Sub
    fee = wsIn.Range("E17").Value2
    If fee < 0 Or fee >= 100 Then Fail "Комиссия должна быть от 0 до 100 (в процентах).": Exit Sub

    If actBuy Then
        eff = R8(rate * (1 + fee / 100))
    Else
        eff = R8(rate * (1 - fee / 100))
    End If

    idRaw = wsIn.Range("C12").Value2
    id = NormId(idRaw)

    ' ===== РЕЖИМ: ID пусто — новая позиция =====
    If Len(id) = 0 Then
        If Len(proto) = 0 Then Fail "Не выбран Протокол.": Exit Sub
        nr = FindFreeRow(wsT)
        If nr = 0 Then Fail "В листе «Торговля» нет свободных строк (до строки 300).": Exit Sub
        If Not IsNum(wsR.Range("G5").Value2) Then Fail "Счётчик ID (Справочники!G5) не число.": Exit Sub
        cnt = CLng(wsR.Range("G5").Value2) + 1
        id = "T-" & Format(cnt, "0000")
        If FindId(wsT, id) > 0 Or FindId(wsC, id) > 0 Then
            Fail "ID " & id & " уже есть в журнале. Проверьте счётчик в Справочники!G5."
            Exit Sub
        End If
        wsT.Cells(nr, 2).Value2 = id
        wsT.Cells(nr, 3).Value2 = d
        wsT.Cells(nr, 4).Value2 = proto
        wsT.Cells(nr, 5).Value2 = pair
        If actBuy Then
            wsT.Cells(nr, 6).Value2 = qty
            wsT.Cells(nr, 7).Value2 = eff
        Else
            wsT.Cells(nr, 9).Value2 = qty
            wsT.Cells(nr, 10).Value2 = eff
        End If
        wsR.Range("G5").Value2 = cnt
        resultMsg = "Открыта новая позиция " & id & " (" & pair & ", " & act & " " & qty & " по " & eff & ")."
        GoTo Done
    End If

    ' ===== ID заполнен: ищем позицию =====
    pr = FindId(wsT, id)
    If pr = 0 Then Fail "ID " & id & " не найден в листе «Торговля» (позиция закрыта или ID неверный).": Exit Sub
    If UCase(Txt(wsT.Cells(pr, 5).Value2)) <> pair Then
        Fail "Пара на «Вводе» (" & pair & ") не совпадает с парой позиции " & id & " (" & Txt(wsT.Cells(pr, 5).Value2) & ")."
        Exit Sub
    End If
    If IsNum(wsT.Cells(pr, 6).Value2) Then
        If wsT.Cells(pr, 6).Value2 > 0 Then
            posBuy = True: posQty = wsT.Cells(pr, 6).Value2: posRate = wsT.Cells(pr, 7).Value2
        End If
    End If
    If Not posBuy Then
        If IsNum(wsT.Cells(pr, 9).Value2) Then
            posQty = wsT.Cells(pr, 9).Value2: posRate = wsT.Cells(pr, 10).Value2
        End If
    End If
    If posQty <= 0 Or posRate <= 0 Then Fail "В строке позиции " & id & " нет Кол-ва или Курса.": Exit Sub

    ' ===== РЕЖИМ: то же Действие — увеличение =====
    If actBuy = posBuy Then
        newQty = posQty + qty
        newRate = R8((posQty * posRate + qty * eff) / newQty)
        wsT.Cells(pr, 3).Value2 = d
        If posBuy Then
            wsT.Cells(pr, 6).Value2 = newQty
            wsT.Cells(pr, 7).Value2 = newRate
        Else
            wsT.Cells(pr, 9).Value2 = newQty
            wsT.Cells(pr, 10).Value2 = newRate
        End If
        resultMsg = "Позиция " & id & " увеличена: Кол-во " & newQty & ", средний курс " & newRate & "."
        GoTo Done
    End If

    ' ===== РЕЖИМ: противоположное Действие — закрытие =====
    If IsNum(wsIn.Range("C13").Value2) Then
        closeQ = wsIn.Range("C13").Value2
    ElseIf Len(Txt(wsIn.Range("C13").Value2)) = 0 Then
        closeQ = qty
    Else
        Fail "«Кол-во закрыть» должно быть числом.": Exit Sub
    End If
    If closeQ <= 0 Then Fail "«Кол-во закрыть» должно быть больше нуля.": Exit Sub
    If closeQ > posQty + EPS Then
        Fail "«Кол-во закрыть» (" & closeQ & ") больше остатка позиции " & id & " (" & posQty & ")."
        Exit Sub
    End If
    cr = FindFreeRow(wsC)
    If cr = 0 Then Fail "В листе «Закрытые» нет свободных строк (до строки 300).": Exit Sub

    wsC.Cells(cr, 2).Value2 = id
    wsC.Cells(cr, 3).Value2 = d
    wsC.Cells(cr, 4).Value2 = Txt(wsT.Cells(pr, 4).Value2)
    wsC.Cells(cr, 5).Value2 = pair
    If posBuy Then
        wsC.Cells(cr, 6).Value2 = closeQ
        wsC.Cells(cr, 7).Value2 = posRate
        wsC.Cells(cr, 9).Value2 = qty
        wsC.Cells(cr, 10).Value2 = eff
    Else
        wsC.Cells(cr, 9).Value2 = closeQ
        wsC.Cells(cr, 10).Value2 = posRate
        wsC.Cells(cr, 6).Value2 = qty
        wsC.Cells(cr, 7).Value2 = eff
    End If

    remQ = R8(posQty - closeQ)
    If remQ <= EPS Then
        ' полное закрытие: строка уходит из «Торговли» (формулы остаются)
        wsT.Cells(pr, 2).ClearContents
        wsT.Cells(pr, 3).ClearContents
        wsT.Cells(pr, 4).ClearContents
        wsT.Cells(pr, 5).ClearContents
        wsT.Cells(pr, 6).ClearContents
        wsT.Cells(pr, 7).ClearContents
        wsT.Cells(pr, 9).ClearContents
        wsT.Cells(pr, 10).ClearContents
        resultMsg = "Позиция " & id & " закрыта полностью (строка перенесена в «Закрытые»)."
    Else
        If posBuy Then
            wsT.Cells(pr, 6).Value2 = remQ
        Else
            wsT.Cells(pr, 9).Value2 = remQ
        End If
        resultMsg = "Позиция " & id & " закрыта частично, остаток " & remQ & " (курс остатка не изменён)."
    End If

Done:
    ClearInput
    MsgBox resultMsg, vbInformation, "ЗАПИСАТЬ — Торговля"
    Exit Sub
EH:
    MsgBox "Ошибка записи: " & Err.Description & vbCrLf & _
           "Проверьте, что формулы листов не затронуты и защита листов включена по умолчанию (без пароля).", _
           vbCritical, "ЗАПИСАТЬ — Торговля"
End Sub

' ---------- кнопка ОЧИСТИТЬ (и очистка после записи) ----------
' Очищает окна ввода всех разделов. Дата (B5), Комиссия (E17) и блок «Обратный расчёт» остаются.
Public Sub ClearInput()
    Dim ws As Worksheet, a As Variant
    Set ws = ThisWorkbook.Worksheets(SH_IN)
    For Each a In Array("C5", "D5", "F5", "C8", "C9", "C10", "C12", "C13", _
                        "E8", "E9", "E10", "E12", "E13", _
                        "G8", "G9", "G10", "G11", "G12", "G13")
        ws.Range(CStr(a)).ClearContents
    Next a
End Sub

' ---------- разовая установка кнопок на лист «Ввод» ----------
Public Sub InstallButtons()
    Dim ws As Worksheet, wasProt As Boolean
    Set ws = ThisWorkbook.Worksheets(SH_IN)
    wasProt = ws.ProtectContents
    If wasProt Then ws.Unprotect
    On Error Resume Next
    ws.Shapes("btnWrite").Delete
    ws.Shapes("btnClear").Delete
    On Error GoTo 0
    MakeBtn ws, ws.Range("F16:F17"), "btnWrite", "WriteTrade"
    MakeBtn ws, ws.Range("G16:G17"), "btnClear", "ClearInput"
    If wasProt Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
    MsgBox "Кнопки ЗАПИСАТЬ и ОЧИСТИТЬ подключены.", vbInformation
End Sub

Private Sub MakeBtn(ByVal ws As Worksheet, ByVal rng As Range, ByVal nm As String, ByVal mac As String)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    s.Name = nm
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.Fill.Transparency = 0.99
    s.Line.Visible = msoFalse
    s.OnAction = "'" & ThisWorkbook.Name & "'!" & mac
End Sub
