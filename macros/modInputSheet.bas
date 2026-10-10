Option Explicit
' Код листа «Ввод» (модуль листа). Подсказка монет при вводе Актива 1 / Актива 2.

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim hit As Range, c As Range
    Set hit = Intersect(Target, Me.Range("C5:D5"))
    If hit Is Nothing Then Exit Sub
    On Error GoTo Fin
    Application.EnableEvents = False
    For Each c In hit.Cells
        CompleteCoin c
    Next c
Fin:
    Application.EnableEvents = True
End Sub
