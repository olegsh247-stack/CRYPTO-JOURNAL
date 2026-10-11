Option Explicit
' Код книги (модуль ThisWorkbook): автодобавление новых пар в листы ИТОГО.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    PairsOnChange Sh, Target
End Sub
