option Explicit

Sub 옵션텍스트1()

Dim r As Integer, i As Integer
Dim c As Integer

r = ActiveSheet.UsedRange.Rows.Count

For i = 2 To r
Cells(i, 5).Value = WorksheetFunction.TextJoin(" ", True, Cells(i, 5), Cells(i, 6))
Cells(i, 6).Value = Cells(i, 13).Value
Cells(i, 13).Value = ""



Next i

End Sub