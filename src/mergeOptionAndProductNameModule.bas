option Explicit

Sub mergeOptionAndProductName()

Dim r As Integer, i As Integer
Dim c As Integer

Workbooks("(주)나길 업로드 양식.xlsx").Worksheets(1).Activate
r = ActiveSheet.UsedRange.Rows.Count

For i = 2 To r
Cells(i, 5).Value = WorksheetFunction.TextJoin(" ", True, Cells(i, 5), Cells(i, 6))
Cells(i, 6).Value = Cells(i, 13).Value
Cells(i, 13).Value = ""



Next i

End Sub