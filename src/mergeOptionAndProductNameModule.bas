option Explicit

Sub mergeOptionAndProductName()

Dim r As Integer, i As Integer
Dim c As Integer
dim productColumn as integer
dim optionColumn as integer
dim quantityColumn as integer

productColumn = 5
optionColumn = 6
quantityColumn = 13

Workbooks("(주)나길 업로드 양식.xlsx").Worksheets(1).Activate
r = ActiveSheet.UsedRange.Rows.Count

For i = 2 To r
Cells(i, productColumn).Value = WorksheetFunction.TextJoin(" ", True, Cells(i, productColumn), Cells(i, optionColumn))
Cells(i, optionColumn).Value = Cells(i, quantityColumn).Value
Cells(i, quantityColumn).Value = ""

Next i

End Sub