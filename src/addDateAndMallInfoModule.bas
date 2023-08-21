Sub addDateAndMallInfo(mall As String, wb as Workbook)

Dim rforFile As Integer, i1 As Integer
rforFile = wb.Worksheets(2).UsedRange.Rows.Count

With wb.Worksheets(2)
.Columns(9).Insert shift:=xlToRight
.Columns(12).Insert shift:=xlToRight
.Columns(14).Insert shift:=xlToRight
.Cells(1, 9).Resize(rforFile).Value = mall
.Cells(1, 14).Resize(rforFile).NumberFormat = "yyyy-mm-dd"
.Cells(1, 14).Resize(rforFile).Value = Date


For i1 = 1 To rforFile
'주문번호에 주문날짜가 나와있지 않은 경우
If Not mall = "w컨셉" And Not mall = "아몬즈" And Not mall = "루앱" And Not mall = "무신사" Then
.Cells(i1, 13).Value = orderDateInfo(.Cells(i1, 1).Value, .Cells(i1, 9).Value)
.Cells(i1, 13).NumberFormat = "yyyy-mm-dd"
'주문번호에 주문날짜가 나와있는 경우
Else
.Cells(i1,13).Value = Left(.Cells(i1, 13), 10)
End If
Next i1

For i1 = 1 To rforFile
If Not mall = "무신사" And Not mall = "29cm" And Not mall = "공홈" And Not mall = "스스" Then
.Cells(i1, 15).Value = "eastindigo"
End If
If mall = "스스" Then
.Cells(i1, 15).Value = "craters"
End If
Next i1
End With

End Sub