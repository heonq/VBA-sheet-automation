Sub copyOrderInfo(ByVal mallOption As Variant, ByVal wb As Workbook)

Dim i1 As Integer, r As Integer
Dim currentOption As Variant

With wb.Worksheets
While .Count < 2
.Add after:=Worksheets(1)
Wend
End With

i1 = 0
r = wb.Worksheets(1).UsedRange.Rows.Count


'파일 안에서 조건에 맞는 열을 mainWb로 복사 붙여넣는다


For Each currentOption In mallOption

wb.Worksheets(1).UsedRange.Find(currentOption, lookat:=xlWhole).Offset(1, 0).Resize(r - 1).Copy
wb.Worksheets(2).Range("A1").Offset(0, i1).Resize(r - 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

i1 = i1 + 1
Next currentOption

End Sub
