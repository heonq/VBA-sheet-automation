Sub copyToUploadForm(ByVal mainWb As Workbook)

Dim r2 As Integer, r3 As Integer, r4 As Integer

i1 = mainWb.Worksheets(1).UsedRange.Rows.Count
r3 = 2: r4 = 2

For r2 = 2 To i1

mainWb.Worksheets(1).Cells(r2, 1).Resize(, 10).Copy

If mainWb.Worksheets(1).Cells(r2, 15).Value = "eastindigo" Then
Workbooks("이스트인디고 업로드 형식.xlsx").Worksheets(1).Cells(r3, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
r3 = r3 + 1
Else
With Workbooks("(주)나길 업로드 양식.xlsx").Worksheets(1)
.Cells(r4, 4).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
.Range("A2:C2").Copy Destination:=.Cells(r4, 1).Resize(, 3)


End With
r4 = r4 + 1
End If
Next r2

End Sub
