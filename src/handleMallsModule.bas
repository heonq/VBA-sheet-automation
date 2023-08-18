Sub handleMalls(ByVal mall As String)

Select Case mall
Case "29cm": Call handle29cm
Case "루앱": Call handleLuaeb
Case "스스": Call handleSmartStore
End Select

End Sub

Sub handle29cm()

With ActiveSheet.Range("AI:AI")
.Replace "CRATERS", "craters", xlPart
.Replace "_서오릉", "", xlPart
.Replace " JEWELRY", "", xlPart
End With

End Sub

Sub handleLuaeb()
with ActiveSheet
.Range("T:T").Replace "(*)", "", xlPart
.Range("G:G").Replace "(", "", xlPart
.Range("G:G").Replace ")", "", xlPart
.Rows(1).Resize(7).Delete
End If

End Sub

Sub handleSmartStore()

Dim i1 As Integer
Dim r1 As Integer
Dim productNameColumn As Integer
Dim optionColumn As Integer
Dim optionInProductName As String

productNameColumn = ActiveSheet.UsedRange.Find("옵션관리코드", lookat:=xlWhole).Column
optionColumn = ActiveSheet.UsedRange.Find("옵션정보", lookat:=xlWhole).Column

With ActiveSheet
.Rows(1).Delete
r1 = .UsedRange.Rows.Count
For i1 = 1 To r1
optionInProductName = Right(.Cells(i1, productNameColumn).Value, 2)
Select Case optionInProductName
Case " L": .Cells(i1, optionColumn).Value = "L": .Cells(i1, productNameColumn).Replace " L", "", xlPart
Case " M": .Cells(i1, optionColumn).Value = "M": .Cells(i1, productNameColumn).Replace " M", "", xlPart
Case "XL": .Cells(i1, optionColumn).Value = "XL": .Cells(i1, productNameColumn).Replace " XL", "", xlPart
End Select
Next i1
End With

End Sub