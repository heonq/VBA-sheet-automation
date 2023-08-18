Attribute VB_Name = "handleSheetsModule"
Option Explicit

Sub handleSheets()
Dim file As Workbook
Dim option As Variant
Dim i As Integer, i2 As Variant, i3 As Integer, i4 As Integer
Dim r As Integer, c As Integer
Dim mainWb As Workbook
Dim nagil As String, eastindigo As String
Dim wb As Workbook
Dim mall As String

nagil = ThisWorkbook.Path & "\(주)나길 업로드 양식.xlsx"
eastindigo = ThisWorkbook.Path & "\이스트인디고 업로드 양식.xlsx"
Set mainWb = Workbooks("배송시트 자동화.xlsm")


With Application
.DisplayAlerts = False
.Calculation = xlCalculationManual
.enableEvents = False
.ScreenUpdating = False
End With



For Each wb In Workbooks
wb.Activate
Set file = ActiveWorkbook
Set sheet = ActiveWorkbook.ActiveSheet


Call selectMall(mall,option)

If mall = "29cm" Then
With file.ActiveSheet.Range("AI:AI")
.Replace "CRATERS", "craters", xlPart
.Replace "_서오릉", "", xlPart
.Replace " JEWELRY", "", xlPart
End With

End If

If mall = "루앱" Then
wb.ActiveSheet.Range("T:T").Replace "(*)", "", xlPart
wb.ActiveSheet.Range("G:G").Replace "(", "", xlPart
wb.ActiveSheet.Range("G:G").Replace ")", "", xlPart
wb.ActiveSheet.Rows(1).Resize(7).Delete
End If


While wb.Worksheets.Count < 2
wb.Worksheets.Add after:=Worksheets(1)
Wend



Dim i5 As Integer

i5 = mainWb.Worksheets(1).Cells(1, 1).CurrentRegion.Rows.Count

i3 = 0
r = wb.Worksheets(1).UsedRange.Rows.Count
c = wb.Worksheets(1).UsedRange.Columns.Count


'파일 안에서 조건에 맞는 열을 mainWb로 복사 붙여넣는다


For Each i2 In option

wb.Worksheets(1).UsedRange.Find(i2, lookat:=xlWhole).Offset(1, 0).Resize(r - 1).Copy
wb.Worksheets(2).Range("A1").Offset(0, i3).Resize(r - 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

i3 = i3 + 1
Next i2

With wb.Worksheets(2)
If mall = "스스" Then
Dim i7 As Integer
For i7 = 1 To r
If Right(.Cells(i7, 2).Value, 2) = " L" Then
.Cells(i7, 3).Value = "L"
.Cells(i7, 2).Replace " L", "", xlPart
End If
If Right(.Cells(i7, 2).Value, 2) = " M" Then
.Cells(i7, 3).Value = "M"
.Cells(i7, 2).Replace " M", "", xlPart
End If
If Right(.Cells(i7, 2).Value, 2) = "XL" Then
.Cells(i7, 3).Value = "XL"
.Cells(i7, 2).Replace " XL", "", xlPart

End If
Next i7
End If
End With

'If mall = "무신사" Then
'파일.Worksheets(2).Cells(1, 12).Resize(r - 1).Value = 파일.Worksheets(1).UsedRange.Find("업체", lookat:=xlWhole).Offset(1).Resize(r - 1).Value
'End If

With wb.Worksheets(2)
.Columns(9).Insert shift:=xlToRight
.Columns(12).Insert shift:=xlToRight
.Columns(14).Insert shift:=xlToRight
.Cells(1, 9).Resize(r - 1).Value = mall
.Cells(1, 14).Resize(r - 1).NumberFormat = "yyyy-mm-dd"
.Cells(1, 14).Resize(r - 1).Value = Date


For i2 = 1 To r - 1
'주문번호에 주문날짜가 나와있지 않은 경우
If Not mall = "w컨셉" And Not mall = "gvg" And Not mall = "아몬즈" And Not mall = "루앱" And Not mall = "무신사" Then
.Cells(i2, 13).Value = 주문날짜(.Cells(i2, 1).Value, .Cells(i2, 9).Value)
.Cells(i2, 13).NumberFormat = "yyyy-mm-dd"
Else
.Cells(i2, 13).Value = Left(.Cells(i2, 13), 10)
End If
Next i2

For i2 = 1 To r - 1
If Not mall = "무신사" And Not mall = "29cm" And Not mall = "공홈" And Not mall = "스스" Then
.Cells(i2, 15).Value = "eastindigo"
End If
If mall = "스스" Then
.Cells(i2, 15).Value = "craters"
End If
Next i2

.UsedRange.Copy
End With
mainWb.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats


'----------------------시트 정리해서 옮기기 까지 완료--------------------------------
wb.Close


skip:

Next

'배송시트 내용중 필요없는 내용 삭제

Call replaceText

'나길 업로드,이스트인디고 업로드 파일 열기
With Application.Workbooks
.Open (nagil): .Open (eastindigo)
End With

'취합한 내용을 배송시트에 필요한 부분만 나길 파일과 판매 데이터 파일에 붙여넣는다

Dim r2 As Integer, r3 As Integer, r4 As Integer

i4 = mainWb.Worksheets(1).UsedRange.Rows.Count
c = mainWb.Worksheets(1).UsedRange.Columns.Count
r3 = 2: r4 = 2

For r2 = 2 To i4

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

Workbooks("(주)나길 업로드 양식.xlsx").Worksheets(1).Activate
Call 옵션텍스트1

With Application
.DisplayAlerts = True
.Calculation = xlCalculationAutomatic
.enableEvents = True
.ScreenUpdating = True
End With

End Sub