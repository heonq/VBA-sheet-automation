Attribute VB_Name = "handleSheetsModule"
Option Explicit

Sub 배송판매취합()
Dim 파일 As Workbook, 시트 As Worksheet, 셀 As Range
Dim 조건 As Variant
Dim i As Integer, i2 As Variant, i3 As Integer, i4 As Integer
Dim r As Integer, c As Integer
Dim 매크로파일 As Workbook
Dim 나길 As String, 판매데이터 As String, 이스트인디고 As String
Dim 워크북 As Workbook
Dim 입점사 As String
Dim 나길파일 As Workbook, 이스턴캐주얼파일 As Workbook

나길 = ThisWorkbook.Path & "\(주)나길 업로드 양식.xlsx"
이스트인디고 = ThisWorkbook.Path & "\이스트인디고 업로드 양식.xlsx"
Set 매크로파일 = ThisWorkbook


With Application
.DisplayAlerts = False
.Calculation = xlCalculationManual
.enableEvents = False
.ScreenUpdating = False
End With



For Each 워크북 In Workbooks
워크북.Activate
Set 파일 = ActiveWorkbook
Set 시트 = ActiveWorkbook.ActiveSheet


Call selectMall(입점사,조건)

If 입점사 = "29cm" Then
With 파일.ActiveSheet.Range("AI:AI")
.Replace "CRATERS", "craters", xlPart
.Replace "_서오릉", "", xlPart
.Replace " JEWELRY", "", xlPart
End With

End If

If 입점사 = "루앱" Then
파일.ActiveSheet.Range("T:T").Replace "(*)", "", xlPart
파일.ActiveSheet.Range("G:G").Replace "(", "", xlPart
파일.ActiveSheet.Range("G:G").Replace ")", "", xlPart
파일.ActiveSheet.Rows(1).Resize(7).Delete
End If


While 파일.Worksheets.Count < 2
파일.Worksheets.Add after:=Worksheets(1)
Wend



Dim i5 As Integer

i5 = 매크로파일.Worksheets(1).Cells(1, 1).CurrentRegion.Rows.Count

i3 = 0
r = 파일.Worksheets(1).UsedRange.Rows.Count
c = 파일.Worksheets(1).UsedRange.Columns.Count


'파일 안에서 조건에 맞는 열을 매크로파일로 복사 붙여넣는다


For Each i2 In 조건

파일.Worksheets(1).UsedRange.Find(i2, lookat:=xlWhole).Offset(1, 0).Resize(r - 1).Copy
파일.Worksheets(2).Range("A1").Offset(0, i3).Resize(r - 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

i3 = i3 + 1
Next i2

With 파일.Worksheets(2)
If 입점사 = "스스" Then
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

'If 입점사 = "무신사" Then
'파일.Worksheets(2).Cells(1, 12).Resize(r - 1).Value = 파일.Worksheets(1).UsedRange.Find("업체", lookat:=xlWhole).Offset(1).Resize(r - 1).Value
'End If

With 파일.Worksheets(2)
.Columns(9).Insert shift:=xlToRight
.Columns(12).Insert shift:=xlToRight
.Columns(14).Insert shift:=xlToRight
.Cells(1, 9).Resize(r - 1).Value = 입점사
.Cells(1, 14).Resize(r - 1).NumberFormat = "yyyy-mm-dd"
.Cells(1, 14).Resize(r - 1).Value = Date


For i2 = 1 To r - 1
'주문번호에 주문날짜가 나와있지 않은 경우
If Not 입점사 = "w컨셉" And Not 입점사 = "gvg" And Not 입점사 = "아몬즈" And Not 입점사 = "루앱" And Not 입점사 = "무신사" Then
.Cells(i2, 13).Value = 주문날짜(.Cells(i2, 1).Value, .Cells(i2, 9).Value)
.Cells(i2, 13).NumberFormat = "yyyy-mm-dd"
Else
.Cells(i2, 13).Value = Left(.Cells(i2, 13), 10)
End If
Next i2

For i2 = 1 To r - 1
If Not 입점사 = "무신사" And Not 입점사 = "29cm" And Not 입점사 = "공홈" And Not 입점사 = "스스" Then
.Cells(i2, 15).Value = "eastindigo"
End If
If 입점사 = "스스" Then
.Cells(i2, 15).Value = "craters"
End If
Next i2

.UsedRange.Copy
End With
매크로파일.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats


'----------------------시트 정리해서 옮기기 까지 완료--------------------------------
파일.Close


skip:

Next

'배송시트 내용중 필요없는 내용 삭제

Call replaceText

'나길 업로드,이스트인디고 업로드 파일 열기
With Application.Workbooks
.Open (나길): .Open (이스트인디고)
End With

'취합한 내용을 배송시트에 필요한 부분만 나길 파일과 판매 데이터 파일에 붙여넣는다

Dim r2 As Integer, r3 As Integer, r4 As Integer

i4 = 매크로파일.Worksheets(1).UsedRange.Rows.Count
c = 매크로파일.Worksheets(1).UsedRange.Columns.Count
r3 = 2: r4 = 2

For r2 = 2 To i4

매크로파일.Worksheets(1).Cells(r2, 1).Resize(, 10).Copy

If 매크로파일.Worksheets(1).Cells(r2, 15).Value = "eastindigo" Then
Workbooks("이스트인디고 업로드 형식.xlsx").Worksheets(1).Cells(r3, 2).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
r3 = r3 + 1
Else
With Workbooks("(주)나길 업로딩 양식.xlsx").Worksheets(1)
.Cells(r4, 4).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
.Range("A2:C2").Copy Destination:=.Cells(r4, 1).Resize(, 3)


End With
r4 = r4 + 1
End If
Next r2

Workbooks("(주)나길 업로딩 양식.xlsx").Worksheets(1).Activate
Call 옵션텍스트1

With Application
.DisplayAlerts = True
.Calculation = xlCalculationAutomatic
.enableEvents = True
.ScreenUpdating = True
End With

End Sub