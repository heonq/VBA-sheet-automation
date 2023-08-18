Attribute VB_Name = "handleSheetsModule"
Option Explicit

Sub handleSheets()
Dim wb As Workbook
Dim mallOption As Variant
Dim i1 As Integer
dim currentOption As Variant
Dim r As Integer
Dim mainWb As Workbook
Dim nagil As String, eastindigo As String
Dim mall As String

nagil = ThisWorkbook.Path & "\(주)나길 업로드 양식.xlsx"
eastindigo = ThisWorkbook.Path & "\이스트인디고 업로드 양식.xlsx"
Set mainWb = Workbooks("배송시트 자동화.xlsm")


Call turnOffAlert

For Each wb In Workbooks
wb.Activate

Call selectMall(mall,mallOption)
If mall = "X" Then GoTo skip

Call handleMalls(mall)

While wb.Worksheets.Count < 2
wb.Worksheets.Add after:=Worksheets(1)
Wend

i1 = 0
r = wb.Worksheets(1).UsedRange.Rows.Count

For Each currentOption In mallOption

wb.Worksheets(1).UsedRange.Find(currentOption, lookat:=xlWhole).Offset(1, 0).Resize(r - 1).Copy
wb.Worksheets(2).Range("A1").Offset(0, i1).Resize(r - 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

i1 = i1 + 1
Next currentOption

With wb.Worksheets(2)
.Columns(9).Insert shift:=xlToRight
.Columns(12).Insert shift:=xlToRight
.Columns(14).Insert shift:=xlToRight
.Cells(1, 9).Resize(r - 1).Value = mall
.Cells(1, 14).Resize(r - 1).NumberFormat = "yyyy-mm-dd"
.Cells(1, 14).Resize(r - 1).Value = Date


For i1 = 1 To r - 1
'주문번호에 주문날짜가 나와있지 않은 경우
If Not mall = "w컨셉" And Not mall = "gvg" And Not mall = "아몬즈" And Not mall = "루앱" And Not mall = "무신사" Then
.Cells(i1, 13).Value = 주문날짜(.Cells(i1, 1).Value, .Cells(i1, 9).Value)
.Cells(i1, 13).NumberFormat = "yyyy-mm-dd"
Else
.Cells(i1, 13).Value = Left(.Cells(i1, 13), 10)
End If
Next i1

For i1 = 1 To r - 1
If Not mall = "무신사" And Not mall = "29cm" And Not mall = "공홈" And Not mall = "스스" Then
.Cells(i1, 15).Value = "eastindigo"
End If
If mall = "스스" Then
.Cells(i1, 15).Value = "craters"
End If
Next i1

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

i1 = mainWb.Worksheets(1).UsedRange.Rows.Count
c = mainWb.Worksheets(1).UsedRange.Columns.Count
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

Workbooks("(주)나길 업로드 양식.xlsx").Worksheets(1).Activate
Call 옵션텍스트1

Call turnOnAlert

End Sub