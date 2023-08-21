Attribute VB_Name = "handleSheetsModule"
Option Explicit

Sub handleSheets()
Dim wb As Workbook
Dim mallOption As Variant
Dim i1 As Integer
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
Call copyOrderInfo(mallOption, wb)
call addDateAndMallInfo(mall,wb)

wb.Worksheets(2).UsedRange.Copy
mainWb.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

wb.Close

skip:

Next

Call replaceText

With Application.Workbooks
.Open (nagil): .Open (eastindigo)
End With

Call copyToUploadForm(mainWb)
Call mergeOptionAndProductName
Call turnOnAlert

End Sub