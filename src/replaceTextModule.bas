Attribute VB_Name = "replaceTextModule"
Option Explicit

Sub replaceText()

Dim before As Variant
Dim after As Variant
dim i1 as integer , i2 as integer

before = array(10000)
after = array(10000)

'변경할 단어를 입력해주세요.
'before를 after로 변경합니다.

before(0) = ""
after(0) = ""

i1 = Ubound(before)-Lbound(before)

With activeWorkbook.Worksheets(1).UsedRange
For i2 = 0 To i1
if before(i2) is Nothing Then goTo Skip
.Replace before(i2), after(i2), xlPart
Next i2

Skip:

.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), Header:=xlNo
End With


End Sub