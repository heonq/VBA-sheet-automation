Sub handleMalls(ByVal mall As String)

Select Case mall
Case "29cm": Call handle29cm
Case "루앱": Call handleLuaeb
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


