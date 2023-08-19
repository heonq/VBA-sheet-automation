Function orderDateInfo(orderNumber As Variant, mall As String) As Variant
Dim dateInfo As Variant
Select Case mall
Case "29cm": dateInfo = Mid(orderNumber, 4, 8)
Case "w컨셉": dateInfo = Left(orderNumber, 10)
Case "하고": dateInfo = "20" & Left(orderNumber, 6)

Case Else: dateInfo = Left(orderNumber, 8)


End Select

Select Case mall
Case "w컨셉": orderDateInfo = dateInfo


Case Else
orderDateInfo = Left(dateInfo, 4) & "/" & Mid(dateInfo, 5, 2) & "/" & Mid(dateInfo, 7, 2)
End Select

End Function

