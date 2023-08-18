Sub turnOffAlert()

With Application
.DisplayAlerts = False
.Calculation = xlCalculationManual
.enableEvents = False
.ScreenUpdating = False
End With

End Sub

Sub turnOnAlert()

With Application
.DisplayAlerts = True
.Calculation = xlCalculationAutomatic
.enableEvents = True
.ScreenUpdating = True
End With

End Sub
