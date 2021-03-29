Attribute VB_Name = "Common"
Option Explicit

Public Sub Optimize_On()
'Turns off non-essential Excel features to make code run faster
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .StatusBar = "Working..."
End With

End Sub

Public Sub Optimize_Off()
'Turns these essential features back on
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .StatusBar = ""
End With

End Sub
