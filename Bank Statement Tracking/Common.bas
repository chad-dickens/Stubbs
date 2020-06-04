Attribute VB_Name = "Common"
Option Explicit

Public Sub OptimizeOn()
'Turns off non-essential Excel features to make code run faster
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .StatusBar = "Working..."
End With

End Sub

Public Sub OptimizeOff()
'Turns these essential features back on
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .StatusBar = ""
End With

End Sub

Public Function CustomFormat(InputValue As Currency) As String
    'Simple function to return a number in currency form
    CustomFormat = "$" & Format(InputValue, "#,###.##")
    If (Right(CustomFormat, 1) = ".") Then
        CustomFormat = Left(CustomFormat, Len(CustomFormat) - 1)
    End If
End Function
