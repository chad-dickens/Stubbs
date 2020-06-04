VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PeriodSelectionForm 
   Caption         =   "Period Selection"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "PeriodSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PeriodSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExitButtonCommand_Click()

ExitButton = True
Unload Me

End Sub

Private Sub MonthListBox_Change()

With PeriodSelectionForm
    If .MonthListBox.ListIndex >= 0 And .YearListBox.ListIndex >= 0 Then
        .SelectButton.Enabled = True
    End If
End With

End Sub

Private Sub SelectButton_Click()

With PeriodSelectionForm
    bSelectedMonth = .MonthListBox.ListIndex + 1
    lSelectedYear = .YearListBox.Value
End With
 
Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim r As Byte
Dim lYear As Long

'Clearing the MonthListBox of Previous Values
PeriodSelectionForm.MonthListBox.Clear

'Adding all 12 months to listbox
For r = 1 To 12
    PeriodSelectionForm.MonthListBox.AddItem MonthName(r)
Next r

'Adding 10 years to years list box
lYear = Year(Date) - 5

PeriodSelectionForm.YearListBox.Clear
For r = 1 To 10
    PeriodSelectionForm.YearListBox.AddItem lYear
    lYear = lYear + 1
Next r

PeriodSelectionForm.SelectButton.Enabled = False

End Sub

Private Sub YearListBox_Change()

    With PeriodSelectionForm
        If .MonthListBox.ListIndex >= 0 And .YearListBox.ListIndex >= 0 Then
            .SelectButton.Enabled = True
        End If
    End With
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then ExitButton = True

End Sub
