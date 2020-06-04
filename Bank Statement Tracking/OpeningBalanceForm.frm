VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpeningBalanceForm 
   Caption         =   "Opening Balance"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "OpeningBalanceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OpeningBalanceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExitCommandButton_Click()

ExitButton = True
Unload Me

End Sub

Private Sub SelectButton_Click()

On Error GoTo ErrHandler

Dim vFileD As Office.FileDialog
Set vFileD = Application.FileDialog(msoFileDialogFilePicker)

With vFileD
    .AllowMultiSelect = False
    .Title = "Please Select The Previous Month's File"
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx, *.xls, *.xlsm"
    .Show
    
    'Setting the path of our workbook
    sPreviousMonthFile = .SelectedItems(1)
    
End With

'If an error has not occurred then we know that a valid file has been selected
Unload Me

ErrHandler:

End Sub

Private Sub TypedBalance_Click()

'Forcing the user to input a number
OpeningBalance = Application.InputBox("Please type in your opening balance:", "Opening Balance", , , , , , 1)

'If the user has inputted a valid value
If OpeningBalance <> 0 Then Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then ExitButton = True

End Sub
