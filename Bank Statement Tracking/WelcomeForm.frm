VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WelcomeForm 
   Caption         =   "Bank Statement Tracking"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "WelcomeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ExitButtonCommand_Click()

ExitButton = True
Unload Me

End Sub

Private Sub InstructionsButton_Click()

InstructionsFrm.Show

End Sub

Private Sub ProceedButton_Click()

Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then ExitButton = True

End Sub
