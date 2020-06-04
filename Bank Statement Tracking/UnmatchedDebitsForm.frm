VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UnmatchedDebitsForm 
   Caption         =   "Unmatched Debits"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "UnmatchedDebitsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UnmatchedDebitsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub BankFeesButton_Click()

Call UpdateFees(12, UnmatchedDebitsForm.ListBox1.Value)

End Sub

Private Sub DishonourButton_Click()

Call UpdateFees(14, UnmatchedDebitsForm.ListBox1.Value)

End Sub

Private Sub ListBox1_Click()

With UnmatchedDebitsForm

    If .ListBox1.Value <> "Null" Then
        
        .BankFeesButton.Enabled = True
        .MCButton.Enabled = True
        .DishonourButton.Enabled = True
        .SecureFeesButton.Enabled = True
        
    End If
    
End With

End Sub

Private Sub MCButton_Click()

Call UpdateFees(15, UnmatchedDebitsForm.ListBox1.Value)

End Sub

Private Sub SecureFeesButton_Click()

Call UpdateFees(13, UnmatchedDebitsForm.ListBox1.Value)

End Sub

Private Sub UserForm_Initialize()

Call UpdateListBox

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then ExitButton = True

End Sub

Private Sub UpdateListBox()

Dim r As Long

With UnmatchedDebitsForm

.ListBox1.Clear

    For r = 2 To UBound(aBankA)
    
        If aBankA(r, 3) <> "MATCHED" Then
                .ListBox1.AddItem aBankA(r, 2)
                .ListBox1.List(ListBox1.ListCount - 1, 1) = aBankA(r, 4)
        End If
    
    Next r

    If .ListBox1.ListCount = 0 Then
        Unload Me
        Exit Sub
    End If

        .BankFeesButton.Enabled = False
        .MCButton.Enabled = False
        .DishonourButton.Enabled = False
        .SecureFeesButton.Enabled = False
    
End With

End Sub

Private Sub UpdateFees(bColNum As Byte, sDescription As String)

Dim r As Long

    'Start off by finding the array row value of the item
    For r = 2 To UBound(aBankA)
        If aBankA(r, 2) = sDescription Then
            Exit For
        End If
    Next r
    
    'Then add it in, using the day of the month it is
    With vBankA.Cells(Day(aBankA(r, 1)) + 4, bColNum)
        .Value = .Value - aBankA(r, 4)
    End With
    
    'Make it matched in the array
    aBankA(r, 3) = "MATCHED"
    
    'Update the list again
    Call UpdateListBox
    
End Sub

