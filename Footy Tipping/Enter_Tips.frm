VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Enter_Tips 
   Caption         =   "Enter Tips"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460.001
   OleObjectBlob   =   "Enter_Tips.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Enter_Tips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'When initializing the user form, this code looks in the fixture sheet to find all games for that
'week and loads the userform accordingly.
    
    Dim start_row As Integer
    'Start row needs to be 2 because this is where the fixtures start
    start_row = 2
    'Setting the person's name in the userform
    person_name_label.Caption = Entrant
    
    'We already know the column number from running this previously, so that is why the
    'parameter is false. This means it will only find the row and round number.
    Call find_row_col(False)
    
    'Setting the round number in the userform
    round_label.Caption = "Round " & CStr(round_num)
    
    'Changing the button caption if it's the last round
    If round_num = 23 Then next_button.Caption = "Complete"
    
    'Find the teams playing in this round
    Do While True
        With fixture_sht
            
            'Making sure we don't iterate longer than we have to
            If .Cells(start_row, 1).Value > round_num Or IsEmpty(.Cells(start_row, 1).Value) Then
                Exit Do
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_1_box.AddItem .Cells(start_row, 4).Text
                game_1_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_2_box.AddItem .Cells(start_row, 4).Text
                game_2_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_3_box.AddItem .Cells(start_row, 4).Text
                game_3_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_4_box.AddItem .Cells(start_row, 4).Text
                game_4_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_5_box.AddItem .Cells(start_row, 4).Text
                game_5_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_6_box.AddItem .Cells(start_row, 4).Text
                game_6_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_7_box.AddItem .Cells(start_row, 4).Text
                game_7_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_8_box.AddItem .Cells(start_row, 4).Text
                game_8_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            If .Cells(start_row, 1).Value = round_num Then
                game_9_box.AddItem .Cells(start_row, 4).Text
                game_9_box.AddItem .Cells(start_row, 5).Text
                start_row = start_row + 1
            End If
            
            start_row = start_row + 1
            
        End With
        
    Loop
    
    'Making sure if there are only 6 games, that the last 3 drop downs aren't selectable.
    If game_7_box.ListCount = 0 Then
        game_7_box.Enabled = False
        game_7_label.Caption = "No Game"
    End If
    
    If game_8_box.ListCount = 0 Then
        game_8_box.Enabled = False
        game_8_label.Caption = "No Game"
    End If
    
    If game_9_box.ListCount = 0 Then
        game_9_box.Enabled = False
        game_9_label.Caption = "No Game"
    End If
    
End Sub

Sub game_box_change()
'Looks through all drop down boxes to check that every one has a value in it.
'If this is the case, then the "next" button will be enabled.
    
    Dim cmb As Control
    
    For Each cmb In Me.Controls
        
        If TypeName(cmb) = "ComboBox" Then
            If cmb.Enabled = True And cmb.Value = vbNullString Then
                Exit Sub
            End If
        End If
        
    Next
    
    Enter_Tips.next_button.Enabled = True
    
End Sub

Private Sub game_1_box_Change()
    Call game_box_change
End Sub

Private Sub game_2_box_Change()
    Call game_box_change
End Sub

Private Sub game_3_box_Change()
    Call game_box_change
End Sub

Private Sub game_4_box_Change()
    Call game_box_change
End Sub

Private Sub game_5_box_Change()
    Call game_box_change
End Sub

Private Sub game_6_box_Change()
    Call game_box_change
End Sub

Private Sub game_7_box_Change()
    Call game_box_change
End Sub

Private Sub game_8_box_Change()
    Call game_box_change
End Sub

Private Sub game_9_box_Change()
    Call game_box_change
End Sub

Private Sub next_button_Click()
'Pastes the tips for each round into the datasheet
    
    With data_sht
    
        'Unlocking data sheet
        .Unprotect Password = Password
        
        .Cells(RowNum, ColumnNum).Value = game_1_box.Value
        RowNum = RowNum + 1
        
        .Cells(RowNum, ColumnNum).Value = game_2_box.Value
        RowNum = RowNum + 1
        
        .Cells(RowNum, ColumnNum).Value = game_3_box.Value
        RowNum = RowNum + 1
        
        .Cells(RowNum, ColumnNum).Value = game_4_box.Value
        RowNum = RowNum + 1
        
        .Cells(RowNum, ColumnNum).Value = game_5_box.Value
        RowNum = RowNum + 1
        
        .Cells(RowNum, ColumnNum).Value = game_6_box.Value
        RowNum = RowNum + 1
        
        'The following 3 sections are necessary because in some rounds there are only 6 matches
        'so this is to ensure that values are only taken where the dropdown is enabled.
        If game_7_box.Enabled = True Then
            .Cells(RowNum, ColumnNum).Value = game_7_box.Value
            RowNum = RowNum + 1
        End If
        
        If game_8_box.Enabled = True Then
            .Cells(RowNum, ColumnNum).Value = game_8_box.Value
            RowNum = RowNum + 1
        End If
        
        If game_9_box.Enabled = True Then
            .Cells(RowNum, ColumnNum).Value = game_9_box.Value
            RowNum = RowNum + 1
        End If
        
        'Locking data sheet again
        .Protect Password = Password
        
    End With
    
    'If the round number is 23 it means that there are no further rounds
    If round_num = 23 Then
        Unload Enter_Tips
    Else
        Unload Enter_Tips
        Enter_Tips.Show
    End If
    
    Call update_graph

End Sub

