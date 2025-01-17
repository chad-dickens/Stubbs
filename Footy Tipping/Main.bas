Attribute VB_Name = "Main"
Option Explicit
Option Base 1

'Public variables to be accessed by all modules
Public Entrant As String
Public ColumnNum As Integer
Public RowNum As Integer
Public RoundNum As Integer

'Password for all sheets
Public Const Password As String = "exonthebeach"

Sub enter_your_tips()
'Loads the input box for the users to choose the name of the person they are entering data for
'If they select a name, then it will load the userform for inputting tips

    Entrant = ""
    ColumnNum = 2
    RowNum = 5
    
    'Check if this person already exists and make sure they don't select nothing
    Entrant = InputBox("Who are you inputting tips for?")
    
    If Entrant = "" Then Exit Sub
    
    'Locate the relevant round, row, and column of the selected person
    Call find_row_col
    
    If RowNum = 203 Then
        MsgBox "You have already inputted all 23 rounds for " & Entrant
    Else
        Enter_Tips.Show
    End If
    
End Sub

Public Sub find_row_col(Optional ByVal want_col As Boolean = True)
'Sets the value of the row, round number, and column for the person selected in the data sheet
'The row number is the next game they haven't selected
'The reason for the optional parameter is that there is no need to find the column a second time
'if this global variable has already been set.
'The column global variable is set when the user chooses the name of the person and doesn't need
'to be found again when the Userform reloads.
'The round number however changes each time the user form is loaded and needs to be found again.

Call Optimize_On

'Find the row, round number, and column number for the person
    With data_sht
        
        'Unlock sheet
        .Unprotect Password = Password
        
        'The optional parameter being used
        If want_col Then
            Do While True
                If .Cells(1, ColumnNum).Value = Entrant Or IsEmpty(.Cells(1, ColumnNum).Value) Then
                    .Cells(1, ColumnNum).Value = Entrant
                    Exit Do
                End If
                ColumnNum = ColumnNum + 1
            Loop
        End If
        
        Do While True
            If IsEmpty(.Cells(RowNum, ColumnNum).Value) Then
                RoundNum = .Cells(RowNum, 1).Value
                Exit Do
            End If
            RowNum = RowNum + 1
        Loop
        
        'Lock sheet
        .Protect Password = Password
        
    End With
    
Call Optimize_Off

End Sub

Public Sub update_graph()
'Updates the top 10 graph in the main sheet.
'It does this by updating the transpose sheet - the data driving the graph.
'The transpose sheet is just the data sheet's top 2 rows but transposed vertically and sorted
'descending. This sheet is hidden.

    Dim i As Integer
    i = 2
    
    With transpose_sht
        
        .Cells.ClearContents
        .Range("A1").Value = "Person"
        .Range("B1").Value = "Score"
        
        While Not IsEmpty(data_sht.Cells(1, i))
                .Cells(i, 1).Value = data_sht.Cells(1, i).Value
                .Cells(i, 2).Value = data_sht.Cells(2, i).Value
                i = i + 1
        Wend
    
        'Sorting the table so only the top 10 show
        .Range("A1:B" & (i - 1)).Sort Key1:=.Range("B1"), Order1:=xlDescending
        
    End With

End Sub

Public Function rounds_completed() As Variant
'Determines the number of rounds that have been fully completed ie, scores for each game are in
'the fixture sheet.
    
    Dim i As Integer
    i = 2

    With fixture_sht

        While .Cells(i, 7).Value <> -1 And Not IsEmpty(.Cells(i, 1)) And Not IsEmpty(.Cells(i, 7))
            i = i + 1
        Wend
        
        If IsEmpty(.Cells(i, 1)) Then
            rounds_completed = "Season Over"
        Else
            rounds_completed = .Cells(i, 1).Value - 1
        End If

    End With

End Function

Public Function decide_winner(game_score As Range, team1 As Range, team2 As Range) As Variant
'Function for deciding the winner in the fixture sheet - to be used as an actual excel function
    
    On Error GoTo ErrHandler
    
    Dim score1 As Integer
    Dim score2 As Integer

    score1 = Int(Left(game_score.Text, InStr(1, game_score.Text, " - ") - 1))
    score2 = Int(Mid(game_score.Text, InStr(1, game_score.Text, " - ") + 3))
        
    If score1 > score2 Then
        decide_winner = team1.Text
    ElseIf score1 < score2 Then
        decide_winner = team2.Text
    Else
        decide_winner = "Draw"
    End If
    
    Exit Function

ErrHandler:
decide_winner = -1

End Function

Sub update_fixture_results()
'Accesses the internet and gets a HTML file of the most recent AFL results.
'Uses this to update the results for each game and the scores for each person.
    
    On Error GoTo ErrHandler
    
    Call Optimize_On
    
    'Unlocking the relevant sheets
    fixture_sht.Unprotect Password = Password
    main_sht.Unprotect Password = Password

    Dim Tr As Object
    Dim iRow As Integer
    Dim HTML_Content As Object
    Dim error_message As String
    Dim arr As Variant
    
    'Create HTMLFile Object
    Set HTML_Content = CreateObject("htmlfile")

    'Get the WebPage Content to HTMLFile Object
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", "https://fixturedownload.com/results/afl-2021", False
        
        'If the internet isn't working, this is where the problem will be found
        On Error Resume Next
        .send
        On Error GoTo ErrHandler
        
        'If this is true then we have a problem
        If .Status <> 200 Then
            MsgBox "There was a problem connecting to the internet at https://fixturedownload.com/results/afl-2021. " & _
                   "Please ensure you have a valid connection and that this website is still legitimate."
            GoTo ErrHandler
        End If
        
        HTML_Content.Body.Innerhtml = .responseText
    End With
    
    'Check if a table exists at all
    error_message = "There is a problem with https://fixturedownload.com/results/afl-2021. " & _
                    "Please ensure this website is correct."
    
    If HTML_Content.getElementsByTagName("table")(0) Is Nothing Then
        MsgBox error_message
        GoTo ErrHandler
    End If
    
    'Put Current Fixture Sheet into Memory
    arr = fixture_sht.Range("A1").CurrentRegion
    
    'Loop through HTML table
    With HTML_Content.getElementsByTagName("table")(0)
        
        'Checking that the table has the right dimensions
        If .Rows.Length < 10 Or .Rows(1).Cells.Length <> 6 Then
            MsgBox error_message
            GoTo ErrHandler
        End If
        
        'Iterating through the html table
        For Each Tr In .Rows
            
            'At the same time, for each iteration of the html table, there is a full loop
            'through arr, to match up teams playing and rounds. If there is a match, the
            'corresponding result is input into the fixture sheet. The reason for this
            'is to ensure that if in future rounds the order of teams playing is changed
            'in the fixture, it won't ruin the workbook. This will also avoid problems
            'if a round isn't played at all.
            
            For iRow = 2 To UBound(arr)
                
                If Tr.Cells(3).innerText = arr(iRow, 4) And _
                Tr.Cells(4).innerText = arr(iRow, 5) And _
                Tr.Cells(0).innerText = CStr(arr(iRow, 1)) Then
                    fixture_sht.Cells(iRow, 6).Value = "'" & Tr.Cells(5).innerText
                    Exit For
                End If
                
            Next iRow
            
        Next Tr
    End With
    
    'Update all formulas - in particular the formulas for who won on the fixture sheet, the
    'custom function on the main sheet, as well as each person's score on the data sheet.
    Application.CalculateFull
    
    'Setting the value of when the last update was run so users are aware.
    main_sht.Range("G6").Value = Now
    
    'Update the new scores for all participants and update the top 10 leaderboard graph.
    Call update_graph
    
    'Locking the relevant sheets
    fixture_sht.Protect Password = Password
    main_sht.Protect Password = Password
    
    'This is to ensure no unexpected behaviour occurs
    main_sht.Select
    
    Call Optimize_Off
    
    Exit Sub
    
ErrHandler:
'For handling errors

If Err.Number <> 0 Then
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description
End If

'Locking the relevant sheets to make sure they aren't left open
fixture_sht.Protect Password = Password
main_sht.Protect Password = Password

main_sht.Select

Call Optimize_Off

End Sub
