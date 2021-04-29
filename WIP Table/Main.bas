Attribute VB_Name = "main"
Option Explicit
Option Base 1

'This is the main module of the WIP table creation process.
'Will create a WIP table to be used in a remuneration report.

Sub create_table()

'Will use the WIP listing table (Amanda or Remuneration Report (Chad) - directly out of APS) on
'the currently active sheet and create a WIP Summary table on a new sheet. This new sheet will
'be selected once the process has finished running.

On Error GoTo ErrHandler

'To make the process run faster
Call optimize_on

'Creating variables
Dim wb As Workbook
Dim ws As Worksheet
Dim myStr As String
Dim rowNum As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim dict As Object
Dim v As Variant
Dim v2 As Variant
Dim categories As Variant
Dim arr As Variant
Dim arr2() As Variant
Dim arr3() As Variant
Dim matchFound As Boolean

'Variables for columns of input table
'This is necessary incase if in future, the input table changes
Dim inputDateCol As Byte: inputDateCol = 2
Dim inputNameCol As Byte: inputNameCol = 3
Dim inputHoursCol As Byte: inputHoursCol = 5
Dim inputCatCol As Byte: inputCatCol = 6
Dim inputRateCol As Byte: inputRateCol = 13

Set wb = Application.ActiveWorkbook
Set ws = Application.ActiveSheet

'Doing some checks to make sure there is nothing wrong with the sheet
arr = Array("SortName", "WIP_Date", "Name", "Rate_Description", "Hours", "Std_Mtr", "Milestone", _
            "Narration", "Value", "Billed", "Write_Off", "Net_WIP", "Actual_Rate", "Standard_Rate")

'Looping through this list of the table headings to ensure that the current sheet has these table
'headings
For i = 1 To UBound(arr)
    
    If arr(i) <> ws.Cells(1, i).Value Then
        'The sheet that this person is trying to use is incorrect
        MsgBox "The current sheet does no satify the requirements for a WIP Listing table. " & _
               "You must use the 'Amanda' table, or the 'Remuneration Report (Chad)' table. " & _
               "Alternatively, you can create your own table, but it must contain the headings " & _
               "of 'SortName', 'WIP_Date', 'Name', 'Rate_Description', 'Hours', 'Std_Mtr', " & _
               "'Milestone', 'Narration', 'Value', 'Billed', 'Write_Off', 'Net_WIP', " & _
               "'Actual_Rate', and 'Standard_Rate', in that order."
        GoTo ErrHandler
    End If
Next i

'Putting the current sheet table into an array
arr = ws.Cells(1, 1).CurrentRegion

'Cleaning data
For i = 2 To UBound(arr)

    'Removing zz at the start of people's names
    If LCase(Left(arr(i, inputNameCol), 2)) = "zz" Then arr(i, inputNameCol) = Mid(arr(i, inputNameCol), 3)
    'Triming peoples names of non letter characters
    arr(i, inputNameCol) = name_trim(arr(i, inputNameCol))
    
    'Removing strings we don't want from start of category names
    For Each v In Array("zzz", "cvl", "cl")
        If LCase(Left(arr(i, inputCatCol), Len(v))) = v Then arr(i, inputCatCol) = Mid(arr(i, inputCatCol), Len(v) + 1)
    Next v
    
    'Apply name trim to description
    arr(i, inputCatCol) = name_trim(arr(i, inputCatCol))
    
Next i

'Determining the number of columns necessary
Set dict = CreateObject("Scripting.Dictionary")

'Getting all unique category names
For i = 2 To UBound(arr)
    dict(arr(i, inputCatCol)) = 1
Next i

'Sorting these category names in ascending order
categories = quicksort(dict.keys())

'Creating an array of each person's name, their highest pay level, and an array of the different
'pay levels they have

'Creating the array to start with
'The reason this is horizontal is because it is going to be constantly resized, and you can only
'resize the last dimension in VBA arrays.
ReDim arr2(3, 1) As Variant
arr2(1, 1) = arr(2, inputNameCol)
arr2(2, 1) = arr(2, inputRateCol)
arr2(3, 1) = Array(arr(2, inputRateCol))

'Populating the new array
For i = 3 To UBound(arr)
    v = array_row(arr2, 1, arr(i, inputNameCol))
    
    'If the person already exists in the new array
    If v Then
        
        'Check that the rate is in the list of rates for this person
        If Not is_in_array(arr2(3, v), arr(i, inputRateCol)) Then
            'If not then add it
            arr2(3, v) = add_to_array(arr2(3, v), arr(i, inputRateCol))
            
            'If the rate for this person is higher than the current highest we'll replace it
            If arr(i, inputRateCol) > arr2(2, v) Then
                arr2(2, v) = arr(i, inputRateCol)
            End If
        End If
        
    'If they don't
    Else
        'New upper bound
        j = UBound(arr2, 2) + 1
        'Resizing the array
        ReDim Preserve arr2(3, j) As Variant
        'Inputting a new row
        arr2(1, j) = arr(i, inputNameCol)
        arr2(2, j) = arr(i, inputRateCol)
        arr2(3, j) = Array(arr(i, inputRateCol))
        
    End If
    
Next i

'Sorting the arrays in column 3 in descending order
For i = 1 To UBound(arr2, 2)
    
    If UBound(arr2(3, i)) > 1 Then
        arr2(3, i) = quicksort(arr2(3, i), True)
    End If
    
Next i

'Sorting this entire array in descending order based on column 2 (highest rate)
arr2 = horizontal_array_sort(arr2, 2)

'Using a third array for the final table
'This will draw the data from the first array and the order from the second array.
'Determining the number of rows required:
j = 0
For i = LBound(arr2, 2) To UBound(arr2, 2)
    j = j + UBound(arr2(3, i))
Next i

'It is plus 4 here because we need to allow 3 columns at the start for the person's name,
'their position, and their rate. It is plus 4 instead of 3 because the categories array starts
'at base 0 and not base 1 (even though option base 1 is selected) because it came from dictionary
'keys and these always start at base 0.
ReDim arr3(j, UBound(categories) + 4) As Variant

'Reset the dictionary and make the categories correspond to their columns in arr3
Set dict = CreateObject("Scripting.Dictionary")
'Starting as 4, as per above comment - we are allowing for the additional first 3 columns
i = 4
For Each v In categories
    dict(v) = i
    i = i + 1
Next v

'Use arr2 to populate arr3
'The following block creates a row for each person and rate combintation and then populates
'all the category columns with zeroes
i = 1

For j = 1 To UBound(arr2, 2)
    
    For Each v In arr2(3, j)
        
        arr3(i, 1) = arr2(1, j)
        arr3(i, 3) = v
        
        k = 4
        For Each v2 In categories
            arr3(i, k) = 0
            k = k + 1
        Next v2
        
        i = i + 1
        
    Next v
    
Next j

'Use arr to populate arr3
For i = 2 To UBound(arr)
    
    'Find the row in arr3 that matches this person and rate in arr
    For j = 1 To UBound(arr3)
        If arr(i, inputNameCol) = arr3(j, 1) And arr(i, inputRateCol) = arr3(j, 3) Then
            Exit For
        End If
    Next j
    
    'If the person currently does not have a position specified, this will populate one
    If IsEmpty(arr3(j, 2)) Then
        arr3(j, 2) = return_position(arr(i, inputDateCol), arr(i, inputRateCol))
    End If
    'This will tally up their hours for each category
    arr3(j, dict(arr(i, inputCatCol))) = arr3(j, dict(arr(i, inputCatCol))) + arr(i, inputHoursCol)
    
Next i

'Creating a new sheet for the finished product to be stored in
'Starting by making sure the name of the worksheet we want to create is unique
myStr = "WIP Table"
i = 1
matchFound = False

'This loop checks if we already have a sheet called 'WIP Table'. If we do, it sees if we have a
'sheet called 'WIP Table 1' and so on and so forth until it finds a free sheet name.
Do While True
    For Each ws In wb.Worksheets
        If ws.name = myStr Then
            matchFound = True
            Exit For
        End If
    Next
    
    If Not matchFound Then Exit Do
    
    i = i + 1
    myStr = "WIP Table " & i
    matchFound = False
    
Loop

'Adding worksheet
Set ws = wb.Worksheets.Add(Type:=xlWorksheet)
With ws
    .name = myStr
    .Activate
End With

'Setting up the sheet
With ws
    .Cells.Font.name = "Arial"
    .Cells.Font.Size = 8
    .Range("A1").Value = "Staff Name"
    .Range("B1").Value = "Classification"
    .Range("C1").Value = "Total Hours"
    .Range("C1:C2").Merge
    .Range("D1").Value = "Hourly Rates ($)"
    .Range("D1:D2").Merge
    .Range("E1").Value = "Total $ (excl GST)"
    .Range("E1:E2").Merge
    
    'Inputting category specific headers onto the new sheet
    i = 6
    For Each v In categories
        
        .Cells(1, i).Value = v
        .Cells(1, i).HorizontalAlignment = xlCenter
        .Range(Cells(1, i), Cells(1, i + 1)).Merge
        .Cells(2, i).Value = "Hours"
        .Cells(2, i + 1).Value = "$"
        i = i + 2
        
    Next v
    
    'Pasting array 3 into the new sheet
    For i = 1 To UBound(arr3)
        'rowNum is the next available row and is plus 2 because the first two rows have headings
        rowNum = i + 2
        
        .Cells(rowNum, 1) = arr3(i, 1)
        .Cells(rowNum, 2) = arr3(i, 2)
        
        'Putting in the addition formula
        'Neccesary to give user assurance
        j = 6
        myStr = "=F" & rowNum
        For Each v In categories
            If j > 6 Then
                myStr = myStr + "+" & col_let(j) & rowNum
            End If
            j = j + 2
        Next
        .Cells(rowNum, 3) = myStr
        
        'Inputting the persons pay rate
        .Cells(rowNum, 4) = arr3(i, 3)
        'Inputting the total amount charged for this person at this rate as a formula
        .Cells(rowNum, 5) = "=C" & rowNum & "*D" & rowNum
        
        'Inputting the total hours and amount for each category
        'The total amount is also a formula
        'j is 6 and increments in steps of 2 because in the final sheet, category columns start
        'at column 6 and increment in 2s. k is 4 and increments in step 1 because categories start
        'at column 4 in arr3.
        j = 6
        k = 4
        For Each v In categories
            
            .Cells(rowNum, j) = arr3(i, k)
            .Cells(rowNum, j + 1) = "=D" & rowNum & "*" & col_let(j) & rowNum
            j = j + 2
            k = k + 1
            
        Next v
        
    Next i
    
    'Adding totals at the bottom
    'rowNum is the next free row
    rowNum = rowNum + 1
    'j is the number of columns we now have
    j = j - 1
    .Cells(rowNum, 1) = "Total"
    .Cells(rowNum, 3) = "=Sum(C3:C" & (rowNum - 1) & ")"
    .Cells(rowNum, 5) = "=Sum(E3:E" & (rowNum - 1) & ")"
    .Cells(rowNum + 1, 1) = "Add: GST at 10%"
    .Cells(rowNum + 1, 5) = "=E" & rowNum & "*0.1"
    .Cells(rowNum + 2, 1) = "Total (incl GST)"
    .Cells(rowNum + 2, 5) = "=E" & rowNum & "+E" & (rowNum + 1)
    .Cells(rowNum + 3, 1) = "Average hourly rate (excl GST)"
    .Cells(rowNum + 3, 5) = "=E" & rowNum & "/C" & rowNum
    
    For i = 6 To j
        .Cells(rowNum, i) = "=Sum(" & col_let(i) & "3:" & col_let(i) & (rowNum - 1) & ")"
        
        'Average hourly rate for each section
        'Only inputs this on alternative columns, hence the mod operator
        If i Mod 2 = 1 Then
            .Cells(rowNum + 3, i) = "=" & col_let(i) & rowNum & "/" & col_let(i - 1) & rowNum
            .Columns(i).NumberFormat = "#,##0.00; [Red] - #,##0.00 ;""-"""
        Else
            .Columns(i).NumberFormat = "#,##0.0; [Red] - #,##0.0 ;""-"""
        End If
        
    Next i
    
    'Fixing up some of the formatting
    For i = 4 To 5
        .Columns(i).NumberFormat = "#,##0.00"
    Next i
    
    .Columns(1).ColumnWidth = 18
    .Columns(2).EntireColumn.AutoFit
    .Rows("1:2").Font.Bold = True
    .Rows(rowNum).Font.Bold = True
    .Rows(rowNum + 2).Font.Bold = True
    .Range("C1:E2").WrapText = True

End With
    
    'The last row formatting
    With ws.Rows(CStr(rowNum + 3) & ":" & CStr(rowNum + 3))
        .Font.Italic = True
        .NumberFormat = "0.00"
    End With

    'Adding thick borders and removing background
    'Border for second last line
    With ws.Range("A" & (rowNum + 2) & ":E" & (rowNum + 2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With ws.Range("A" & (rowNum + 2) & ":E" & (rowNum + 2)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    'Borders for fourth last line
    With ws.Range("A" & rowNum & ":" & col_let(j) & rowNum).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With ws.Range("A" & rowNum & ":" & col_let(j) & rowNum).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    'Border for first line
    With ws.Range("A1:" & col_let(j) & "2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With ws.Range("A1:" & col_let(j) & "2").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    ws.Range("A1").CurrentRegion.Interior.ColorIndex = 2

'Reverting back to the regular settings
Call optimize_off

Exit Sub

ErrHandler:
'Simple error handler to keep end users out of the code
If Err.Number <> 0 Then
    MsgBox "The following error has occurred. " & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Description: " & Err.Description
End If
       
Call optimize_off

End Sub


