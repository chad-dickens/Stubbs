Attribute VB_Name = "Common"
Option Explicit
Option Base 1

'This module contains general functions and procedures used by other modules

'===========================================================================================================
'GENERAL TOOLS
'===========================================================================================================

Public Sub optimize_on()
'Turns off non-essential Excel features to make code run faster
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .StatusBar = "Working..."
End With

End Sub

Public Sub optimize_off()
'Turns these essential features back on
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .StatusBar = ""
End With

End Sub

Public Function col_let(colNum As Long) As String
'Returns column letter from number
    
    col_let = Split(Cells(1, colNum).Address, "$")(1)

End Function

Public Function name_trim(ByVal name As String) As String
'Will remove any non-letter characters at the start and end of a string
    
    'Removing start characters
    Do While Not Left(name, 1) Like "[a-zA-Z]" And Len(name) > 1
        name = Right(name, Len(name) - 1)
    Loop
    
    'Removing end characters
    Do While Not Right(name, 1) Like "[a-zA-Z]" And Len(name) > 1
        name = Left(name, Len(name) - 1)
    Loop
    
    name_trim = StrConv(name, vbProperCase)

End Function

'===========================================================================================================
'ARRAY AND SORTING TOOLS
'===========================================================================================================

Sub run_quicksort(vArray As Variant, inLow As Long, inHi As Long, Optional descending As Boolean = False)
'Quicksort algorithm to use on one dimesional arrays that will sort in ascending
'or descending order
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)
    
  While (tmpLow <= tmpHi)
    
    If descending Then
        While (vArray(tmpLow) > pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (pivot > vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
    Else
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend

        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
    End If

    If (tmpLow <= tmpHi) Then
       tmpSwap = vArray(tmpLow)
       vArray(tmpLow) = vArray(tmpHi)
       vArray(tmpHi) = tmpSwap
       tmpLow = tmpLow + 1
       tmpHi = tmpHi - 1
    End If
    
  Wend
  
  If descending Then
    If (inLow < tmpHi) Then run_quicksort vArray, inLow, tmpHi, True
    If (tmpLow < inHi) Then run_quicksort vArray, tmpLow, inHi, True
  Else
    If (inLow < tmpHi) Then run_quicksort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then run_quicksort vArray, tmpLow, inHi
  End If
  
End Sub

Public Function quicksort(ByVal myArray As Variant, Optional descending As Boolean = False) As Variant
'For the purpose of creating a function out of quicksort that returns a value
    Call run_quicksort(myArray, LBound(myArray), UBound(myArray), descending)
    quicksort = myArray

End Function

Public Function array_row(anArray As Variant, colNum As Integer, searchTerm As Variant) As Variant
'Will return the row number of requested item in a 2 dimensional array
'If the item does not exist it will return False
'The orientation of the array must be horizontal ie Array(3, 20)

    Dim i As Integer
    
    For i = LBound(anArray, 2) To UBound(anArray, 2)
        If anArray(colNum, i) = searchTerm Then
            array_row = i
            Exit Function
        End If
    Next i
    
    array_row = False

End Function

Public Function is_in_array(anArray As Variant, searchTerm As Variant) As Boolean
'Returns whether item exists in one dimensional array
    
    Dim v As Variant
    
    For Each v In anArray
        If v = searchTerm Then
            is_in_array = True
            Exit Function
        End If
    Next v
    
    is_in_array = False

End Function

Public Function add_to_array(anArray As Variant, addition As Variant) As Variant
'Adds item to one dimensional array

    Dim v As Variant
    v = anArray
    ReDim Preserve v(LBound(anArray) To UBound(anArray) + 1)
    
    v(UBound(v)) = addition
    add_to_array = v

End Function

Sub run_horizontal_array_sort(vArray As Variant, colNum As Long, inLow As Long, inHi As Long)
'Custom function that sorts a horizontal array (ie,  Array(3, 20)) in descending order
'based on column number (first dimension) specified.
'Pretty much just the quicksort function above, but repurposed.
    
    Dim pivot As Variant
    Dim i As Long
    Dim lowBound As Long
    Dim hiBound As Long
    
    lowBound = LBound(vArray)
    hiBound = UBound(vArray)
    
    Dim tmpSwap() As Variant
    ReDim tmpSwap(lowBound To hiBound) As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray(colNum, (inLow + inHi) \ 2)
      
    While (tmpLow <= tmpHi)
      
        While (vArray(colNum, tmpLow) > pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
    
        While (pivot > vArray(colNum, tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
          
        If (tmpLow <= tmpHi) Then
        
            For i = lowBound To hiBound
                 tmpSwap(i) = vArray(i, tmpLow)
            Next i
            
            For i = lowBound To hiBound
                 vArray(i, tmpLow) = vArray(i, tmpHi)
            Next i
            
            For i = lowBound To hiBound
                 vArray(i, tmpHi) = tmpSwap(i)
            Next i
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
      
    Wend
    
    If (inLow < tmpHi) Then run_horizontal_array_sort vArray, colNum, inLow, tmpHi
    If (tmpLow < inHi) Then run_horizontal_array_sort vArray, colNum, tmpLow, inHi
    
End Sub

Public Function horizontal_array_sort(ByVal myArray As Variant, colNum As Long) As Variant
'For the purpose of creating a function out of run_horizontal_array_sort that returns a value
    Call run_horizontal_array_sort(myArray, colNum, LBound(myArray, 2), UBound(myArray, 2))
    horizontal_array_sort = myArray

End Function

Public Function remove_from_array(myArray As Variant, element As Variant) As Variant
'Will remove all instances of an element from a one dimensional array
    
    'Variable to determine how large the new array will be
    Dim count As Long
    count = 0
    
    'Counting the number of items in the new array after removal
    Dim v As Variant
    For Each v In myArray
        If v <> element Then
            count = count + 1
        End If
    Next v
    
    'Returning blank array and exiting if no items left after removal
    If count = 0 Then
        remove_from_array = Array()
        Exit Function
    End If
    
    'Building new array
    Dim newArray() As Variant
    ReDim newArray(count)
    count = 0
    For Each v In myArray
        If v <> element Then
            count = count + 1
            newArray(count) = v
        End If
    Next v
    
    remove_from_array = newArray

End Function

