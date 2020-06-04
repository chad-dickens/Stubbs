Attribute VB_Name = "MainModule"
Option Explicit
Option Base 1
Public ExitButton As Boolean
Public OpeningBalance As Currency
Public bSelectedMonth As Byte
Public lSelectedYear As Long
Public sPreviousMonthFile As String
Public aBankA

Sub Preparing_Tracking_Files()

On Error GoTo ErrHandler

Call OptimizeOn

'Declaring Variables
Dim dFirstDay As Date
Dim objOutlook As Object, objMail As Object
Dim sEmail As String
Dim vFileD As Office.FileDialog
Dim sBankAFile As String, sBankBFile As String, sSaveFolder As String, sSaveFileName
Dim aBankB
Dim bDaysinmonth As Byte, Vcol As Byte, Ucol As Byte, Wcol As Byte
Dim i As Long, r As Long
Dim sNarrative As String
Dim cRowValue As Currency
Dim Sht As Worksheet
Dim ShtExists As Boolean
Dim iAnswer As Integer
Dim DatesRight As Boolean
Dim BankAWrongDate As String
Dim BankBWrongDate As String
Dim sUserName As String
Dim sFileExists As String
Dim boolMatched As Boolean
Dim msg1 As String, msg2 As String

'Setting variable values
sEmail = "myemail@gmail.com"

'Setting default value
ExitButton = False
ShtExists = False
OpeningBalance = 0
bSelectedMonth = 0
lSelectedYear = 0
sPreviousMonthFile = ""
DatesRight = True
boolMatched = False

'Setting our sUserName Variable value
    If InStr(1, Application.UserName, " ") > 0 Then
        sUserName = Split(Application.UserName, " ")(0)
    Else
        sUserName = Application.UserName
    End If

'Setting Error message for when the file selected by users does not confirm to the rules
msg2 = "Hi " & sUserName & "," & vbNewLine
msg2 = msg2 & "The file you have selected does not conform to the rules of this spreadsheet. "

'Showing Front Userform
    'This gives the user the options of either displaying detailed instructions on the process, proceeding, or exiting
WelcomeForm.Show

    'Checking whether our user has clicked exit
    If ExitButton = True Then GoTo ExitButton1

'Showing Month Selection Userform
PeriodSelectionForm.Show

    'Checking whether our user has clicked exit
    If ExitButton = True Then GoTo ExitButton1

OpeningBalance1:
'Showing the Opening Balance Form
OpeningBalanceForm.Show

    'Checking whether our user has clicked exit
    If ExitButton = True Then GoTo ExitButton1

'Opening the Previous Month File and Checking that there is an opening balance
If OpeningBalance = 0 Then

    If sPreviousMonthFile = ThisWorkbook.FullName Then
        MsgBox msg2 & "You cannot choose this current workbook as your opening balance file. Please choose again."
        GoTo OpeningBalance1
    End If

    With Workbooks.Open(sPreviousMonthFile)
    
        For Each Sht In .Worksheets
            If Sht.Name = "Bank A" Then ShtExists = True
        Next Sht
        
        'If we can't find this sheet name it means the file is wrong
        If ShtExists = False Then
            MsgBox "A worksheet with the name 'Bank A' could not be found in the file you selected. Please choose your file again."
            .Close False
            GoTo OpeningBalance1
        End If
        
        With .Sheets("Bank A").Range("Q1")
        
            If IsNumeric(.Value) And Len(.Value) > 0 Then
                OpeningBalance = .Value
            Else
                MsgBox msg2 & "While the file you selected does have a sheet called 'Bank A', range Q1 does not hold a numeric value. Please choose again."
                .Close False
                GoTo OpeningBalance1
            End If
            
        End With
        
        .Close False
    
    End With
End If

'Double checking that users are happy with the opening balance
iAnswer = MsgBox("The opening balance is " & CustomFormat(OpeningBalance) & ". Do you wish to proceed?", vbYesNo + vbQuestion + vbDefaultButton1, "Opening Balance")

        If iAnswer = vbNo Then
            OpeningBalance = 0
            GoTo OpeningBalance1
        End If

BankAFileSelect:
Set vFileD = Application.FileDialog(msoFileDialogFilePicker)
'Getting users to choose their Bank A file
With vFileD
    .Title = "Please Select Your Bank A Data File"
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx, *.csv, *.xls, *.xlsm"
    .Show
    
    'Setting the path of this workbook
    sBankAFile = .SelectedItems(1)
End With

'Opening the Bank A Workbook
With Workbooks.Open(Filename:=sBankAFile, Local:=True)
    
    'Checking that the file only has 1 sheet
    If .Sheets.Count > 1 Then
        MsgBox msg2 & "Your file has more than 1 sheet. Please choose again."
        .Close False
        GoTo BankAFileSelect
    End If
    
    'Checking that there is actually data in the sheet
    If .Sheets(1).Range("A1").CurrentRegion.Count < 12 Then
        MsgBox msg2 & "Your selected file does not have enough data in it to be valid. Please choose again."
        .Close False
        GoTo BankAFileSelect
    End If
    
    With .Sheets(1)
        'Putting all the data into an array
        aBankA = .Range("A1").CurrentRegion.Value
    End With
        
        .Close False


End With

'Check that the data in this Bank A file was correct
'Resting the boolean value
DatesRight = True

'Setting our error message
BankAWrongDate = "Hi " & sUserName & "," & vbNewLine & "It looks like the Bank A file you have selected does not cover the period of "
BankAWrongDate = BankAWrongDate & MonthName(bSelectedMonth) & " " & lSelectedYear & ". "
BankAWrongDate = BankAWrongDate & "Please choose your file again."
    
    'Checking Column count
    If UBound(aBankA, 2) <> 6 Then
        MsgBox msg2 & "Your data region does not have 6 columns. Please choose again."
        GoTo BankAFileSelect
    End If
    
    'Checking all the column headings are right
    If aBankA(1, 1) <> "Process date" Or _
        aBankA(1, 2) <> "Description" Or _
        aBankA(1, 3) <> "Currency Code" Or _
        aBankA(1, 4) <> " Debit" Or _
        aBankA(1, 5) <> " Credit" Or _
        aBankA(1, 6) <> " Balance" Then
    
        MsgBox msg2 & "Your column headings are incorrect. Please choose again. If you are not sure of the correct format, exit the process and refer to the instructions."
        GoTo BankAFileSelect
    
    End If
    
    'Checking the dates
    For r = 2 To UBound(aBankA)
        If Year(aBankA(r, 1)) <> lSelectedYear Then DatesRight = False
        If Month(aBankA(r, 1)) <> bSelectedMonth Then DatesRight = False
        
            If DatesRight = False Then
                MsgBox BankAWrongDate
                GoTo BankAFileSelect
            End If
    Next r


BankBFileSelect:
'Getting Users to Select their Bank B data file

With vFileD
    .Title = "Please Select Your Bank B Data File"
    .Show
    
    'Setting the path of this workbook
    sBankBFile = .SelectedItems(1)
End With

'Opening the Bank B Workbook
With Workbooks.Open(sBankBFile)
    
    'Checking that the file only has 1 sheet
    If .Sheets.Count > 1 Then
        MsgBox msg2 & "Your file has more than 1 sheet. Please choose again."
        .Close False
        GoTo BankBFileSelect
    End If
    
    'Checking that there is actually data in the sheet
    If .Sheets(1).Range("A1").CurrentRegion.Count < 18 Then
        MsgBox msg2 & "Your selected file does not have enough data in it to be valid. Please choose again."
        .Close False
        GoTo BankBFileSelect
    End If
    
    'Putting all the data into an array
    aBankB = .Sheets(1).Range("A1").CurrentRegion.Value
    .Close False

End With

'Check that the data in this Bank B file was correct
'Reseting the boolean value
DatesRight = True

'Setting our error message
BankBWrongDate = "Hi " & sUserName & "," & vbNewLine & "It looks like the Bank B file you have selected does not cover the period of "
BankBWrongDate = BankBWrongDate & MonthName(bSelectedMonth) & " " & lSelectedYear & ". "
BankBWrongDate = BankBWrongDate & "Please choose your file again."

    'Checking Column count
    If UBound(aBankB, 2) <> 9 Then
        MsgBox msg2 & "Your data region does not have 9 columns. Please choose again."
        GoTo BankBFileSelect
    End If
    
    'Checking all the column headings are right
    If aBankB(1, 1) <> "TRAN_DATE" Or _
        aBankB(1, 2) <> " ACCOUNT_NO" Or _
        aBankB(1, 3) <> " SEGMENT_ID" Or _
        aBankB(1, 4) <> " CCY" Or _
        aBankB(1, 5) <> " CLOSING_BAL" Or _
        aBankB(1, 6) <> " AMOUNT" Or _
        aBankB(1, 7) <> " TRAN_CODE" Or _
        aBankB(1, 8) <> " NARRATIVE" Or _
        aBankB(1, 9) <> " SERIAL" Then
    
        MsgBox msg2 & "Your column headings are incorrect. Please choose again. If you are not sure of the correct format, exit the process and refer to the instructions."
        GoTo BankBFileSelect
    
    End If

    'Checking the dates
    For r = 2 To UBound(aBankB)
        If Left(aBankB(r, 1), 4) <> lSelectedYear Then DatesRight = False
        If Val(Mid(aBankB(r, 1), 5, 2)) <> bSelectedMonth Then DatesRight = False
        
            If DatesRight = False Then
                MsgBox BankBWrongDate
                GoTo BankBFileSelect
            End If
    Next r
     

'Getting Users to Select where they want to save their new report
Set vFileD = Application.FileDialog(msoFileDialogFolderPicker)

With vFileD
    .AllowMultiSelect = False
    .Title = "Please select where you want to save your new report:"
    .Show
    
    'Setting the path of our save folder
    sSaveFolder = .SelectedItems(1)
End With

'Setting the full path of the file we'll be saving
sSaveFileName = sSaveFolder & "\Bank Statement Tracking " & MonthName(bSelectedMonth, True) & Right(lSelectedYear, 2) & ".xlsx"

'Checking this file does no already exist and if it does, asking if the user wants to overwrite it
sFileExists = "Hi " & sUserName & "," & vbNewLine
sFileExists = sFileExists & "A file currently exists with the filepath of " & sSaveFileName & ". "
sFileExists = sFileExists & "Do you wish to overwrite it? If you click 'No', the Macro will be exited."

If Dir(sSaveFileName) <> vbNullString Then
    iAnswer = MsgBox(sFileExists, vbYesNo + vbQuestion + vbDefaultButton1, "Overwrite file?")
    If iAnswer = vbNo Then GoTo ExitButton1
End If

'Clearing out the old stuff
With vBankB
.Range("A4:A34").ClearContents
.Range("D4:E34").ClearContents
.Range("D4:E34").ClearComments
.Range("G4:J34").ClearContents
.Range("G4:J34").ClearComments
End With

With vBankA
.Range("A5:A35").ClearContents
.Range("B5:E35").ClearContents
.Range("A40:A70").ClearContents
.Range("L5:O35").ClearContents
.Range("C40:D70").ClearContents
.Range("F40:K70").ClearContents
.Range("M40:P70").ClearContents
.Range("F40:K70").ClearComments
.Range("M40:P70").ClearComments
End With

'Putting our new dates in

        'Finding the last day in the month
        bDaysinmonth = Day(DateSerial(lSelectedYear, bSelectedMonth + 1, 1) - 1)
        'Getting first day of the month, this will be reset a few times
        dFirstDay = DateSerial(lSelectedYear, bSelectedMonth, 1)

    'Updating vBankB Dates First
    For i = 4 To bDaysinmonth + 3
        vBankB.Cells(i, 1).Value = dFirstDay
        dFirstDay = dFirstDay + 1
    Next i

'Updating vBankA Dates Second
dFirstDay = DateSerial(lSelectedYear, bSelectedMonth, 1)

For i = 5 To bDaysinmonth + 4
    vBankA.Cells(i, 1).Value = dFirstDay
    dFirstDay = dFirstDay + 1
Next i

dFirstDay = DateSerial(lSelectedYear, bSelectedMonth, 1)

For i = 40 To bDaysinmonth + 39
    vBankA.Cells(i, 1).Value = dFirstDay
    dFirstDay = dFirstDay + 1
Next i

'Moving The vBankBDump Values To Their Correct Location
    'Rule is that If the narrative string contains Merchant Fees if it will be moved there, if it contains chargeback it will be moved there, otherwise it will go to normal

For r = 4 To bDaysinmonth + 3

        'We'll use this as our chargeback column and reset it for every new date
        Vcol = 7
        'We'll use this as our general column
        Ucol = 4
        
        For i = 2 To UBound(aBankB)
            
            'If the day is the same
            If Val(Right(aBankB(i, 1), 2)) = Day(vBankB.Cells(r, 1).Value) Then
            
                sNarrative = aBankB(i, 8)
                cRowValue = aBankB(i, 6)
                
                    If InStr(1, sNarrative, "Merchant Fee", vbTextCompare) >= 1 Then
                    
                        With vBankB.Cells(r, 10)
                        .Value = .Value + cRowValue
                        
                            'Checking if we're combining two or more values together
                            If .Value > cRowValue Then
                                .ClearComments
                                .AddComment "Batched due to lack of space"
                            Else
                                .AddComment sNarrative
                            End If
                        
                        End With
                        
                        Wcol = Wcol + 1
                    ElseIf InStr(1, sNarrative, "Chargeback", vbTextCompare) >= 1 Then
                        
                        If Vcol < 10 Then
                            With vBankB.Cells(r, Vcol)
                                .Value = cRowValue
                                .AddComment sNarrative
                            End With
                        Else
                            With vBankB.Cells(r, 9)
                                .Value = .Value + cRowValue
                                .ClearComments
                                .AddComment "Batched due to lack of space"
                            End With
                        End If
                        
                        Vcol = Vcol + 1
                    
                    Else
                        
                        If Ucol < 6 Then
                            With vBankB.Cells(r, Ucol)
                                .Value = cRowValue
                                .AddComment sNarrative
                            End With
                        Else
                            
                            With vBankB.Cells(r, 5)
                                .Value = .Value + cRowValue
                                .ClearComments
                                .AddComment "Batched due to lack of space"
                            End With
                            
                        End If
                            
                        Ucol = Ucol + 1
                        
                    End If
            
            End If
            
        Next i
        
Next r

'Putting in our Bank A Values now for the top part
    'Using column 4 as a way of dumping 'matched' markers
    
        'Putting in the opening balance
        vBankA.Range("Q1").Value = OpeningBalance

For r = 5 To bDaysinmonth + 4

        'We'll use this as our cheque deposit column and reset it for every new date
        Vcol = 6
        'We'll use this as our total bank transfer column
        Ucol = 13
        'We'll use this as our payplus column
        Wcol = 3

    For i = UBound(aBankA) To 2 Step -1
    
        'If the day is the same and we haven't already matched it
            
            If Day(aBankA(i, 1)) = Day(vBankA.Cells(r, 1).Value) And aBankA(i, 3) <> "MATCHED" Then
                
                sNarrative = aBankA(i, 2)
                cRowValue = aBankA(i, 5)
                
                    'If its not a credit then its a debit
                    If cRowValue = 0 Then cRowValue = -aBankA(i, 4)
                    
                    'Looking for Amex first
                    If InStr(1, sNarrative, "amex", vbTextCompare) > 0 Then
                    
                        With vBankA.Cells(r, 3)
                            .Value = .Value + cRowValue
                        End With
                        
                        With vBankA.Cells(r, 2)
                            .Value = .Value + Val(Split(WorksheetFunction.Trim(sNarrative), " ")(5))
                        End With
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "diners", vbTextCompare) > 0 Then
                    
                        With vBankA.Cells(r, 5)
                            .Value = .Value + cRowValue
                        End With
                        
                        With vBankA.Cells(r, 4)
                            .Value = .Value + Val(Split(WorksheetFunction.Trim(sNarrative), " ")(4))
                        End With
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "securepay", vbTextCompare) > 0 And cRowValue < 0 Then
                        
                        With vBankA.Range("M" & r)
                            .Value = .Value + cRowValue
                        End With
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "payment centre fee", vbTextCompare) > 0 And cRowValue < 0 Then
                        
                        With vBankA.Range("L" & r)
                            .Value = .Value + cRowValue
                        End With
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "returned cheque", vbTextCompare) > 0 And cRowValue < 0 Then
                        
                        With vBankA.Range("N" & r)
                            .Value = .Value + cRowValue
                        End With
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "payplus", vbTextCompare) > 0 And cRowValue > 0 Then
                        
                        If Wcol < 5 Then
                        
                            With vBankA.Cells(r + 35, Wcol)
                                .Value = .Value + cRowValue
                            End With
                            
                        Else
                            With vBankA.Cells(r + 35, 4)
                                .Value = .Value + cRowValue
                            End With
                            
                        End If
                        
                        Wcol = Wcol + 1
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf InStr(1, sNarrative, "cash dep", vbTextCompare) > 0 And cRowValue > 0 Or _
                    InStr(1, sNarrative, "chq dep", vbTextCompare) > 0 And cRowValue > 0 Then
                        
                        If Vcol < 12 Then
                            With vBankA.Cells(r + 35, Vcol)
                                .Value = cRowValue
                                .AddComment sNarrative
                            End With
                        Else
                            With vBankA.Cells(r + 35, 11)
                                .Value = .Value + cRowValue
                                .ClearComments
                                .AddComment "Batched due to lack of space"
                            End With
                        End If
                        
                        Vcol = Vcol + 1
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    ElseIf cRowValue > 0 Then
                        
                        If Ucol < 17 Then
                            With vBankA.Cells(r + 35, Ucol)
                                .Value = cRowValue
                                .AddComment sNarrative
                            End With
                        Else
                            With vBankA.Cells(r + 35, 16)
                                .Value = .Value + cRowValue
                                .ClearComments
                                .AddComment "Batched due to lack of space"
                            End With
                        End If
                        
                        Ucol = Ucol + 1
                        
                        'Marking it so we don't match it again
                        aBankA(i, 3) = "MATCHED"
                        
                    End If
                
            End If
    
    Next i

Next r

'Updating all the formulas
Application.Calculate

'Deciding if there are any unmatched Bank A Debits
For r = 2 To UBound(aBankA)
    If aBankA(r, 3) <> "MATCHED" Then
        boolMatched = True
        Exit For
    End If
Next r

    'Showing the userform if there were unmatched debits
    If boolMatched = True Then UnmatchedDebitsForm.Show
    
    'If form was closed we want to exit the process here
    If ExitButton = True Then GoTo ExitButton1

'Letting the user know that all items have been matched
msg1 = "Hi " & sUserName & "," & vbNewLine
msg1 = msg1 & "All items have now been matched. Please double check the opening and closing balances against the PDF "
msg1 = msg1 & "bank statements before you send the email."
MsgBox msg1

'Setting our mouse to Cells A1 in both sheets
vBankB.Select
Range("A1").Select

vBankA.Select
Range("A1").Select

vStart.Select
Range("A1").Select

'Opening a new workbook
With Workbooks.Add
    
    vBankA.Copy .Sheets(1)
    vBankB.Copy .Sheets(2)
    .Sheets(3).Delete
    .Sheets(1).Select
    .SaveAs sSaveFileName
    .Close
    
End With

'Clearing the sheets out again
With vBankB
.Range("A4:A34").ClearContents
.Range("D4:E34").ClearContents
.Range("D4:E34").ClearComments
.Range("G4:J34").ClearContents
.Range("G4:J34").ClearComments
End With

With vBankA
.Range("Q1:Q2").ClearContents
.Range("A5:A35").ClearContents
.Range("B5:E35").ClearContents
.Range("A40:A70").ClearContents
.Range("L5:O35").ClearContents
.Range("C40:D70").ClearContents
.Range("F40:K70").ClearContents
.Range("M40:P70").ClearContents
.Range("F40:K70").ClearComments
.Range("M40:P70").ClearComments
End With

'Setting up the email we'll be displaying

        Set objOutlook = CreateObject("Outlook.Application")
        Set objMail = objOutlook.CreateItem(0)

        With objMail
            .To = sEmail
            .Subject = "Bank Statement Tracking " & MonthName(bSelectedMonth) & " " & lSelectedYear
            .HTMLBody = "Hi all,<br><br>Here is the bank statement tracking excel file for " & MonthName(bSelectedMonth) & " " & lSelectedYear & ".<br><br>Please let me know if there are any issues.<br><br>Thanks,"
            .Attachments.Add sSaveFileName
            .Display
        End With

Call OptimizeOff

Exit Sub

'Exit Button Location
ExitButton1:
    
    MsgBox "Process Exited"
    Call OptimizeOff

Exit Sub

'Our Error Handler
ErrHandler:
    'This is error number that occurs when someone closes a filedialog box without selecting a file or folder
    If Err.Number = 5 Then
        MsgBox "FileDialog Exited"
    Else
        MsgBox "An error has occurred. " & Err.Number & Err.Description
    End If
    
Call OptimizeOff

End Sub

