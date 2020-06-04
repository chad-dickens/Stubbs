Attribute VB_Name = "Main_Module"
Option Explicit

Public DataBaseName As String
Public SalesDataTable As String
Public RatesDataTable As String
Public AddressDataTable As String
Public ExitButton As Boolean

Public Const SalesDataFields = ",ID,Account_Type,Sub_Account_Type,Despatch_Year,Despatch_Month,Despatch_Date,Qtr,Despatch_ID,Serial_Number,Origin,Destination,Mail_Category,Class,Subclass,PL,Country_Office,Country_Code,Country_Name,ETOE Bilaterals_Source,No_of_ItRates,Weight_Kgs"
Public Const RatesDataFields = ",ID,Rate_Reference,Operator,Despatch_Year,Sub_Account_Type,PL,Country_Code,Mail_Category,Subclass,Rate_Ltr_Kg,Rate_Ltr_Itm,Rate_Bulk_Kg,Rate_Bulk_Itm"
Public Const AddressDataFields = ",ID,Country_Code,Physical_Address,Country_Name,Email_Address"
'

Sub Choose_Database()

'Prompts user to select the Access database they wish to use.
'Once selected, will check this database to ensure that it has the same table names as default.
'If this is not the case, it will let the user assign the default table names to their alternate names.
'Will then check that these tables contain the right field names.
'If all of this criteria is met, the main userform will be displayed.

Call Optimize_On

'Choosing the database file
DataBaseName = FilePicker("Select your database file", "MS Access", "*.accdb")
If DataBaseName = "Error" Then Exit Sub

'Checking it has the right table names
Dim DefaultTableNames() As Variant
DefaultTableNames = Array("sales_data", "rates_data", "address_data")

'Setting the public variable values
If TablesExist(DefaultTableNames, DataBaseName) Then

    SalesDataTable = DefaultTableNames(0)
    RatesDataTable = DefaultTableNames(1)
    AddressDataTable = DefaultTableNames(2)
    
Else

    MsgBox "Hi " & UserName() & ", it looks the Access database you have selected does not contain tables with names we were expecting. " & _
           "Please use the following Userform to select the right tables."

    'Showing a userform to let the user decide which tables they want to use
    DBTableAllocator.Show
    If ExitButton = True Then GoTo ExitLine
    
End If

'Checking the tables all have the right fields
If Not FieldsExist(DataBaseName, SalesDataTable, SalesDataFields) Or _
   Not FieldsExist(DataBaseName, RatesDataTable, RatesDataFields) Or _
   Not FieldsExist(DataBaseName, AddressDataTable, AddressDataFields) Then
   
   MsgBox "Hi " & UserName() & ", Unfortuneately the fields of the tables of the database you selected are not correct. " & _
          "Please refer to the instructions, rectify, and then run again."
    
   GoTo ExitLine
   
End If

'All checks complete. Running the main userform
Main_Window.Show

'If the user exits prematurely
ExitLine:

Call Optimize_Off

End Sub

Sub Create_Invoices(CustomerArray As Variant, Category As String, FilePath As String, Year As String, Quarter As String)
'Main procedure to create all PDF invoices and prepare emails
'@CustomerArray should be a string array of all the customers that need invoices created. Only one dimension.
'@Category is the sub account type
'@Filepath is where the invoices should be saved
'@Year is the year of the records
'@Quarter is the quarter of the records

'Setting the category, name, quarter and year fields on the invoices
With Sht_Summary
    .Range("N1").Value = Year
    .Range("N2").Value = Quarter
    .Range("M6").Value = Category
End With

'Starting our loop to go through each item in the customerarry
Dim Customer As Variant
Dim SQL As String
Dim AddressArray As Variant
Dim SalesArray As Variant
Dim RowNum As Long
Dim RowStart As Long
Dim R As Long, Y As Long
Dim RateReference As String
Dim SDT As String
Dim RDT As String
Dim ADT As String
Dim RP As Long
Dim EmailCount As Long
Dim LastRow As Long, CurrentRow As Long
Dim CurrentString As String, NewString As String

Dim A As String, B As String, C As String, D As String, E As String, F As String, G As String, H As String, I As String, J As String, K As String
Dim DY As String, PL As String, RR As String, DD As String, SN As String
Dim SummaryFile As String, DetailFile As String

Dim outlookApp As Outlook.Application
Dim myMail As Outlook.MailItem

SDT = SalesDataTable
RDT = RatesDataTable
ADT = AddressDataTable
EmailCount = 0

For Each Customer In CustomerArray

SummaryFile = FilePath & "\" & Customer & " " & Category & " " & Quarter & " " & Year & " Summary Invoice.pdf"
DetailFile = FilePath & "\" & Customer & " " & Category & " " & Quarter & " " & Year & " Detail Invoice.pdf"

'==========================================================================================================
'INVOICE SUMMARY
'==========================================================================================================
    
    'Getting our physical address and email address
          SQL = SQLString("SELECT DISTINCT ?.?, ?.? ", ADT, Split(AddressDataFields, ",")(3), _
                                                       ADT, Split(AddressDataFields, ",")(5))
    SQL = SQL + SQLString("FROM ? ", ADT)
    SQL = SQL + SQLString("INNER JOIN ? ON ?.? = ?.? ", SDT, ADT, Split(AddressDataFields, ",")(2), _
                                                        SDT, Split(SalesDataFields, ",")(17))
    SQL = SQL + SQLString("WHERE ?.? = '?';", SDT, Split(SalesDataFields, ",")(18), Customer)
    
    AddressArray = SQLSelect(DataBaseName, SQL)
    
    'Pasting in the physical address and saving the email address for later
    Sht_Summary.Range("A1").Value = AddressArray(0, 0)
    
    'Getting the summary lines
    A = Split(SalesDataFields, ",")(16)
    B = Split(SalesDataFields, ",")(10)
    C = Split(SalesDataFields, ",")(11)
    D = Split(SalesDataFields, ",")(12)
    E = Split(SalesDataFields, ",")(14)
    
    F = Split(SalesDataFields, ",")(20)
    G = Split(SalesDataFields, ",")(21)
    
    H = Split(RatesDataFields, ",")(10)
    I = Split(RatesDataFields, ",")(11)
    J = Split(RatesDataFields, ",")(12)
    K = Split(RatesDataFields, ",")(13)
    
    DY = Split(SalesDataFields, ",")(4)
    PL = Split(SalesDataFields, ",")(15)
    RR = Split(RatesDataFields, ",")(2)
    
          SQL = SQLString("SELECT ?.?, ?.?, ?.?, ?.?, ?.?, SUM(?.?), SUM(?.?), ?.?, ?.?, ?.?, ?.? ", _
                          SDT, A, SDT, B, SDT, C, SDT, D, SDT, E, _
                          SDT, F, SDT, G, _
                          RDT, H, RDT, I, RDT, J, RDT, K)
                          
    SQL = SQL + SQLString("FROM ? ", SDT)
    
    SQL = SQL + SQLString("INNER JOIN ? ON ?.? + ?.? + ?.? + ?.? + ?.? = ?.? ", _
                          RDT, _
                          SDT, DY, SDT, A, SDT, PL, SDT, D, SDT, E, _
                          RDT, RR)
    
    SQL = SQL + SQLString("WHERE ?.? = '?' AND ?.? = '?' AND ?.? = '?' AND ?.? = ? ", _
                                                            SDT, Split(SalesDataFields, ",")(18), Customer, _
                                                            SDT, Split(SalesDataFields, ",")(3), Category, _
                                                            SDT, DY, Year, _
                                                            SDT, Split(SalesDataFields, ",")(7), CInt(Right(Quarter, 1)))
                                                    
    SQL = SQL + SQLString("GROUP BY ?.?, ?.?, ?.?, ?.?, ?.?, ?.?, ?.?, ?.?, ?.? ", _
                          SDT, A, SDT, B, SDT, C, SDT, D, SDT, E, RDT, H, RDT, I, RDT, J, RDT, K)
                          
    SQL = SQL + SQLString("ORDER BY ?.?, ?.?, ?.?, ?.?, ?.?;", _
                          SDT, A, SDT, B, SDT, C, SDT, D, SDT, E)
    
    SalesArray = SQLSelect(DataBaseName, SQL)
    
    'Inserting the required rows
    RowStart = 13
    RowNum = UBound(SalesArray, 2)
    If RowNum > 0 Then
        Sht_Summary.Range("A" & RowStart).EntireRow.Offset(1).Resize(RowNum).Insert Shift:=xlDown
    End If
    
    'Inserting our new Array
    With Sht_Summary
        .Range("A" & RowStart & ":G" & (RowStart + RowNum)).Value = WorksheetFunction.Transpose(SalesArray)
    'Calculating the sums for items and weight
        .Range("F" & (RowStart + 1 + RowNum)).Formula = "=SUM(F13:F" & (RowStart + RowNum) & ")"
        .Range("G" & (RowStart + 1 + RowNum)).Formula = "=SUM(G13:G" & (RowStart + RowNum) & ")"
    
        For R = 0 To RowNum
            RP = R + RowStart
            
            'Inputting the rates
            .Cells(RP, 11).Value = SalesArray(7, R)
            .Cells(RP, 8).Value = SalesArray(8, R)
            .Cells(RP, 12).Value = SalesArray(9, R)
            .Cells(RP, 9).Value = SalesArray(10, R)
            
            'Calculating the rate charges
            .Range("J" & RP).Formula = Replace("=IF(H?=0, F?*I?, F?*H?)", "?", RP)
            .Range("M" & RP).Formula = Replace("=IF(K?=0, G?*L?, G?*K?)", "?", RP)
            .Range("N" & RP).Formula = Replace("=SUM(J?,M?)", "?", RP)
            
        Next R
        
        'Grand totals formulas
        .Range("J" & (RowStart + 1 + RowNum)).Formula = Replace("=SUM(J" & RowStart & ":J?)", "?", RowStart + RowNum)
        .Range("M" & (RowStart + 1 + RowNum)).Formula = Replace("=SUM(M" & RowStart & ":M?)", "?", RowStart + RowNum)
        .Range("N" & (RowStart + 2 + RowNum)).Formula = Replace("=SUM(N" & RowStart & ":N?)", "?", RowStart + RowNum)
        
        'Fixing up the formatting
        .Range("A" & RowStart & ":N" & RowStart + RowNum).Borders.LineStyle = xlContinuous
                    
        'Exporting to PDF
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=SummaryFile, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        'Clearing the sheet
        If RowNum > 0 Then
            .Rows((RowStart + 1) & ":" & (RowStart + RowNum)).EntireRow.Delete
        End If
        
        .Range("A" & RowStart & ":N" & RowStart).ClearContents
        
    End With
    
    '==========================================================================================================
    'INVOICE DETAIL
    '==========================================================================================================
    DD = Split(SalesDataFields, ",")(6)
    SN = Split(SalesDataFields, ",")(9)
    
    'Revisiting our SQL query and simplifying it
          SQL = SQLString("SELECT ?, ?, ?, ?, ?, ?, ?, ? ", _
                          DD, B, C, D, E, SN, F, G)
                          
    SQL = SQL + SQLString("FROM ? ", SDT)
    
    SQL = SQL + SQLString("WHERE ? = '?' AND ? = '?' AND ? = '?' AND ? = ? ", _
                                                            Split(SalesDataFields, ",")(18), Customer, _
                                                            Split(SalesDataFields, ",")(3), Category, _
                                                            DY, Year, _
                                                            Split(SalesDataFields, ",")(7), CInt(Right(Quarter, 1)))
                                                    
    SQL = SQL + SQLString("ORDER BY ?, ?, ?, ?;", _
                          B, C, D, E)
    
    'Returning a new array specifically for the detail invoice
    SalesArray = SQLSelect(DataBaseName, SQL)
    
    RowStart = 9
    RowNum = UBound(SalesArray, 2)
    
    'Pasting this data onto the sheet and putting the customer name and date at the top
    With Sht_Detail
        .Range("A" & RowStart & ":H" & (RowStart + RowNum)).Value = WorksheetFunction.Transpose(SalesArray)
        .Range("D4").Value = Customer
        .Range("H6").Formula = "=TODAY()"
        .Range("B6").Value = Year & " " & Quarter
    
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        'Inputting the subtotals on the invoice detail sheet via a loop
        CurrentRow = LastRow + 1
        
        For Y = 2 To 5
            CurrentString = CurrentString + .Cells(LastRow, Y).Value
        Next Y
        
        For R = LastRow - 1 To RowStart - 1 Step -1
        
                NewString = vbNullString
                
                For Y = 2 To 5
                    NewString = NewString + .Cells(R, Y).Value
                Next Y
                
                'Means we've found a new combination
                If NewString <> CurrentString Then
                    
                    'Inputting our grand total
                    .Cells(CurrentRow, 1).Value = "Grand Total"
                    .Range("G" & CurrentRow).Formula = "=SUM(G" & (R + 1) & ":G" & (CurrentRow - 1) & ")"
                    .Range("H" & CurrentRow).Formula = "=SUM(H" & (R + 1) & ":H" & (CurrentRow - 1) & ")"
                    .Range("A" & CurrentRow & ":H" & CurrentRow).Font.Bold = True
                    
                    If R <> RowStart - 1 Then
                        'Inserting a new row
                        .Range("A" & R).EntireRow.Offset(1).Resize(1).Insert Shift:=xlDown
                        
                        'Resetting our current row
                        CurrentRow = R + 1
                        
                        'Resetting the value of currentstring
                        CurrentString = vbNullString
                            
                            For Y = 2 To 5
                                CurrentString = CurrentString + .Cells(R, Y).Value
                            Next Y
                    End If
                    
                End If
                
        Next R
        
        'Fixing up the formatting
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        .Range("A" & RowStart & ":H" & LastRow).Borders.LineStyle = xlContinuous
        
        'Export to PDF
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=DetailFile, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        'Delete Rows
        .Rows(RowStart & ":" & LastRow).EntireRow.Delete
        
    End With
    
    '==========================================================================================================
    'PREPARING EMAIL
    '==========================================================================================================
    EmailCount = EmailCount + 1
    
    Set outlookApp = New Outlook.Application
    Set myMail = outlookApp.CreateItem(olMailItem)
    
    With myMail
        .To = AddressArray(1, 0)
        .Subject = Customer & " " & Category & " " & Year & " " & Quarter & " Invoices"
        .Body = "Hi," & vbNewLine & "Here is your invoice." & vbNewLine & "Thank You"
        .Attachments.Add SummaryFile
        .Attachments.Add DetailFile
        .Save
    End With
    
Next Customer

    MsgBox "You have " & EmailCount & " email(s) in your drafts folder."

End Sub
