VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Main_Window 
   Caption         =   "Invoice Generator"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   OleObjectBlob   =   "Main_Window.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Main_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This public boolean variable is to stop the updateissues sub being run for each item in the customer listbox when the all or none buttons are being clicked
'Each added item is counted as an update
'This is turned on and off when the none or all buttons are clicked
Public SuppressUpdate As Boolean

'These are public variables that store lists of the issues for rates and addresses
'They are updated when the update issues sub is run and are used to display the issues to users
Public ArrayRatesIssues As Variant
Public ArrayAddressIssues As Variant

'===========================================================================================================
'FUNCTIONS
'===========================================================================================================

Private Function CustomerStringList() As String
'Used to return a string of the current customer selection
'Will be used in conjunction with 'IN' SQL queries

Dim I As Integer
Dim Start As Integer
    
    Start = 0
    CustomerStringList = "("
    
    For I = 0 To CustomersListBox.ListCount - 1
    
        If CustomersListBox.Selected(I) = True Then
            
            If Start > 0 Then CustomerStringList = CustomerStringList & ", "
            
            CustomerStringList = CustomerStringList & "'" & CustomersListBox.List(I) & "'"
            
            Start = Start + 1
            
        End If
        
    Next I
    
    CustomerStringList = CustomerStringList + ")"
    

End Function

Private Function CustomerArray() As Variant
'Returns an array made of strings of customers currently selected
'If no customers have been selected then it will return an array with a single null string

Dim I As Integer
Dim ArrayList() As String
Dim ArrayCount As Long
ArrayCount = 0
    
    For I = 0 To CustomersListBox.ListCount - 1
        If CustomersListBox.Selected(I) = True Then
            ReDim Preserve ArrayList(ArrayCount)
            ArrayList(ArrayCount) = CustomersListBox.List(I)
            ArrayCount = ArrayCount + 1
        End If
    Next I
   
If ArrayCount = 0 Then
    CustomerArray = Array("")
Else
    CustomerArray = ArrayList
End If

End Function

'===========================================================================================================
'PROCEDURES
'===========================================================================================================

Private Sub Update_Issues()
'Updates both issues count labels

Me.AddressesIssuesLabel.Caption = "0"
Me.RatesIssuesLabel.Caption = "0"
Me.PopUpText.Caption = ""

Dim StringList As String
StringList = CustomerStringList()

If StringList <> "()" Then
    Call Load_Address_Issues(StringList)
    Call Load_Rates_Issues(StringList)
Else
    'This is to clear out these arrays if nothing has been selected
    ReDim ArrayAddressIssues(0)
    ReDim ArrayRatesIssues(0)
End If

End Sub

Private Sub ComboBox_Populate()
'Loads in the default values for the three comboboxes
'The year box has the current year with 5 years either side
'The quarter box has 4 quarters (obviously)
'The category box runs a SQL query for Distinct sub account types in the main database

    Dim R As Long
    Dim CategoryArray As Variant
    Dim Element As Variant
    Dim SQLQuery As String
    
    For R = 1 To 4
        QuarterCB.AddItem "Q" & R
    Next R
    
    For R = Year(Date) - 5 To Year(Date) + 5
        YearCB.AddItem R
    Next R
    
    'Getting the category values
    SQLQuery = "SELECT DISTINCT ? FROM ?;"
    
    CategoryArray = SQLSelect(DataBaseName, SQLQuery, Split(SalesDataFields, ",")(3), SalesDataTable)
    
    For Each Element In CategoryArray
        CategoryCB.AddItem Element
    Next Element
    
End Sub

Private Sub Load_Address_Issues(StringList As String)
'To find how many customers do not have a physical address or an email address in the system
'StringList should be a list in SQL format of the current customer selection

Dim AddressArray As Variant
Dim CountryCode As Variant
Dim CountryCodeArray As Variant
Dim R As Long
Dim RequiredNums As Long
Dim NumCount As Long
Dim ArrayCount As Long
Dim FoundMatch As Boolean

'Returns a distinct array of the country codes for the selected customers
CountryCodeArray = SQLSelect(DataBaseName, "SELECT DISTINCT ? FROM ? WHERE ? IN ?;", _
                                            Split(SalesDataFields, ",")(17), SalesDataTable, Split(SalesDataFields, ",")(18), StringList)
                                            
'Returns the entire address table
AddressArray = SQLSelect(DataBaseName, "SELECT * FROM ?;", AddressDataTable)

'The number of correct entries we need to achieve
RequiredNums = UBound(CountryCodeArray, 2) + 1
NumCount = 0

'Redimming the array we're going to use to store a list of country codes that don't have an address
ReDim ArrayAddressIssues(0)

'Looping through both arrays to check that each customer has an address
For Each CountryCode In CountryCodeArray
    
    FoundMatch = False
    
    For R = 0 To UBound(AddressArray, 2)
        
        If CountryCode = AddressArray(1, R) Then
            
            If AddressArray(4, R) <> vbNullString And AddressArray(2, R) <> vbNullString Then
                NumCount = NumCount + 1
                FoundMatch = True
                Exit For
            End If
            
        End If
        
    Next R
    
    'Storing problematic items in our list
    If FoundMatch = False Then
        ReDim Preserve ArrayAddressIssues(ArrayCount)
        ArrayAddressIssues(ArrayCount) = CountryCode
        ArrayCount = ArrayCount + 1
    End If
    
Next CountryCode

'Setting the value of the number of errors
Me.AddressesIssuesLabel.Caption = CStr(RequiredNums - NumCount)

End Sub

Private Sub Load_Rates_Issues(StringList As String)
'To find how many selected line items do not have rates available in the database
'StringList should be a list in SQL format of the current customer selection

Dim RatesTableArray As Variant
Dim SalesRefArray As Variant
Dim StringArray() As String
Dim RequiredNums As Long
Dim NumCount As Long
Dim SQL As String
Dim MyString As String
Dim R As Long
Dim Y As Byte
Dim Rate As Variant
Dim Selection As Variant
Dim FoundMatch As Boolean
Dim ArrayCount As Long

Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String

'The five fields that make up the rates reference
A = Split(SalesDataFields, ",")(4)
B = Split(SalesDataFields, ",")(16)
C = Split(SalesDataFields, ",")(15)
D = Split(SalesDataFields, ",")(12)
E = Split(SalesDataFields, ",")(14)

'Returns a list of the unique rates references for customers selected, for category selected
      SQL = SQLString("SELECT DISTINCT ?, ?, ?, ?, ? ", A, B, C, D, E)
SQL = SQL + SQLString("FROM ? ", SalesDataTable)
SQL = SQL + SQLString("WHERE ? IN ? ", Split(SalesDataFields, ",")(18), StringList)
SQL = SQL + SQLString("AND ? = '?';", Split(SalesDataFields, ",")(3), Me.CategoryCB.Value)

SalesRefArray = SQLSelect(DataBaseName, SQL)

'Setting the new length of our Stringarray and loading in each unique reference code from selection
ReDim StringArray(0 To UBound(SalesRefArray, 2))

For R = 0 To UBound(SalesRefArray, 2)

    MyString = ""
    
    For Y = 0 To 4
        MyString = MyString + SalesRefArray(Y, R)
    Next Y
    
    StringArray(R) = MyString
    
Next R

'Returns a list of reference codes from the rates master table
RatesTableArray = SQLSelect(DataBaseName, "SELECT DISTINCT ? FROM ?;", Split(RatesDataFields, ",")(2), RatesDataTable)

'The number of correct entries we need to achieve
RequiredNums = UBound(StringArray) + 1
NumCount = 0

'Looping through both arrays to check that each rate has a match
'Also storing the rates with no match in our public variable
ReDim ArrayRatesIssues(0)

For Each Selection In StringArray
    
    FoundMatch = False
    
    For Each Rate In RatesTableArray
        
        If Selection = Rate Then
            NumCount = NumCount + 1
            FoundMatch = True
            Exit For
        End If
        
    Next Rate
    
    'Storing problematic items in our list
    If FoundMatch = False Then
        
        ReDim Preserve ArrayRatesIssues(ArrayCount)
        ArrayRatesIssues(ArrayCount) = Selection
        ArrayCount = ArrayCount + 1
        
    End If
    
Next Selection

'Setting the value of the number of errors
Me.RatesIssuesLabel.Caption = CStr(RequiredNums - NumCount)

End Sub

Private Sub Update_Customer_Table()
'Will update the customer listbox in accordance with the combobox selection if they are all not blank
'Runs a SQL query on the database to get the list

    Me.AddressesIssuesLabel.Caption = "0"
    Me.RatesIssuesLabel.Caption = "0"
    Me.PopUpText.Caption = ""
    
    If CategoryCB.Value <> vbNullString And _
       QuarterCB.Value <> vbNullString And _
       YearCB.Value <> vbNullString Then
        
        'Update Customer Table
        Dim Dict As Scripting.Dictionary
        Set Dict = New Scripting.Dictionary
        Dim Q As Byte
        Dim Month As Variant
        Dim SQL As String
        Dim Element As Variant
        Dim CustomerList As Variant
        
        'Dictionary for the four quarters and corresponding months
        For Q = 1 To 4
                Dict.Add "Q" & Q, Q
        Next Q
        
        Month = Dict(QuarterCB.Value)
        
        'Constructing our SQL Query
        SQL = "SELECT DISTINCT ? " & _
              "FROM ? " & _
              "WHERE ? = '?' AND ? = '?' AND ? = ?;"
            
        CustomerList = SQLSelect(DataBaseName, SQL, _
                        Split(SalesDataFields, ",")(18), _
                        SalesDataTable, _
                        Split(SalesDataFields, ",")(4), _
                        YearCB.Value, _
                        Split(SalesDataFields, ",")(3), _
                        CategoryCB.Value, _
                        Split(SalesDataFields, ",")(7), _
                        Month)
        
        CustomersListBox.Clear
        
        For Each Element In CustomerList
            CustomersListBox.AddItem Element
        Next Element
        
        
    End If
    

End Sub

Private Sub Caption_Output_Folder()
'Will set the value of the output folder if the previous user is the same

    If ReturnPreviousUser() = UserName() Then
        OutputFolderLabel.Caption = ReturnOutputFolder()
    Else
        OutputFolderLabel.Caption = "None"
    End If
        
End Sub

Private Sub Enable_Run_Button()
'If critera is met, the run button will be enabled

RunButton.Enabled = False

If CategoryCB.Value <> vbNullString And _
       QuarterCB.Value <> vbNullString And _
       YearCB.Value <> vbNullString And _
       CustomerArray(0) <> vbNullString Then
       
       RunButton.Enabled = True

End If

End Sub

'===========================================================================================================
'EVENTS
'===========================================================================================================

Private Sub UserForm_Initialize()
    Call ComboBox_Populate
    Call Caption_Output_Folder
    Call Update_Issues
End Sub

Private Sub btnBrowseOutputFolder_Click()

    Dim Folder As String
    
    Folder = BrowseForFolder("Select where you would like to save your files:")
    
    If Folder <> vbNullString Then
        OutputFolderLabel.Caption = Folder
        Call Set_Output_Folder(Folder)
        Call Set_Current_User
    End If

End Sub

Private Sub CategoryCB_Change()

    Call Update_Customer_Table

End Sub

Private Sub YearCB_Change()

    Call Update_Customer_Table

End Sub

Private Sub QuarterCB_Change()
    
    Call Update_Customer_Table
    
End Sub

Private Sub CustomersListBox_Change()

If Me.SuppressUpdate = False Then
    Call Update_Issues
    Call Enable_Run_Button
End If

End Sub

Private Sub SelectAllButton_Click()
'Turns the suppress update button on and off to stop the updateissues sub being run

    Me.SuppressUpdate = True
    
    Dim I As Integer
    
    For I = 0 To CustomersListBox.ListCount - 1
        CustomersListBox.Selected(I) = True
    Next I
    
    Call Update_Issues
    
    Me.SuppressUpdate = False
    
    Call Enable_Run_Button

End Sub

Private Sub SelectNoneButton_Click()
'Turns the suppress update button on and off to stop the updateissues sub being run

    Me.SuppressUpdate = True
    
    Dim I As Integer
    
    For I = 0 To CustomersListBox.ListCount - 1
        CustomersListBox.Selected(I) = False
    Next I
    
    Me.AddressesIssuesLabel.Caption = "0"
    Me.RatesIssuesLabel.Caption = "0"
    Me.RunButton.Enabled = False
    Me.PopUpText.Caption = ""
    
    Me.SuppressUpdate = False
    
End Sub

Private Sub RunButton_Click()
'The big dance. Actually creating the invoices

If Me.AddressesIssuesLabel.Caption <> "0" Or Me.RatesIssuesLabel.Caption <> "0" Then
    
    Me.PopUpText.Caption = "Fix Issues"

Else
       'Passing in all the required arguments
       Call Create_Invoices(CustomerArray(), Me.CategoryCB.Value, Me.OutputFolderLabel.Caption, Me.YearCB.Value, Me.QuarterCB.Value)
       Unload Me
       
End If

End Sub

Private Sub btnBrowseOutputFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnBrowseOutputFolder.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btnBrowseOutputFolder_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnBrowseOutputFolder.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub btnRatesView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnRatesView.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btnRatesView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnRatesView.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub btnAddressesView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnAddressesView.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub btnAddressesView_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.btnAddressesView.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub btnAddressesView_Click()
'The issues form doubles up as being used for both addresses and rates
'Here we load our address issues
    
    With Issues_Form
    
        .ListBox1.Clear
        .TitleLabel.Caption = "Country Codes With Missing Address"
        .IssueNumber.Caption = "0"
        
        If Not IsEmpty(ArrayAddressIssues(0)) Then
            .ListBox1.List = ArrayAddressIssues
            .IssueNumber.Caption = CStr(.ListBox1.ListCount)
        End If
        
        .Show
        
    End With
    
End Sub

Private Sub btnRatesView_Click()
'Loading rates issues
    
    With Issues_Form
    
        .ListBox1.Clear
        .TitleLabel.Caption = "Rates Codes Missing From Rates Table"
        .IssueNumber.Caption = "0"
        
        If Not IsEmpty(ArrayRatesIssues(0)) Then
            .ListBox1.List = ArrayRatesIssues
            .IssueNumber.Caption = CStr(.ListBox1.ListCount)
        End If
        
        .Show
        
    End With
    
End Sub
