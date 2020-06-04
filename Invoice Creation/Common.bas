Attribute VB_Name = "Common"
Option Explicit

'===========================================================================================================
'GENERAL TOOLS
'===========================================================================================================

Public Function UserName() As String
'Returns the first name of the computer user

    If InStr(1, Application.UserName, " ") > 0 Then
        UserName = Split(Application.UserName, " ")(0)
    Else
        UserName = Application.UserName
    End If

End Function

Public Sub Optimize_On()
'Turns off non-essential Excel features to make code run faster
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .StatusBar = "Working..."
End With

End Sub

Public Sub Optimize_Off()
'Turns these essential features back on
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .StatusBar = ""
End With

End Sub

Public Function MoneyFormat(InputValue As Currency) As String
    'Simple function to return a number in currency form
    MoneyFormat = "$" & Format(InputValue, "#,###.##")
    If (Right(MoneyFormat, 1) = ".") Then
        MoneyFormat = Left(MoneyFormat, Len(MoneyFormat) - 1)
    End If
End Function

Public Function Capitalize(MyString As Variant) As String
'Returns string capitalized
    Capitalize = (UCase(Left(MyString, 1)) & LCase(Mid(MyString, 2)))

End Function

Public Function PreviousUserExists() As Boolean
'Checks if a property called previous user exists

    Dim DocProperty As Object

    For Each DocProperty In ThisWorkbook.CustomDocumentProperties
        If DocProperty.Name = "PreviousUser" Then
            PreviousUserExists = True
            Exit Function
        End If
    Next DocProperty
    
    PreviousUserExists = False

End Function

Public Function ReturnPreviousUser() As String
'Returns the previous user name if it exists

    If PreviousUserExists() Then
        ReturnPreviousUser = ThisWorkbook.CustomDocumentProperties("PreviousUser")
    Else
        ReturnPreviousUser = ""
    End If
    
End Function

Public Sub Set_Current_User()
'Sets the PreviousUser property to the current user

        Dim User As String
        User = UserName()

        If PreviousUserExists() Then
            ThisWorkbook.CustomDocumentProperties("PreviousUser") = User
        Else
            ThisWorkbook.CustomDocumentProperties.Add "PreviousUser", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=User
        End If

End Sub

'===========================================================================================================
'FILES AND FOLDERS TOOLS
'===========================================================================================================

Public Function FilePicker(Title As String, FiltersTitle As String, Filters As String) As String
'Simple function that lets user choose a file and returns the file name as a string
'@Title - The text displayed at the top of the filedialog box
'@FiltersTitle - The text displayed on the filters button eg. "Excel Files"
'@Filters - The file types allowed eg. "*.xlsx, *.csv, *.xls, *.xlsm"

Dim vFileD As Office.FileDialog
Set vFileD = Application.FileDialog(msoFileDialogFilePicker)

With vFileD
    .Title = Title
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add FiltersTitle, Filters
    
    If .Show = True Then
        FilePicker = .SelectedItems(1)
    Else
        FilePicker = "Error"
    End If
        
End With

End Function

Public Function BrowseForFolder(Title As String) As String
'Returns The location of a folder selected by the user
'If the user doesn't select anything, it will return a null string

    Dim dlgFolder As FileDialog
    Set dlgFolder = Application.FileDialog(msoFileDialogFolderPicker)
    With dlgFolder
        .Title = Title
        'Display the folder selection dialog box and check return value for Cancel button
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        Else
            BrowseForFolder = ""
        End If
    End With

    Set dlgFolder = Nothing

End Function

Public Function OutputFolderExists() As Boolean
'Returns True is this document already has an output folder

    Dim DocProperty As Object

    For Each DocProperty In ThisWorkbook.CustomDocumentProperties
        If DocProperty.Name = "OutputFolder" Then
            OutputFolderExists = True
            Exit Function
        End If
    Next DocProperty
    
    OutputFolderExists = False

End Function

Public Function ReturnOutputFolder() As String
'Returns the name of the output folder
    ReturnOutputFolder = ThisWorkbook.CustomDocumentProperties("OutputFolder")

End Function

Public Sub Set_Output_Folder(Directory As String)
'Sets the output folder property for this workbook

        If OutputFolderExists() Then
            ThisWorkbook.CustomDocumentProperties("OutputFolder") = Directory
        Else
            ThisWorkbook.CustomDocumentProperties.Add "OutputFolder", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=Directory
        End If

End Sub

'===========================================================================================================
'SQL ACCESS TOOLS
'===========================================================================================================

Public Function SQLSelect(DataBase As String, SQL As String, ParamArray Parameters() As Variant) As Variant
'Function designed to run SQL SELECT queries on Access database
'Don't forget to add a ';' on the end of your SQL query
'Input parameters to replace '?' placeholders

'ADO 2.8 selected in references
Dim rsData As New ADODB.Recordset
Dim sConnect As String
Dim Param As Variant

For Each Param In Parameters
    SQL = Replace(SQL, "?", Param, , 1)
Next

'Creating the connection string
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & DataBase & ";"
        
'Create the recordset object and run the query
rsData.Open SQL, sConnect, adOpenForwardOnly, adLockReadOnly, adCmdText

'Returning an array
If Not rsData.EOF Then
    SQLSelect = rsData.GetRows
Else
    SQLSelect = Array()
End If

'Closing the database connection
rsData.Close

End Function

Public Sub SQL_Command(DataBase As String, SQL As String, ParamArray Parameters() As Variant)
'Procdure designed to run UPDATE, DELETE, INSERT SQL commands for Access Database
'Input parameters to replace '?' placeholders

Dim cn  As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim Param As Variant

For Each Param In Parameters
    SQL = SQLString(SQL, Param)
Next

cn.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" & DataBase & ";"

With cmd
    .ActiveConnection = cn
    .CommandText = SQL
    .CommandType = adCmdText
    .Execute
End With

cn.Close

End Sub

Public Function FieldsExist(DataBase As String, TableName As String, Fields As String) As Boolean
'Simple function to return whether or not a series of fields exist in a database table

Dim cnnDB As ADODB.Connection
Dim rst As New ADODB.Recordset
Dim ii As Integer
Dim ss As String

Set cnnDB = New ADODB.Connection

cnnDB.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" & DataBase & ";"

rst.Open "SELECT * FROM " & TableName, cnnDB, adOpenForwardOnly, adLockReadOnly

For ii = 0 To rst.Fields.Count - 1
    ss = ss & "," & rst.Fields(ii).Name
Next ii

rst.Close
cnnDB.Close

FieldsExist = (ss = Fields)

End Function


Public Function TablesExist(ByVal DefaultTables As Variant, DataBase As String) As Boolean
'Simple function that checks if each string element of an array exists as a table in an Access database

Dim cnnDB As ADODB.Connection
Dim rstList As ADODB.Recordset
Dim DataBaseTables As Variant
Dim Table As Variant
Dim Default As Variant
Dim NumFound As Long

Set cnnDB = New ADODB.Connection

' Open the connection.
cnnDB.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" & DataBase & ";"

' Open the tables schema rowset.
Set rstList = cnnDB.OpenSchema(adSchemaTables)
DataBaseTables = rstList.GetRows

NumFound = -1
TablesExist = False

'Checking if the tables exist
For Each Table In DataBaseTables

        For Each Default In DefaultTables
            
            If Table = Default Then
                NumFound = NumFound + 1
            End If
            
            If NumFound = UBound(DefaultTables) Then
                TablesExist = True
                GoTo CloseDB
            End If
            
        Next Default
    
Next Table

CloseDB:
rstList.Close
cnnDB.Close

End Function

Public Function SQLString(SQL As String, ParamArray Parameters() As Variant) As String
'Used to replace '?' placeholders in sql strings

Dim Param As Variant

For Each Param In Parameters
    
    If InStr(1, Param, ";") > 0 Then
        Err.Raise 316, "SQL Parameter", "SQL Parameters cannot contain ';'"
    End If
    
    SQL = Replace(SQL, "?", Param, , 1)
Next

SQLString = SQL

End Function


