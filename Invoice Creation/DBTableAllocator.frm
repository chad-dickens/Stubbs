VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DBTableAllocator 
   Caption         =   "Select Tables"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "DBTableAllocator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DBTableAllocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===========================================================================================================
'PROCEDURES
'===========================================================================================================

Private Sub Load_List_Box()
'Loads ListBox with all the tables in the DataBase the User has selected

Dim cnnDB As ADODB.Connection
Dim rstList As ADODB.Recordset
Dim DataBaseTables As Variant
Dim Element As Variant

Set cnnDB = New ADODB.Connection

' Open the connection.
cnnDB.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" & DataBaseName & ";"

' Open the tables schema rowset.
Set rstList = cnnDB.OpenSchema(adSchemaTables)
DataBaseTables = rstList.GetRows

'Checking if the tables exist

For Each Element In DataBaseTables
        If Element <> vbNullString Then
            ListBox1.AddItem Element
        End If
Next Element

rstList.Close
cnnDB.Close

End Sub

Private Sub If_Check_Buttons_Disabled()
'Once all three buttons have been clicked, the form will be unloaded

    If SalesButton.Enabled = False And _
       RatesButton.Enabled = False And _
       AddressButton.Enabled = False Then
       
       Unload Me
    End If

End Sub

'===========================================================================================================
'EVENTS
'===========================================================================================================

Private Sub UserForm_Initialize()
    ExitButton = False
    Call Load_List_Box
End Sub

'The following three subs let users select which tables they want to use

Private Sub AddressButton_Click()
    
    If ListBox1.Value <> vbNullString Then
        AddressDataTable = ListBox1.Value
        ListBox1.RemoveItem ListBox1.ListIndex
        AddressButton.Enabled = False
    End If
    
    Call If_Check_Buttons_Disabled
    
End Sub

Private Sub RatesButton_Click()

    If ListBox1.Value <> vbNullString Then
        RatesDataTable = ListBox1.Value
        ListBox1.RemoveItem ListBox1.ListIndex
        RatesButton.Enabled = False
    End If
    
    Call If_Check_Buttons_Disabled

End Sub

Private Sub SalesButton_Click()
    
    If ListBox1.Value <> vbNullString Then
        SalesDataTable = ListBox1.Value
        ListBox1.RemoveItem ListBox1.ListIndex
        SalesButton.Enabled = False
    End If
    
    Call If_Check_Buttons_Disabled
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then ExitButton = True

End Sub

Private Sub QuitButton_Click()
    ExitButton = True
    Unload Me
End Sub

