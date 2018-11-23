VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7650
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMoreCustomerInfo 
      Height          =   360
      Left            =   810
      Picture         =   "frmCustomer.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   180
      Width           =   480
   End
   Begin VB.TextBox txtHomeCountry 
      DataField       =   "HomeCountry"
      DataSource      =   "adoCustomer"
      Height          =   405
      Left            =   1500
      TabIndex        =   5
      Top             =   2973
      Width           =   3255
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataSource      =   "adoCustomer"
      Height          =   405
      Left            =   1500
      TabIndex        =   2
      Top             =   1362
      Width           =   3255
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataSource      =   "adoCustomer"
      Height          =   405
      Left            =   1500
      TabIndex        =   1
      Top             =   825
      Width           =   3255
   End
   Begin VB.TextBox txtPhoneNumber 
      DataField       =   "PhoneNumber"
      DataSource      =   "adoCustomer"
      Height          =   405
      Left            =   1500
      TabIndex        =   3
      Top             =   1899
      Width           =   3255
   End
   Begin VB.TextBox txtHomeState 
      DataField       =   "HomeState"
      DataSource      =   "adoCustomer"
      Height          =   405
      Left            =   1500
      TabIndex        =   4
      Top             =   2436
      Width           =   3255
   End
   Begin VB.TextBox txtComments 
      DataField       =   "Comments"
      DataSource      =   "adoCustomer"
      Height          =   1005
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3510
      Width           =   3255
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   5025
      TabIndex        =   10
      Top             =   3405
      Width           =   2400
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   495
      Left            =   5025
      TabIndex        =   9
      Top             =   2679
      Width           =   2400
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove this Customer"
      Height          =   495
      Left            =   5025
      TabIndex        =   8
      Top             =   1956
      Width           =   2400
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New Customer"
      Height          =   495
      Left            =   5025
      TabIndex        =   7
      Top             =   1233
      Width           =   2400
   End
   Begin VB.ComboBox cboCustomerName 
      Height          =   360
      Left            =   1500
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc adoCustomer 
      Height          =   375
      Left            =   135
      Top             =   4140
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoCustomer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   825
      Left            =   5025
      Picture         =   "frmCustomer.frx":0454
      Top             =   180
      Width           =   2400
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Country:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   105
      TabIndex        =   16
      Top             =   3015
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home State:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   105
      TabIndex        =   15
      Top             =   2475
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   105
      TabIndex        =   14
      Top             =   1950
      Width           =   1350
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   5
      Left            =   105
      TabIndex        =   13
      Top             =   3570
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   105
      TabIndex        =   12
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   105
      TabIndex        =   11
      Top             =   1425
      Width           =   1020
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: A form added to allow the user to conveniently add and
'           delete customers from the database
'

Option Explicit

' *************************************************************
' When the user chooses a new customer, parse the first and last
' name from the text (in the form <last name>, <first name>) and
' bring up the relevant customer record.  Note that the first and
' last names combined must be unique (ie. there can be only one
' John Doe in the database)
'
Private Sub cboCustomerName_Click()
    Dim s As String                         ' Temporary string variable
    Dim Name() As String                    ' Used to help parse cboCustomerName.text
    
    If (cboCustomerName.Text <> "") Then    ' Don't do anything if the field is somehow blank
        Name = Split(cboCustomerName.Text, ",")     ' Parse into first and last name
        ' Build the SQL query
        s = "Select * from tblCustomer " & _
            "Where LastName = '" & Trim(Name(0)) & "'" & _
            " and FirstName = '" & Trim(Name(1)) & "'"
        adoCustomer.RecordSource = s
        adoCustomer.Refresh
    End If
End Sub

' *************************************************************
' When the user chooses to add a new customer, initialize the record
' source.
'
Private Sub cmdAdd_Click()
    adoCustomer.Recordset.CancelUpdate
    adoCustomer.Recordset.AddNew
    cmdAdd.Enabled = False
End Sub

' *************************************************************
' All done - exit the form
'
Private Sub cmdDone_Click()
    Unload Me
End Sub

' *************************************************************
' Display frmBilling to provide a complete listing of this customer's
' activity at the hotel
'
Private Sub cmdMoreCustomerInfo_Click()
    CustomerName = Trim(cboCustomerName.Text)       ' Save the current customer's name to a public
                                                    '      variable for later use in frmBilling for Report filtering.
    ReportChoice = vbGroupName                      ' Set the option that we want the report to show Records for this customer
    m_RoomNumber = ""                               ' Since we want to know the customer's stays in ALL rooms, reset room number to ""
    m_StartDate = "12:00:00 AM"                     ' We want to show all visits regardless of date that this customer has been at the
                                                    '      hotel -- therefore set the starttime to the default value
    Screen.MousePointer = vbHourglass
    frmBilling.Show vbModal                         ' Show the form
End Sub

' *************************************************************
' Delete the currently selected customer from the database
'
Private Sub cmdRemove_Click()
    Dim ct As Control                               ' Used for cycling through all the controls on the form

    On Error GoTo Problem
    
    If adoCustomer.Recordset.RecordCount > 0 Then   ' Delete the customer
        adoCustomer.Recordset.Delete
        adoCustomer.Recordset.Requery
        For Each ct In Me.Controls                  ' A way of clearing all the textboxes on the form with fewer lines of code.
            If TypeOf ct Is TextBox Then
                ct.Text = ""
            End If
        Next
        With cboCustomerName                        ' Also emove the customer name from the combo list
            .RemoveItem .ListIndex
            If .ListCount > 0 Then .ListIndex = 0
        End With
    End If
    
    ' If this was the only remaining customer, don't allow any more removing
    If adoCustomer.Recordset.RecordCount = 0 Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
    cmdSave.Enabled = False
    Exit Sub

Problem:
    ' Provide a nice error message if the user tried to remove a price category that a room is currently using (table relationship error)
    If Err.Number = -2147217887 Then
        MsgBox "Sorry, a reservation exists for " & cboCustomerName.Text & "." & vbNewLine & _
                "You can't remove this person while he or she has a reservation.", vbExclamation, "Unable to comply"
    End If
End Sub

' *************************************************************
' Save the newly added or recently edited customer data to the database.
' Note that for new users, the user must have clicked on "Add New Customer" or
' else the new customer info won't be added
'
Private Sub cmdSave_Click()
    Dim i As Integer                        ' counter
    Dim s As String                         ' temporary string holder
    Dim RS As ADODB.Recordset               ' temporary recordset
    
    Set RS = adoCustomer.Recordset
        
    With cboCustomerName
        On Error GoTo Problem
        s = RS!LastName & ", " & RS!FirstName
        If (RS.EditMode = adEditInProgress) Then        ' The user wants to save changes to an existing record
            RS.Update
            MsgBox "Changes succesfully applied to " & s & ".", vbInformation, "Changes Successful"
            .RemoveItem .ListIndex                      ' Temporarily remove the customer name from the combobox
                                                        '     (name re-added with possible name changes below)
        ElseIf (RS.EditMode = adEditAdd) Then           ' The user has clicked on "Add New Customer" and wants to save their changes
            RS.Update
            MsgBox s & " has been successfully added", vbInformation, "Addition Successful"
        End If

        .AddItem s                                      ' Add the new customer name to the customer name navigation combobox
        For i = 0 To .ListCount - 1
            If .List(i) = s Then .ListIndex = i
        Next
        cboCustomerName_Click                           ' Refresh the textboxes
    End With
    Set RS = Nothing
    cmdAdd.Enabled = True                               ' Complete the visual cues on the form
    cmdRemove.Enabled = True
    cmdSave.Enabled = False
    Exit Sub

Problem:
    ' Display a more eye-pleasing error message if the user attempted to add a duplicate customer name
    If Err.Number = -2147467259 Then
        MsgBox "Sorry, " & s & " already exists in the database. " & vbNewLine & _
                "Please enter a different first or last name."
    End If
End Sub

' *************************************************************
' Initialize the form controls
'
Private Sub Form_Load()
    Dither Me                                           ' Add the nice teal shading
    InitAdoRecordsets
    PopulatecboCustomerName
    ' Initialize the form to display the first record in the database (if it exists)
    If (adoCustomer.Recordset.RecordCount <> 0) Then adoCustomer.Recordset.MoveFirst
    If cboCustomerName.ListCount > 0 Then cboCustomerName.ListIndex = 0
    If adoCustomer.Recordset.RecordCount = 0 Then cmdRemove.Enabled = False         ' Don't want to allow deleting if there's no customers in the database
    cmdSave.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Open an ADO connection to Hotel.mdb, which must be in the
' same directory as the application
'
Private Sub InitAdoRecordsets()
    Dim Connection As New Connection
    Dim RS As New ADODB.Recordset
    
    Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read|Write"
    Connection.Open
    
    RS.Open "Select * FROM tblCustomer", Connection, adOpenDynamic, adLockPessimistic, adCmdText
    Set adoCustomer.Recordset = RS
    
    ' Cleanup
    Set Connection = Nothing
    Set RS = Nothing
End Sub
' *************************************************************
' cboCustomerName allows the user to navigate among the names of customers
' The text of the combo box is always in the form <Last Name>, <First Name>
' This function populates the combo box
Private Sub PopulatecboCustomerName()
    Dim RS As New Recordset
    Dim s As String
    
    Set RS = adoCustomer.Recordset

    If (RS.BOF And RS.EOF) Then                         ' Make sure we have a record to use
        Exit Sub
    Else
        RS.MoveFirst
    End If
    
    cboCustomerName.Clear                               ' Clear the combobox of any text that may be inadvertently be in it
    While Not RS.EOF                                    ' Add an item for each customer in the database
        cboCustomerName.AddItem RS!LastName & ", " & RS!FirstName
        RS.MoveNext
    Wend
    Set RS = Nothing
End Sub

' *************************************************************
Private Sub txtComments_Change()
    cmdSave.Enabled = True
End Sub

' *************************************************************
Private Sub txtFirstName_Change()
    cmdSave.Enabled = True
End Sub

' *************************************************************
Private Sub txtHomeCountry_Change()
    cmdSave.Enabled = True
End Sub

' *************************************************************
Private Sub txtHomeState_Change()
    cmdSave.Enabled = True
End Sub

' *************************************************************
Private Sub txtLastName_Change()
    cmdSave.Enabled = True
End Sub

' *************************************************************
Private Sub txtPhoneNumber_Change()
    cmdSave.Enabled = True
End Sub
