VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rooms"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmRoom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMoreRoomInfo 
      Height          =   540
      Left            =   4410
      Picture         =   "frmRoom.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5100
      Width           =   615
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   7
      Top             =   5100
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5100
      Width           =   2535
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Selected Room"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   6
      Top             =   4380
      Width           =   2895
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "txtComments"
      Top             =   4620
      Width           =   3195
   End
   Begin VB.TextBox txtRoomNumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "txtRoomNumber"
      Top             =   4620
      Width           =   1515
   End
   Begin VB.ComboBox cboPrice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3660
      Width           =   3195
   End
   Begin VB.CommandButton cmdSetPrice 
      Caption         =   "&Set Price for Selected Rooms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   5
      ToolTipText     =   "Set the price for the selected rooms"
      Top             =   3660
      Width           =   2895
   End
   Begin MSComctlLib.ListView lvwRoom 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Price:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lblComments 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblRoom 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4290
      Width           =   1515
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' A form added to allow the user to conveniently add and remove
' rooms from the database
'
Option Explicit

Dim m_DB As DAO.Database                    ' Holds the database we're currently accessing (Hotel.mdb) - used throughout the code in this form
Dim m_wrkJet As DAO.Workspace               ' The workspace of the database we're accessing
Dim m_RS As DAO.Recordset                   ' The recordset we use throughout the form to read/write data from the database

' *************************************************************
' Two possibilities:
' 1. If the caption for cmdAdd is "Add New Room", prepare the form for adding a new customer
' 2. If the caption for cmdAdd is "Save New Room", commit the changes to the database
'     and update the listview
'
Private Sub cmdAdd_Click()
    Dim s As String
    Dim li As ListItem
    Dim Prices() As String
    
    On Error GoTo Problem
       
    If cmdAdd.Caption = "&Add New Room" Then
        ' Preparing to add new room - make the textboxes used to type in room information visible to the user
        m_RS.AddNew
        ClearTextboxes
        txtRoomNumber.Visible = True
        txtComments.Visible = True
        lblRoom.Visible = True
        lblComments.Visible = True
        cmdAdd.Caption = "&Save New Room"
    Else
        ' We're saving the data the user just typed in
        Prices = Split(cboPrice.Text, Chr(151))         ' Get the price from the cboPrice combo box
        s = txtRoomNumber.Text
        If s = "" Then                                  ' Make sure we have a valid room number (can be a string)
            MsgBox "Please enter a Room number", vbExclamation, "Room Number missing"
            Exit Sub
        End If
        
        ' Save the rest of info
        m_RS!RoomNumber = s
        m_RS!PriceID = cboPrice.ItemData(cboPrice.ListIndex)
        If txtComments.Text <> "" Then m_RS!Comments = txtComments.Text
        m_RS.Update
        
        ' Now add this new info to the listview
        Set li = lvwRoom.ListItems.Add
        li.Text = s                                     ' Room number
        li.SubItems(1) = Trim(Prices(0))                ' Price
        li.SubItems(2) = Trim(Prices(1))                ' Price Category
        li.SubItems(3) = txtComments.Text               ' Comments
        
        ' Let the user know we had success, and update the form visual cues
        MsgBox "Room " & s & " has been successfully added", vbInformation, "Addition Successful"
        cmdRemove.Enabled = True
        txtRoomNumber.Visible = False
        txtComments.Visible = False
        lblRoom.Visible = False
        lblComments.Visible = False
        cmdAdd.Caption = "&Add New Room"
    End If
    Exit Sub
    
Problem:
    ' Provide a nice error message if the room already exists
    If Err.Number = -2147467259 Then
        MsgBox "Sorry, " & s & " already exists in the database. " & vbNewLine & _
                "Please enter a different room number.", vbExclamation, "Problem"
    Else
        MsgBox "Unhandled error:" & vbNewLine & _
            "Error #" & Err.Number & vbNewLine & _
            "Error Description:" & Err.Description & vbNewLine & _
            "Source: " & Err.Source
    End If
    
End Sub

' *************************************************************
' All done -- unload the form
'
Private Sub cmdDone_Click()
    Unload Me
End Sub

' *************************************************************
' Display frmBilling to provide a complete listing of everyone
' who has ever stayed in this room
'
Private Sub cmdMoreRoomInfo_Click()
    If lvwRoom.SelectedItem Is Nothing Then Exit Sub    ' Ensure we have valid data
    
    m_RoomNumber = lvwRoom.SelectedItem.Text            ' Save the current room number to a public
                                                        '   variable for later use in frmBilling for Report filtering.
    ReportChoice = vbGroupRoom                          ' Set the option that we want the report to show Records for this room
    m_StartDate = "12:00:00 AM"                         ' We want to show all visits regardless of date that this customer has been at the
    CustomerName = ""                                   ' Since we want to know the stays in this room for ALL customers, reset customername to ""
    Screen.MousePointer = vbHourglass
    frmBilling.Show vbModal                             ' Show the form
End Sub

' *************************************************************
' Remove the customer from the database and the listview
'
Private Sub cmdRemove_Click()
    Dim RS As DAO.Recordset
    Dim SQL As String
    Dim li As ListItem
            
    On Error GoTo Problem
    Set li = lvwRoom.SelectedItem
       
    ' Open the about-to-be removed record from the database
    SQL = "Select * from tblRoom where RoomNumber = '" & li.Text & "'"
    Set RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
    
    ' Delete the room from the database and the listview
    If RS.RecordCount > 0 Then
        RS.Delete
        lvwRoom.ListItems.Remove li.index
    End If
    If lvwRoom.ListItems.Count = 0 Then cmdRemove.Enabled = False
    Exit Sub

Problem:
    ' Give the user a nicer message if this room is used in a reservation
    If Err.Number = 3200 Then
        MsgBox "Sorry, a reservation exists for " & li.Text & ". " & vbNewLine & _
                "Please remove the reservation before deleting the room.", vbExclamation, "Problem"
    Else
        MsgBox "Unhandled error:" & vbNewLine & _
            "Error #" & Err.Number & vbNewLine & _
            "Error Description:" & Err.Description & vbNewLine & _
            "Source: " & Err.Source
    End If
End Sub

' *************************************************************
' Add visual enhancements to the form
'
Private Sub Form_Activate()
    Dither Me                           ' Add the teal shading
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Initialize the form controls and database
'
Private Sub Form_Load()
    InitDatabase                        ' Initialize database
    InitPrices                          ' Get price categories from DB
    InitlvwRoom                         ' Initialize listview
    
    ' Set the form to the "Not adding a new room" state
    ClearTextboxes
    txtRoomNumber.Visible = False
    txtComments.Visible = False
    lblRoom.Visible = False
    lblComments.Visible = False
    cboPrice.Visible = False
    
    If m_RS.RecordCount = 0 Then cmdRemove.Enabled = False
    lvwRoom_ItemClick lvwRoom.SelectedItem
End Sub

' *************************************************************
Private Sub ClearTextboxes()
    txtRoomNumber.Text = ""
    txtComments.Text = ""
End Sub

' *************************************************************
' Open a handle to the database and the table we want to work with
'
Private Sub InitDatabase()
    Set m_wrkJet = CreateWorkspace("", "admin", "", dbUseJet)       ' Create Microsoft Jet Workspace object.
    Set m_DB = m_wrkJet.OpenDatabase(App.Path & "\Hotel.mdb")       ' Open Database object from saved Microsoft Jet database (Hotel.mdb)
    Set m_RS = m_DB.OpenRecordset("tblRoom", dbOpenDynaset)         ' All our modifications are from tblRoom -- open this table now.
End Sub

' *************************************************************
' Get the price categories from the database
'
Private Sub InitPrices()
    Dim RS As DAO.Recordset
            
    Set RS = m_DB.OpenRecordset("tblPrice", dbOpenSnapshot)         ' Open the Price table
    cboPrice.Clear                                                  ' Reset the cboPrice combobox
    If Not (RS.BOF And RS.EOF) Then RS.MoveFirst
    
    ' Add an item to cboPrice for each price category
    While Not RS.EOF
        cboPrice.AddItem Format(RS!Price, "Currency") & " " & Chr(151) & " " & RS!PriceName
        cboPrice.ItemData(cboPrice.NewIndex) = RS!PriceID
        RS.MoveNext
    Wend
    If cboPrice.ListCount > 0 Then cboPrice.ListIndex = 0
    Set RS = Nothing
End Sub

' *************************************************************
' Initialize lvwRoom by adding four columns and displaying the info
' about the already existing rooms in lvwRoom
'
Private Sub InitlvwRoom()
    Dim lvw As ListView
    Dim RS As DAO.Recordset
    Dim li As ListItem
    
    ' Add the columns using the AddColumn helper function (in Utilities.bas)
    Set lvw = lvwRoom
    AddColumn lvw, "Room", 1000
    AddColumn lvw, "Price", 1000
    AddColumn lvw, "Price Category", 1800
    AddColumn lvw, "Room Comments", 3950
    lvw.ListItems.Clear
    
    ' Open the stored query that joins the rooms with their corresponding prices
    ' and add this information to the listview
    Set RS = m_DB.OpenRecordset("qryRoomPrices", dbOpenSnapshot)
    If Not (RS.BOF And RS.EOF) Then RS.MoveFirst
    While Not RS.EOF
        Set li = lvw.ListItems.Add
        li.Text = RS!RoomNumber
        li.SubItems(1) = Format(RS!Price, "Currency")
        li.SubItems(2) = RS!PriceName
        If Not IsNull(RS!Comments) Then li.SubItems(3) = RS!Comments
        li.Tag = RS!PriceID
        RS.MoveNext
    Wend
    Set RS = Nothing
End Sub

' *************************************************************
' The user has the ability to highlight multiple lines in lvwRooom
' and set the price for all of them simultaneously by clicking on
' cmdSetPrice
'
Private Sub cmdSetPrice_Click()
    Dim li As ListItem
    Dim RS As DAO.Recordset
    Dim SQL As String
    Dim Prices() As String
      
    ' Get the currently selected price
    Prices = Split(cboPrice.Text, Chr(151))
    
    ' Save the new price info in the database and listview
    For Each li In lvwRoom.ListItems
        If li.Selected Then
            ' Update the database
            SQL = "Select * From tblRoom Where RoomNumber = '" & li.Text & "'"
            Set RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
            RS.Edit
            RS!PriceID = cboPrice.ItemData(cboPrice.ListIndex)
            RS.Update
            ' Update the listvew
            li.SubItems(1) = Trim(Prices(0))        ' The new Price category
            li.SubItems(2) = Trim(Prices(1))        ' The new Price
        End If
    Next
    
    ' Cleanup
    Set li = Nothing
    Set RS = Nothing
End Sub

' *************************************************************
' A housekeeping function for lvwRoom -- ensures that the user
' can't click on remove if nothing is selected in the listview
'
Private Sub lvwRoom_Click()
    Dim li As ListItem
    Dim NumberSelected As Long
               
    ' Count how many listitems are highlighted
    NumberSelected = 0
    For Each li In lvwRoom.ListItems
        If li.Selected Then NumberSelected = NumberSelected + 1
    Next
    If NumberSelected = 0 Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
    ClearTextboxes
End Sub

' *************************************************************
' Provided as a convenience - allows the user to click on the
' column headers to sort the listview.
'
Private Sub lvwRoom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRoom.SortKey = ColumnHeader.index - 1
    If (lvwRoom.SortOrder = lvwAscending) Then
        lvwRoom.SortOrder = lvwDescending
    Else
        lvwRoom.SortOrder = lvwAscending
    End If
End Sub

' *************************************************************
' Show the billing form for this particular room
'
Private Sub lvwRoom_DblClick()
    cmdMoreRoomInfo_Click
End Sub

' *************************************************************
' When an item is clicked on, make the combo box reflect the price
' of the currently selected room
'
Private Sub lvwRoom_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer                            ' Counter variable
    
    cboPrice.Visible = True
    
    If Item Is Nothing Then Exit Sub
    ' Cycle through the combobox and find the price to select
    For i = 0 To cboPrice.ListCount - 1
        If cboPrice.ItemData(i) = Item.Tag Then
            cboPrice.ListIndex = i
            Exit For
        End If
    Next i
End Sub

