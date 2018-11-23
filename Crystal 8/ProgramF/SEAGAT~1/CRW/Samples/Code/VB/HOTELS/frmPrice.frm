VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrice 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Price Categories"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmPrice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Price Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4245
      TabIndex        =   5
      Top             =   2610
      Width           =   2625
   End
   Begin VB.TextBox txtPrice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   1
      Text            =   "txtPrice"
      Top             =   4275
      Width           =   2505
   End
   Begin VB.TextBox txtPriceName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2730
      TabIndex        =   2
      Text            =   "txtPriceName"
      Top             =   4275
      Width           =   4095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "    &Remove Selected Price                 Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4245
      TabIndex        =   4
      Top             =   1845
      Width           =   2625
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New Price Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4245
      TabIndex        =   3
      Top             =   1095
      Width           =   2625
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
      Height          =   570
      Left            =   4245
      TabIndex        =   6
      Top             =   3360
      Width           =   2625
   End
   Begin MSComctlLib.ListView lvwPrice 
      Height          =   3840
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Image imgLogo 
      Height          =   825
      Left            =   4365
      Picture         =   "frmPrice.frx":030A
      Top             =   90
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Category"
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
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   3975
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Height          =   315
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   3990
      Width           =   975
   End
End
Attribute VB_Name = "frmPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: A form added to allow the user to conveniently add and
'           delete customers from the database
'

Option Explicit

Dim m_DB As DAO.Database                    ' Holds the database we're currently accessing (Hotel.mdb) - used throughout the code in this form
Dim m_wrkJet As DAO.Workspace               ' The workspace of the database we're accessing
Dim m_RS As DAO.Recordset                   ' The recordset we use throughout the form to read/write data from the database

Dim m_SortAsc As Boolean

' *************************************************************
' Prepare the form for adding a new customer
'
Private Sub cmdAdd_Click()
    ClearTextboxes
    m_RS.AddNew
    cmdAdd.Enabled = False
End Sub

' *************************************************************
' All done -- exit the form
'
Private Sub cmdDone_Click()
    Unload Me
End Sub

' *************************************************************
' Remove a customer from the database
'
Private Sub cmdRemove_Click()
    Dim SQL As String
    Dim li As ListItem
    
    ' Check that the currently selected listitem is valid
    If IsError(lvwPrice.SelectedItem.Text) Then
        cmdRemove.Enabled = False
        Exit Sub
    End If
    
    Set li = lvwPrice.SelectedItem
    On Error GoTo Problem
    
    ' Select the correct price category and then delete the corresponding record from the database.
    SQL = "Select * From tblPrice Where PriceID = " & li.Tag
    Set m_RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
    If m_RS.RecordCount = 1 Then
        m_RS.MoveFirst
        m_RS.Delete
        ' Remove the item from the listview
        lvwPrice.ListItems.Remove li.index
    End If
    If lvwPrice.ListItems.Count < 1 Then cmdRemove.Enabled = False
    Exit Sub
    
Problem:
    ' Provide a nice error message if the user tried to remove a price category that a room is currently using (table relationship error)
    If Err.Number = 3200 Then
        MsgBox "Sorry, a room exists that uses the '" & li.Text & " - " & li.SubItems(1) & _
               "' price category. " & vbNewLine & "You can't remove this price category while a room is set to it.", vbExclamation, "Unable to comply"
    Else
        MsgBox "Unhandled error:" & vbNewLine & _
            "Error #" & Err.Number & vbNewLine & _
            "Error Description:" & Err.Description & vbNewLine & _
            "Source: " & Err.Source
    End If
End Sub

' *************************************************************
' Save the price category changes to the database and update the
' listview
'
Private Sub cmdSave_Click()
    Dim SQL As String
    Dim li As ListItem
    
    On Error GoTo Problem
    ' We can only do changes if there is something in the listview to edit
    If lvwPrice.ListItems.Count < 1 And m_RS.EditMode <> dbEditAdd Then
        cmdSave.Enabled = False
        Exit Sub
    End If
    
    ' Do a little extra error checking
    If IsError(CCur(txtPrice.Text)) Then
        MsgBox "Please enter a valid room price.", vbOKOnly, "Unrecognizable Room Price"
        Exit Sub
    End If
    
    If m_RS.EditMode = dbEditAdd Then
        ' If we're adding a new price category, update the database and then update the listview
        m_RS!Price = CCur(txtPrice.Text)
        m_RS!PriceName = txtPriceName.Text
        m_RS.Update
        ' Move to the last record (the most recently added one) and update the listview
        SQL = "Select * From tblPrice Where PriceName = '" & txtPriceName.Text & "'"
        Set m_RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
        m_RS.MoveLast
        Set li = lvwPrice.ListItems.Add
    Else
        ' We're editing an existing record -- update the database and listview
        SQL = "Select * From tblPrice Where PriceID = " & lvwPrice.SelectedItem.Tag
        Set m_RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
        If m_RS.RecordCount = 1 Then
            m_RS.MoveFirst
            m_RS.Edit
            m_RS!Price = CCur(txtPrice.Text)
            m_RS!PriceName = txtPriceName.Text
            m_RS.Update
            Set li = lvwPrice.SelectedItem
        End If
    End If
    
    ' Set the listview to reflect the changes
    li.Text = Format(txtPrice.Text, "Currency")
    li.SubItems(1) = Trim(txtPriceName.Text)
    li.Tag = m_RS!PriceID

Problem:
    If Err.Number <> 0 Then
        MsgBox "Please enter valid data into both the Price and Price Category textboxes", _
                vbInformation, "Correction Needed"
        Exit Sub
    End If
    ClearTextboxes
    InitlvwPrice
    
    ' Update the visual form cues
    cmdSave.Enabled = False
    cmdAdd.Enabled = True
End Sub

' *************************************************************
Private Sub ClearTextboxes()
    txtPriceName.Text = ""
    txtPrice.Text = ""
End Sub

' *************************************************************
' Add some screen visuals to make the form look nicer
'
Private Sub Form_Activate()
    Dither Me
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Initialize the form controls and form database
'
Private Sub Form_Load()
    InitDatabase
    AddColumn lvwPrice, "Price", 1500
    AddColumn lvwPrice, "Price Category", 2500
    
    m_SortAsc = True
    InitlvwPrice
    ClearTextboxes
    If lvwPrice.ListItems.Count > 0 Then
        lvwPrice_ItemClick lvwPrice.ListItems(1)
        cmdRemove.Enabled = True
    Else
        cmdRemove.Enabled = True
    End If
    cmdSave.Enabled = False
End Sub

' *************************************************************
' Open a handle to the database and the table we want to work with
'
Private Sub InitDatabase()
    Set m_wrkJet = CreateWorkspace("", "admin", "", dbUseJet)       ' Create Microsoft Jet Workspace object.
    Set m_DB = m_wrkJet.OpenDatabase(App.Path & "\Hotel.mdb")                   ' Open Database object from saved Microsoft Jet database (Hotel.mdb)
    Set m_RS = m_DB.OpenRecordset("tblPrice", dbOpenDynaset)        ' All our modifications are from tblPrice -- open this table now.
End Sub

' *************************************************************
' Initialize the lvwPrice listview to 2 columns - Price and Price Category
'
Private Sub InitlvwPrice()
    Dim lvw As ListView
    Dim RS As DAO.Recordset
    Dim li As ListItem
    Dim SortType As String
    
    Set lvw = lvwPrice
    
    lvw.ListItems.Clear
    If m_SortAsc = True Then
        SortType = "ASC"
    Else
        SortType = "DESC"
    End If
    
    ' Populate the listview with the corresponding data from tblPrice from the database
    Set RS = m_DB.OpenRecordset("Select * From tblPrice Order by Price " & SortType, dbOpenSnapshot)
    If Not (RS.BOF And RS.EOF) Then RS.MoveFirst
    While Not RS.EOF
        Set li = lvw.ListItems.Add
        li.Text = Format(RS!Price, "Currency")
        li.SubItems(1) = RS!PriceName
        li.Tag = RS!PriceID
        RS.MoveNext
    Wend
    Set RS = Nothing                                    ' Cleanup

End Sub

' *************************************************************
' When the user clicks on the listview, display the data in the textboxes
' for possible updating by the user.  If the user clicks on the listview
' without saving their changes, they will lose the changes.
'
Private Sub lvwPrice_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If m_RS.EditMode <> dbEditNone Then m_RS.CancelUpdate
    
    txtPriceName.Text = Item.SubItems(1)
    txtPrice.Text = Item.Text
    
    cmdSave.Enabled = False
    cmdRemove.Enabled = True
    cmdAdd.Enabled = True
End Sub

' *************************************************************
' The user has changed something -- enable the save button
'
Private Sub txtPrice_Change()
    If (txtPrice.Visible = True) Then cmdSave.Enabled = True
End Sub

' *************************************************************
' The user has changed something -- enable the save button
'
Private Sub txtPriceName_Change()
    If (txtPrice.Visible = True) Then cmdSave.Enabled = True
End Sub

' *************************************************************
' Provided as a convenience -- the user can choose which column
' to sort the listitems ascending
'
Private Sub lvwPrice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_SortAsc = Not m_SortAsc
    
    If ColumnHeader.index = 1 Then
        lvwPrice.Sorted = False
        InitlvwPrice
    Else
        lvwPrice.Sorted = True
        lvwPrice.SortKey = 1
        If m_SortAsc Then
            lvwPrice.SortOrder = lvwAscending
        Else
            lvwPrice.SortOrder = lvwDescending
        End If
    End If
End Sub

