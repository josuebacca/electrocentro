VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReservation 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservation Management"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmReservation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMoreRoomInfo 
      Height          =   360
      Left            =   6645
      Picture         =   "frmReservation.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1635
      Width           =   480
   End
   Begin VB.CommandButton cmdMoreCustomerInfo 
      Height          =   420
      Left            =   5745
      Picture         =   "frmReservation.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   45
      Width           =   480
   End
   Begin VB.ComboBox cboRoom 
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
      Left            =   3630
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1635
      Width           =   2880
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "     &Remove Reservations              for these days"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7260
      TabIndex        =   10
      Top             =   1425
      Width           =   2475
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add Reservation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7260
      TabIndex        =   9
      Top             =   738
      Width           =   2475
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
      Height          =   585
      Left            =   7260
      TabIndex        =   8
      Top             =   52
      Width           =   2475
   End
   Begin VB.ComboBox cboCustomerName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2445
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   52
      Width           =   3180
   End
   Begin CRVIEWERLibCtl.CRViewer crvReservationView 
      Height          =   6765
      Left            =   75
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2115
      Width           =   9645
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   0   'False
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComCtl2.DTPicker DTShowEndDate 
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   8880
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22937600
      CurrentDate     =   36298
   End
   Begin MSComCtl2.DTPicker DTShowStartDate 
      Height          =   375
      Left            =   1695
      TabIndex        =   6
      Top             =   8880
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22937600
      CurrentDate     =   36298
   End
   Begin MSComCtl2.DTPicker DTEndDate 
      Height          =   375
      Left            =   3630
      TabIndex        =   3
      Top             =   1155
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22937600
      CurrentDate     =   36298
   End
   Begin MSComCtl2.DTPicker DTStartDate 
      Height          =   375
      Left            =   3630
      TabIndex        =   2
      Top             =   675
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22937600
      CurrentDate     =   36298
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date:"
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
      Index           =   3
      Left            =   2430
      TabIndex        =   16
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date:"
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
      Index           =   2
      Left            =   2430
      TabIndex        =   15
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room #:"
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
      Left            =   2430
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose a customer and a time period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   6
      Left            =   270
      TabIndex        =   13
      Top             =   705
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Start Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   60
      TabIndex        =   12
      Top             =   8910
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview End Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   4965
      TabIndex        =   11
      Top             =   8940
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Reservations for:"
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
      Left            =   90
      TabIndex        =   4
      Top             =   105
      Width           =   2655
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: The main form in the application -- allows the user to add and
'          remove reservations, and display the reservations for any
'          range of dates
'
Option Explicit

Dim m_Report As New rptRoomReport                       ' The instantiation of rptRoomReport used throughout the form
Dim DateKeyPressed As Boolean                           ' Set to true if the user modifies the start or end date with the keyboard
Dim m_Connection As ADODB.Connection                    ' The connection for the data source assigned to the report
Dim adoRS As ADODB.Recordset                            ' The ADO recordset assigned to the report data source

' *************************************************************
' The Crystal Reports viewer allows the user to detect what type
' of data field has been double clicked.  We take advantage of this
' feature to provide a convenient method for the user to bring up
' a report about a particular customer or room.
'
Private Sub crvReservationView_DblClicked(ByVal x As Long, ByVal Y As Long, EventInfo As Variant, UseDefault As Boolean)
    Dim m As New CRVEventInfo
    Dim t As Long
    
    Set m = EventInfo
    ' Now we're going to figure out the name of the field
    If (m.Type = crDatabaseFieldType) Then
        ' Has to be one of the two database fields in the report
        ' Use co-ordinates
        If (x >= 40 And x <= 181) Then
            ' Must have double clicked on a room number - show the report for that room
            CustomerName = ""
            ReportChoice = vbGroupRoom
            m_RoomNumber = m.Text
            ' Clean up any carriage-returns and line feeds that might be in the room number
            m_RoomNumber = Replace(m_RoomNumber, vbCr, "")
            m_RoomNumber = Replace(m_RoomNumber, vbLf, "")
            ShowBill
        ElseIf (x >= 182 And x <= 285) Then
            ' Must have double clicked on a price
            MsgBox "Room Price: " & m.Text, vbInformation, "Price Info"
        End If
    ElseIf (m.Type = crFormulaFieldType) Then
        ' Only one formula in the report (CustomerName) - show the report for that customer
        ReportChoice = vbGroupName
        m_RoomNumber = ""
        CustomerName = m.Text
        ' Clean up any carriage-returns and line feeds that might be in the name
        CustomerName = Replace(CustomerName, vbCr, "")
        CustomerName = Replace(CustomerName, vbLf, "")
        ShowBill
    End If
End Sub


' *************************************************************
' Display frmBilling to provide a complete listing of this customer's
' activity at the hotel between the start and end dates shown on screen
'
Private Sub cmdMoreCustomerInfo_Click()
    CustomerName = Trim(cboCustomerName.Text)       ' Save the current customer's name to a public
                                                    '      variable for later use in frmBilling for Report filtering.
    ReportChoice = vbGroupName                      ' Set the option that we want the report to show Records for this customer
    m_RoomNumber = ""                               ' Since we want to know the customer's stays in ALL rooms, reset room number to ""
    Screen.MousePointer = vbHourglass
    ShowBill
End Sub

' *************************************************************
' Set the start and end dates that the report in frmBilling will
' use as filters
'
Private Sub ShowBill()
    Screen.MousePointer = vbHourglass
    DoEvents
    m_StartDate = DTShowStartDate.Value
    m_EndDate = DTShowEndDate.Value
    frmBilling.Show vbModal
End Sub

' *************************************************************
' Display frmBilling to provide a complete listing of everyone
' who has ever stayed in this room between the start and end dates
' shown on screen
'
Private Sub cmdMoreRoomInfo_Click()
    Dim Room() As String
    
    CustomerName = ""                                   ' Since we want to know the stays in this room for ALL customers, reset customername to ""
    ReportChoice = vbGroupRoom                          ' Set the option that we want the report to show Records for this room
    Room = Split(cboRoom.Text, Chr(151))
    m_RoomNumber = Trim(Room(0))                        ' Save the current room number to a public
                                                        '   variable for later use in frmBilling for Report filtering.
    ShowBill
End Sub

' *************************************************************
' Save the customer's new reservation to the database.
' To allow for flexibility, a separate record is saved for each day that the
' user stays in the hotel.
'
Private Sub cmdSave_Click()
    Dim m_Connection2 As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim strConnect As String
    Dim d As Date
    Dim m_EndDate As Date
    Dim RoomNumber() As String
    Dim s As String
    
    If DTStartDate.Value > DTEndDate.Value Then
        MsgBox "The Reservation start date must be before the Reservation end date", _
                vbInformation, "Change Reservation Dates"
        Exit Sub
    End If
        
    ' By design, a reservation can only be added for a specific customer
    If cboCustomerName.Text = vbAllCustomers Then
        MsgBox "Sorry, you aren't permitted to add reservations for all customers. " & vbNewLine & _
                "Please select a specific customer. ", vbExclamation, "Unable to comply"
        Exit Sub
    End If
        
    ' Do some preparation work
    d = DTStartDate.Value
    m_EndDate = DTEndDate.Value
    RoomNumber = Split(cboRoom.Text, Chr(151))
    
    Set m_Connection2 = New ADODB.Connection
    Set RS = New ADODB.Recordset
            
    ' Open the connection
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read|Write"
    m_Connection2.Open strConnect
    
    RS.CursorType = adOpenDynamic
    RS.LockType = adLockOptimistic
    RS.Open "tblReservation", m_Connection2, adOpenDynamic, adLockOptimistic, adCmdTable
    
    ' Add the records to the database
    With RS
        While d <= m_EndDate
            .AddNew
            ' Open a recordset and enter the date, customer ID, room number for each date
            !Date = d
            !Room = RoomNumber(0)
            !CustomerID = cboCustomerName.ItemData(cboCustomerName.ListIndex)
            d = d + 1
            .Update
        Wend
    End With
     ' Let the user know that the addition of the reservation was successful
    MsgBox "Reservation added for " & cboCustomerName.Text & " in Room " & RoomNumber(0) & _
            "from " & Format(DTStartDate.Value, "dddd, mmm d yyyy") & _
            " to " & Format(m_EndDate, "dddd, mmm d yyyy") _
            , vbInformation, "Successful Reservation"
            
    RS.Close
    ReservationRefresh
End Sub

' *************************************************************
' A helper function that ensures the viewer has the most recent
' Reservation data
'
Private Sub ReservationRefresh()
    adoRS.Requery
    Dim i As Long
    Screen.MousePointer = vbHourglass
    ' Yield the processor to ADO to give it time to update the Recordset
    For i = 1 To 550000
        DoEvents
    Next i
    RefreshViewer
    Screen.MousePointer = vbNormal
End Sub

' *************************************************************
' Remove all reservations between the specified dates for a
' specific customer
'
Private Sub cmdRemove_Click()
    Dim m_Connection2 As ADODB.Connection
    Dim cmdDelete As ADODB.Command
    Dim strConnect As String
    Dim SQL As String
    Dim Response As Integer
    Dim s As String
    
    If DTStartDate.Value > DTEndDate.Value Then
        MsgBox "The Reservation start date must be before the Reservation end date", _
                vbInformation, "Change Reservation Dates"
        Exit Sub
    End If
    
    ' By design, the user is not permitted to delete all reservations for
    ' all customers (to prevent accidents)
    If cboCustomerName.Text = vbAllCustomers Then
        MsgBox "Sorry, you aren't permitted to delete reservations for all customers. " & vbNewLine & _
                "Please select a specific customer. ", vbExclamation, "Unable to comply"
        Exit Sub
    End If
    
    ' Ensure that the user is REALLY sure about this
    Response = MsgBox("This will remove all reservations for " & cboCustomerName.Text & _
            " between " & DTStartDate.Value & " and " & DTEndDate.Value _
            , vbOKCancel, "Remove Reservation")
    If Response = vbOK Then
        ' Open the connection
        Set m_Connection2 = New ADODB.Connection
        strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read|Write"
        m_Connection2.Open strConnect
       
        Set cmdDelete = New ADODB.Command
        Set cmdDelete.ActiveConnection = m_Connection2
        ' Delete the all records between the two dates
        cmdDelete.CommandText = "Delete * from tblReservation Where " & _
              " CustomerID = " & cboCustomerName.ItemData(cboCustomerName.ListIndex) & _
              " and (Date >= #" & Format(DTStartDate.Value, "mmm d yyyy") & "#) " & _
              " and (Date <= #" & Format(DTEndDate.Value, "mmm d yyyy") & "#)"
        cmdDelete.Execute
        m_Connection2.Close
        ReservationRefresh
    End If
End Sub

' *************************************************************
' All done - exit the form
'
Private Sub cmdDone_Click()
    Unload Me
End Sub

' *************************************************************
' Two possibilities:
' 1. <All Customers> is selected - disable reservation adding and deleting
' 2. A specific customer is selected - enable reservation adding and deleting
Private Sub cboCustomerName_Click()
    If cboCustomerName.Enabled = False Then Exit Sub ' Prevent the user from requesting another update
                                                     ' while the report is still refreshing
    If Me.Visible Then RefreshViewer                 ' Get the report to automatically update
    If cboCustomerName.Text = vbAllCustomers Then
        cmdSave.Enabled = False
        cmdRemove.Enabled = False
        cmdMoreCustomerInfo.Enabled = False
        CustomerName = ""
    Else
        ' Specific customer name selected - enable the reservation options
        cmdSave.Enabled = True
        cmdRemove.Enabled = True
        cmdMoreCustomerInfo.Enabled = True
        CustomerName = Trim(cboCustomerName.Text)
    End If
End Sub

' *************************************************************
' Update the report to reflect the user's desire to view all records
' until a later date.
'
Private Sub DTShowEndDate_CloseUp()
    If Me.Visible And DTShowEndDate.Enabled Then RefreshViewer
End Sub

' *************************************************************
Private Sub DTShowEndDate_KeyUp(KeyCode As Integer, Shift As Integer)
    DateKeyPressed = True
End Sub

' *************************************************************
' Handle the possibility the user changes the value with the keyboard
'
Private Sub DTShowEndDate_LostFocus()
    If DateKeyPressed = True Then
        DateKeyPressed = False
        DTShowEndDate_CloseUp
    End If
End Sub

' *************************************************************
' Update the report to reflect the user's desire to view all records
' from an earlier date.
'
Private Sub DTShowStartDate_CloseUp()
    If Me.Visible Then
        ' If the user chooses a starting date ahead of the end date,
        ' move the end date to 30 days after the start date
        If DTShowStartDate.Value > DTShowEndDate.Value Then
            If DTShowStartDate.Month = mvwDecember Then
                DTShowEndDate.Month = mvwJanuary
                DTShowEndDate.Year = DTShowStartDate.Year + 1
            Else
                DTShowEndDate.Month = DTShowStartDate.Month + 1
            End If
        End If
        If DTShowStartDate.Enabled Then RefreshViewer
    End If
End Sub

' *************************************************************
Private Sub DTShowStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    DateKeyPressed = True
End Sub

' *************************************************************
' Handle the possibility the user changes the value with the keyboard
'
Private Sub DTShowStartDate_LostFocus()
    If DateKeyPressed = True Then
        DateKeyPressed = False
        DTShowEndDate_CloseUp
    End If
End Sub

' *************************************************************
' Filter the data in the viewer to display only the records that we want to see.
'
Private Sub RefreshViewer()
    Dim SQL As String
    Dim Name() As String                    ' Used to help parse cboCustomerName.text
    
    If crvReservationView.IsBusy Then Exit Sub            ' Extra safety precaution
    Screen.MousePointer = vbHourglass
    DTShowStartDate.Enabled = False
    DTShowEndDate.Enabled = False
    cboCustomerName.Enabled = False
    m_Report.IsNoData = False
    
    SQL = "Select * from qryCustomerReservation Where " & _
          "tblReservation.Date >= #" & Format(DTShowStartDate, "yyyy/mmm/dd") & "#" & _
          " and tblReservation.Date <= #" & Format(DTShowEndDate, "yyyy/mmm/dd") & "#"
          
    If cboCustomerName.Text <> vbAllCustomers Then
        Name = Split(cboCustomerName.Text, ",")     ' Parse into first and last name
        SQL = SQL & " and LastName = '" & Trim(Name(0)) & "'" & _
                 " and FirstName = '" & Trim(Name(1)) & "'"
    End If
    adoRS.Close
    adoRS.Open SQL, m_Connection, adOpenDynamic, adLockBatchOptimistic
    m_Report.Database.SetDataSource adoRS
    
    If crvReservationView.Visible Then
        m_Report.ReadRecords
        crvReservationView.Refresh
    End If
    DTShowStartDate.Enabled = True
    DTShowEndDate.Enabled = True
    cboCustomerName.Enabled = True
    
    Screen.MousePointer = vbDefault
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
    Dim d As Date                       ' Temporary date holder
    Screen.MousePointer = vbHourglass
        
    ' Set the reservation start and end date to May 1.
    ' It would be better to set the start date to the current date (using 'Date')
    d = CDate("May 1, 1999")
    DTStartDate.Value = d
    DTEndDate.Value = d
    DateKeyPressed = False
    
    ' Set the view to show a month from the start date
    DTShowStartDate.Value = DTStartDate.Value
    DTShowEndDate.Value = DTShowStartDate.Value + 30
    
    PopulatecboCustomerName
    PopulatecboRoom
    OpenReport
End Sub

' *************************************************************
' Open the Crystal report into the viewer
'
Private Sub OpenReport()
    Dim SQL As String
    Dim strConnect As String
    Set cmdButton = Me.cmdMoreCustomerInfo              ' Save a reference to the command button so that we can reference it later
    
    ' Create and bind the ADO Recordset object
    Set m_Connection = New ADODB.Connection
    Set adoRS = New ADODB.Recordset

    ' Open the connection
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read"
    m_Connection.Open strConnect
    
    adoRS.Open "Select * From qryCustomerReservation", m_Connection, adOpenDynamic, adLockBatchOptimistic
    m_Report.Database.SetDataSource adoRS
   
    crvReservationView.ReportSource = m_Report          ' other code modules can enable and disable it
    m_Report.PaperSize = crPaperEnvelope14              ' Set a visually pleasing paper size for the display
                                                        ' If you're going to print the report, you should change the papersize to letter
    RefreshViewer
    crvReservationView.ViewReport
End Sub

' *************************************************************
' Populate cboCustomerName with all the customer names from the database
'
Private Sub PopulatecboCustomerName()
    Dim m_Connection As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim strConnect As String
    
    Set m_Connection = New ADODB.Connection
    Set RS = New ADODB.Recordset
        
    ' Open the connection
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read|Write"
    m_Connection.Open strConnect
    
    RS.CursorType = adOpenDynamic
    RS.LockType = adLockOptimistic
    RS.Open "tblCustomer", m_Connection, , , adCmdTable
    

    Dim s As String
    
    With cboCustomerName
        ' Add "<All Customers>"
        .Clear
        .AddItem vbAllCustomers
        .ItemData(.NewIndex) = 0        ' No customer will ever have ID = 0
        
        If (RS.BOF And RS.EOF) Then
            Exit Sub                    ' No customers in the database
        Else
            RS.MoveFirst
        End If
                
        ' Populate the control with names from the database
        While Not RS.EOF
            .AddItem RS!LastName & ", " & RS!FirstName
            .ItemData(.NewIndex) = RS!CustomerID
            RS.MoveNext
        Wend
    End With
    
    Set RS = Nothing                    ' Cleanup
    ' Default to the first name in the list
    If cboCustomerName.ListCount > 0 Then cboCustomerName.ListIndex = 0
End Sub

' *************************************************************
' Populate the cboRoom combobox with all the existing room numbers
' in the database
'
Private Sub PopulatecboRoom()
    Dim wrkJet As DAO.Workspace
    Dim m_DB As DAO.Database
    Dim RS As DAO.Recordset
    
    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)     ' Create Microsoft Jet Workspace object.
    Set m_DB = wrkJet.OpenDatabase(App.Path & "\Hotel.mdb")     ' Open Database object from saved Microsoft Jet database (Hotel.mdb)
    Set RS = m_DB.OpenRecordset("qryRoomPrices", dbOpenSnapshot) ' All our modifications are from qryRoomPrices -- open this table now.
    
    cboRoom.Clear
    If Not (RS.BOF And RS.EOF) Then RS.MoveFirst
    
    ' Add an item to cboRoom for each record in the database
    While Not RS.EOF
        cboRoom.AddItem RS!RoomNumber & " " & Chr(151) & " " & Format(RS!Price, "Currency")
        cboRoom.ItemData(cboRoom.NewIndex) = RS!PriceID
        RS.MoveNext
    Wend
    If cboRoom.ListCount > 0 Then cboRoom.ListIndex = 0
    Set RS = Nothing

End Sub

' *************************************************************
' Clean-up the report to prevent possible memory leaks
Private Sub Form_Unload(Cancel As Integer)
    Set m_Report = Nothing
End Sub
