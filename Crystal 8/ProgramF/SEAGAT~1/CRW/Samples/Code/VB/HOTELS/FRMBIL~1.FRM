VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmBilling 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmBilling.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPaperSize 
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
      Left            =   7380
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7410
      Width           =   2190
   End
   Begin VB.ComboBox cboPaperOrientation 
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
      Left            =   4995
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7410
      Width           =   2190
   End
   Begin VB.ComboBox cboPrinterDuplex 
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
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   7410
      Width           =   2190
   End
   Begin VB.ComboBox cboPaperSource 
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
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7410
      Width           =   2190
   End
   Begin VB.CommandButton cmdPrintOptions 
      Caption         =   "&Show Printer Options"
      Height          =   585
      Left            =   7365
      TabIndex        =   0
      Top             =   3780
      Width           =   2265
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "P&rint"
      Height          =   585
      Left            =   7365
      TabIndex        =   1
      Top             =   4560
      Width           =   2265
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "&Pay for these visits"
      Height          =   585
      Left            =   7365
      TabIndex        =   2
      Top             =   5355
      Width           =   2265
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   585
      Left            =   7365
      TabIndex        =   3
      Top             =   6135
      Width           =   2265
   End
   Begin CRVIEWERLibCtl.CRViewer crvBilling 
      Height          =   6675
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   7170
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label lblOwed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Owed:"
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
      Height          =   240
      Left            =   7380
      TabIndex        =   17
      Top             =   1845
      Width           =   1035
   End
   Begin VB.Label lblPaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Visits:"
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
      Height          =   240
      Left            =   7380
      TabIndex        =   16
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label lblUnpaid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unpaid Visits:"
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
      Height          =   240
      Left            =   7380
      TabIndex        =   15
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress: "
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
      Height          =   240
      Left            =   7380
      TabIndex        =   14
      Top             =   2250
      Width           =   2205
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size:"
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
      Height          =   240
      Index           =   3
      Left            =   7395
      TabIndex        =   13
      Top             =   7155
      Width           =   1020
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Orientation:"
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
      Height          =   240
      Index           =   2
      Left            =   5010
      TabIndex        =   11
      Top             =   7155
      Width           =   1575
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duplexing:"
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
      Height          =   240
      Index           =   1
      Left            =   2610
      TabIndex        =   10
      Top             =   7155
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Source:"
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
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   7155
      Width           =   1245
   End
   Begin VB.Label lblPrinterName 
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
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
      Height          =   285
      Left            =   75
      TabIndex        =   6
      Top             =   6825
      Width           =   5505
   End
   Begin VB.Image imgLogo 
      Height          =   825
      Left            =   7290
      Picture         =   "frmBilling.frx":030A
      Top             =   75
      Width           =   2400
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Based on m_StartDate, m_EndDate, CustomerName and ReportChoice,
'          displays a listing of all the visits that fit within the desired
'          criteria.
'          For example, if m_StartDate = "June 3, 1999", m_EndDate = "June 5, 1999",
'           and CustomerName = "John Doe", the report will list all stays for
'           John Doe between June 3rd and June 5th.
'
Option Explicit
  
Dim BillReport As New rptCustomer               ' The instantiation of rptCustomer
Dim ShowPrintOptions As Boolean                 ' A flag that holds whether or not the printer options are visible to the user.
Dim OriginalFormHeight As Long                  ' Holds the original, expanded height of the form
                                                ' (used when changing the height of the form)
Dim FirstShowComplete As Boolean                ' A flag that is set to true after the first time that the printer options are shown.
                                                ' This flag is used to prevent needless reloading of printer settings.
Dim m_Connection As ADODB.Connection
Dim adoRS As ADODB.Recordset
                                                
                                                
' *************************************************************
' Changes the size of the paper for the report.
' Note that your printer may not accept all the available paper sizes.
'
Private Sub cboPaperSize_Click()
    BillReport.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)    ' Set the papersize option
    If FirstShowComplete Then crvBilling.Refresh                            ' Show our changes only after the initial printer option enumeration
End Sub

' *************************************************************
' Changes the paper orientation for the displayed report.
'
Private Sub cboPaperOrientation_Click()
    BillReport.PaperOrientation = cboPaperOrientation.ItemData(cboPaperOrientation.ListIndex)
    If FirstShowComplete Then crvBilling.Refresh                        ' Show our changes only after the initial printer option enumeration
End Sub

' *************************************************************
' Changes the paper bin source for the displayed report.
' To enumerate the printer bins available for your printer, see EnumPrinterBins
' in Utilities.bas
' Note that your printer may override this setting to accommodate the papersize setting.
'
Private Sub cboPaperSource_Click()
    BillReport.PaperSource = cboPaperSource.ItemData(cboPaperSource.ListIndex)
End Sub

' *************************************************************
' Changes the printer duplex setting for the displayed report.
'
Private Sub cboPrinterDuplex_Click()
    BillReport.PrinterDuplex = cboPrinterDuplex.ItemData(cboPrinterDuplex.ListIndex)
End Sub

' *************************************************************
' The user is finished here; exit the form.
'
Private Sub cmdDone_Click()
    Unload Me                                   ' Unload this form
End Sub

' *************************************************************
' The user has chosen to pay their bill for visits not yet paid for
'
Private Sub cmdPay_Click()
    SetPaidRooms                                ' Pay the bill for all visible visits
    crvBilling.Refresh                          ' Refresh the viewer to reflect the changes
    RefreshLabels                               ' Refresh the Unpaid, Paid and Total Owing labels to reflect the new data
End Sub

' *************************************************************
' Print the current report, while showing the user the printer progress
' using lblPrinting
'
Private Sub cmdPrint_Click()
    lblPrinting.Visible = True                  ' Show the Printer progress label
    BillReport.PrintOut True                    ' Print the report and prompt the user for number of pages and collate option.
    lblPrinting.Visible = False                 ' All done printing.  Hide the printer progress label.
End Sub

' *************************************************************
' Show or hide the printer options.
' Three possible actions can occur:
' 1. If the printer options are currently hidden, and the options have not yet
' been enumerated (flagged by FirstShowComplete = False), enumerate all the printer options
' 2. If the printer options are currently hidden, but the options have already
' been enumerated (flagged by FirstShowComplete = True), just show the printer options
' 3. If the printer options are currently shown, hide them from view by resizing the form
'
Private Sub cmdPrintOptions_Click()
    Screen.MousePointer = vbHourglass           ' Show the user that we're working
    If ShowPrintOptions Then
        cmdPrintOptions.Caption = "&Hide Printer Options"   ' Change the button text
        Me.Height = OriginalFormHeight          ' Resize the form to show the printer options
        ShowPrintOptions = False                ' Set the printer option flag to show that the options are now hidden
        If Not FirstShowComplete Then           ' If the options have not already been enumerated, then
            GetPrinterOptions                   ' Populate the printer option combo boxes with all the available options
            FirstShowComplete = True            ' Set the flag to show we've enumerated the printer options
        End If
    Else
        cmdPrintOptions.Caption = "&Show Printer Options"   ' Change the button text
        Me.Height = OriginalFormHeight - 1020   ' Resize the form to hide the printer options
        ShowPrintOptions = True                 ' Set the printer option flag to show that the options are now shown
    End If
    Screen.MousePointer = vbDefault             ' Show the user that we've completed our work.
End Sub

' *************************************************************
' Enumerate all the available printer options for the report
' *** Note that GetPaperSource, GetPaperSize, GetPrinterDuplex, GetPaperOrientation will not retrieve
' accurate settings unless the printer settings have been saved in the report or the properties have
' been set some place in code.  In this case, rptCustomer has printer settings saved in it.
'
Private Sub GetPrinterOptions()
    Dim i As Integer                            ' Counter
    Dim PaperSource As Integer
    Dim PrinterDuplex As CRPrinterDuplexType
    Dim PaperOrientation As CRPaperOrientation
    Dim PaperSize As CRPaperSize

    lblPrinterName.Caption = "Report Printer Settings for: " & BillReport.PrinterName       ' Display the printer name
    ' Display the list of available printer bins in the cboPaperSource combo box.
    EnumPrinterBins BillReport.PrinterName, cboPaperSource
    With cboPaperSource
        PaperSource = BillReport.PaperSource    ' Get the report's paper source
        ' Cycle through the combo box and select the correct currently selected type of papersource in the report
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperSource Then .ListIndex = i
        Next i
    End With
    
    ' Display the list of available printer duplexing types in the cboPrinterDuplex combo box.
    ' Addcbo is a helper function to make the code cleaner
    ' Addcbo format:   <combo name>, <item caption>, <.itemdata(.listindex) to assign>
    Addcbo cboPrinterDuplex, "Simplex", crPRDPSimplex
    Addcbo cboPrinterDuplex, "Horizontal", crPRDPHorizontal
    Addcbo cboPrinterDuplex, "Vertical", crPRDPVertical
    PrinterDuplex = BillReport.PrinterDuplex    ' Get the report's printer duplex setting
    ' Cycle through the combo box and select the correct currently selected type of printer duplexing in the report
    With cboPrinterDuplex
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PrinterDuplex Then .ListIndex = i
        Next i
    End With
    
    ' Display the list of available paper orientations in the cboPaperOrientation combo box.
    Addcbo cboPaperOrientation, "Portrait", crPortrait
    Addcbo cboPaperOrientation, "Landscape", crLandscape
    PaperOrientation = BillReport.PaperOrientation  ' Get the report's paper orientation setting
    ' Cycle through the combo box and select the correct currently selected type of paper orientation in the report
    With cboPaperOrientation
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperOrientation Then .ListIndex = i
        Next i
    End With
        
    ' Add the large number of supported paper sizes to the cboPaperSize combobox
    Addcbo cboPaperSize, "Default", crDefaultPaperSize
    Addcbo cboPaperSize, "Letter", crPaperLetter
    Addcbo cboPaperSize, "Small Letter", crPaperLetterSmall
    Addcbo cboPaperSize, "Legal", crPaperLegal
    Addcbo cboPaperSize, "10x14", crPaper10x14
    Addcbo cboPaperSize, "11x17", crPaper11x17
    Addcbo cboPaperSize, "A3", crPaperA3
    Addcbo cboPaperSize, "A4", crPaperA4
    Addcbo cboPaperSize, "A4 Small", crPaperA4Small
    Addcbo cboPaperSize, "A5", crPaperA5
    Addcbo cboPaperSize, "B4", crPaperB4
    Addcbo cboPaperSize, "B5", crPaperB5
    Addcbo cboPaperSize, "C Sheet", crPaperCsheet
    Addcbo cboPaperSize, "D Sheet", crPaperDsheet
    Addcbo cboPaperSize, "Envelope 9", crPaperEnvelope9
    Addcbo cboPaperSize, "Envelope 10", crPaperEnvelope10
    Addcbo cboPaperSize, "Envelope 11", crPaperEnvelope11
    Addcbo cboPaperSize, "Envelope 12", crPaperEnvelope12
    Addcbo cboPaperSize, "Envelope 14", crPaperEnvelope14
    Addcbo cboPaperSize, "Envelope B4", crPaperEnvelopeB4
    Addcbo cboPaperSize, "Envelope B5", crPaperEnvelopeB5
    Addcbo cboPaperSize, "Envelope B6", crPaperEnvelopeB6
    Addcbo cboPaperSize, "Envelope C3", crPaperEnvelopeC3
    Addcbo cboPaperSize, "Envelope C4", crPaperEnvelopeC4
    Addcbo cboPaperSize, "Envelope C5", crPaperEnvelopeC5
    Addcbo cboPaperSize, "Envelope C6", crPaperEnvelopeC6
    Addcbo cboPaperSize, "Envelope C65", crPaperEnvelopeC65
    Addcbo cboPaperSize, "Envelope DL", crPaperEnvelopeDL
    Addcbo cboPaperSize, "Envelope Italy", crPaperEnvelopeItaly
    Addcbo cboPaperSize, "Envelope Monarch", crPaperEnvelopeMonarch
    Addcbo cboPaperSize, "Envelope Personal", crPaperEnvelopePersonal
    Addcbo cboPaperSize, "E Sheet", crPaperEsheet
    Addcbo cboPaperSize, "Executive", crPaperExecutive
    Addcbo cboPaperSize, "Fanfold Legal German", crPaperFanfoldLegalGerman
    Addcbo cboPaperSize, "Fanfold Standard German", crPaperFanfoldStdGerman
    Addcbo cboPaperSize, "Fanfold US", crPaperFanfoldUS
    Addcbo cboPaperSize, "Folio", crPaperFolio
    Addcbo cboPaperSize, "Ledger", crPaperLedger
    Addcbo cboPaperSize, "Note", crPaperNote
    Addcbo cboPaperSize, "Quarto", crPaperQuarto
    Addcbo cboPaperSize, "Statement", crPaperStatement
    Addcbo cboPaperSize, "Tabloid", crPaperTabloid
    PaperSize = BillReport.PaperSize    ' Get the report's paper size setting
    ' Cycle through the combo box and select the correct currently selected type of paper size in the report
    With cboPaperSize
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperSize Then .ListIndex = i
        Next i
    End With
End Sub

' *************************************************************
' A small helper function for GetPrinterOptions that helps reduce the amount
' of code to write
'   Addcbo format:   <combo name to add item to>, <item caption>, <.itemdata(.listindex) to assign>
Private Sub Addcbo(cbo As ComboBox, Name As String, index As Integer)
    cbo.AddItem Name                        ' Add the name of the item to the combo box
    cbo.ItemData(cbo.NewIndex) = index      ' Set the .itemdata(.listindex) for later retrieval
End Sub

' *************************************************************
' Display some useful info about the currently displayed report:
'   Unpaid visits, paid visits, and total amount yet to be paid.
'
Private Sub RefreshLabels()
    DoEvents                                 ' Bring everything up to date (ie. refresh the viewer and the display
    ' Get the appropriate numbers that have been saved in the report.
    lblUnpaid.Caption = "Unpaid Visits: " & BillReport.DaysOwed
    lblPaid.Caption = "Paid Visits: " & BillReport.DaysPaid
    lblOwed.Caption = "Amount Owed: " & Format(BillReport.TotalBill, "Currency")
End Sub

' *************************************************************
' Filter the data in the viewer to display only the records that we want to see.
'
Private Sub RefreshViewer()
    Dim s As String                         ' Temporary variable of convenience
        
    Screen.MousePointer = vbHourglass
    
    ' If m_StartDate = "12:00:00 AM", then we should display records for all available dates,
    ' otherwise only show the records between the dates we've set in frmReservation
    If m_StartDate <> "12:00:00 AM" Then
        s = "{qryCustomerReservation.Date} >= Date(" & Format(m_StartDate, "yyyy,mm,dd") & ") and " & _
            "{qryCustomerReservation.Date} <= Date(" & Format(m_EndDate, "yyyy,mm,dd") & ") and"
    End If
    
    ' If we are showing information only for a particular customer, then filter by that customer name
    If ReportChoice = vbGroupName Then
        s = s & " {@CustomerName} = '" & Trim(CustomerName) & "'"
        ' If we are showing information only for a particular room, then filter by that room number
    ElseIf ReportChoice = vbGroupRoom Then
        s = s & " {qryCustomerReservation.RoomNumber} = '" & Trim(m_RoomNumber) & "'"
    End If
    
    BillReport.RecordSelectionFormula = s   ' Apply the new filter-formula to the report
    RefreshLabels                           ' Refresh the "useful information" labels
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Initialize the form, setting the report source and make sure the visual
' aspects of the form are correct
'
Private Sub Form_Load()
    Dim fld1 As FieldObject                 ' A temporary holder
    Dim strConnect As String
    
    Set lblPrinting = Me.lblProgress        ' Allows the caption for the label to be set by code in other modules (ie. rptCustomer.dsr)
    lblProgress.Visible = False             ' Reset and hide the report printing progress label
    lblProgress.Caption = "0"

    ' Create and bind the ADO Recordset object
    Set m_Connection = New ADODB.Connection
    Set adoRS = New ADODB.Recordset

    ' Open the connection
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hotel.mdb;Mode=Read"
    m_Connection.Open strConnect
    
    adoRS.Open "Select * From qryCustomerReservation", m_Connection, adOpenKeyset, adLockBatchOptimistic
    BillReport.Database.SetDataSource adoRS
    crvBilling.ReportSource = BillReport    ' Set the viewer source to our instantiation of rptCustomer
    RefreshViewer                           ' Set the correct datasource for the report

    ' FieldGroupDisplay is used in the rptCustomer report to let the report know
    ' what type of heading to display in FieldReportTitle, as well as which fields
    ' of CustomerName, RoomNumber, Price, and Date should be displayed or made invisible
    ' SetUnboundFieldSource allows you to send an unbound value to a report that can be set at runtime.
    ' In this case, I'm using the unbound field as a flag that other components of the report
    ' can use to set their display options.
    Set fld1 = BillReport.FieldGroupDisplay
    fld1.SetUnboundFieldSource ReportChoice
    
    ' To provide a nicer look, set the paper size to something that fits the window,
    ' and make the margins equal
    BillReport.PaperSize = crPaperEnvelope11
    BillReport.RightMargin = BillReport.LeftMargin
    
    crvBilling.ViewReport                   ' Finally, view the report
    RefreshLabels                           ' Get the newly updated info from the report and display it.
    
    ' Set the form caption
     If ReportChoice = vbGroupName Then
        Me.Caption = "Report for: " & CustomerName
    Else
        Me.Caption = "Report for Room " & m_RoomNumber
        cmdPay.Enabled = False              ' Only allow bill paying when a specific customer is selected
    End If
    
    ' Some final initializations
    ShowPrintOptions = False                ' Report printer options are not initially displayed
    OriginalFormHeight = Me.Height          ' Save the expanded form height for later use
    cmdPrintOptions_Click                   ' Use the cmdPrintOptions_Click function to resize the form to its smaller size to hide the print options
    FirstShowComplete = False               ' Ensure that the form knows that we haven't enumerated the printer options yet.
End Sub

' *************************************************************
' Add some visual enhancements to the form
'
Private Sub Form_Activate()
    Dither Me                               ' Make the nice teal background shading
    Screen.MousePointer = vbDefault         ' All initializations should have been completed; provide the user with a visual clue
End Sub

' *************************************************************
' Clean-up after ourselves
'
Private Sub Form_Unload(Cancel As Integer)
    Set BillReport = Nothing                ' Terminate this instance of the form to prevent a potential memory leak problem
    m_StartDate = CDate("12:00:00 AM")      ' Reset the start date so that nothing is left behind from this form.
End Sub
