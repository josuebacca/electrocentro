Attribute VB_Name = "Utilities"
' *************************************************************
' Purpose: General utilities used throughout the application
'
Option Explicit

' Used in frmBilling to set the filtering options for the report
Public CustomerName As String                   ' The name of the customer to filter by
Public m_StartDate As Date
Public m_EndDate As Date
Public m_RoomNumber As String

' Used in the code parts of rptCustomer and rptRoomReport to provide the user
' with visual cues
Public cmdButton As CommandButton
Public lblPrinting As Label

' Used to set what type of report the user wishes to see
Public ReportChoice As Integer  ' 1 = Group by Customer Name
                                ' 2 = Group by RoomNumber
Public Const vbGroupName = 1
Public Const vbGroupRoom = 2

' A string used in frmReservation
Public Const vbAllCustomers = "<All Customers>"

' API functions used in the "Find and Open Browser" routine
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
(ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As _
String) As Long


' API functions and constants used in EnumPrinterBins
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
   ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" ( _
   ByVal hPrinter As Long) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" _
   Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
   ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
   ByVal dev As Long) As Long

Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12

' *************************************************************
' Set all the visits that are currently being displayed to "Paid"
'
Public Sub SetPaidRooms()
    Dim m_DB As DAO.Database                    ' The handle to the open database
    Dim m_wrkJet As DAO.Workspace
    Dim m_RS As DAO.Recordset
    Dim SQL As String
    Dim d As Date                               ' Temporary date variable
    Dim Name() As String                        ' Used to parse the customer name
    
    Name = Split(CustomerName, ",")             ' Parse the customer name into first and last name
    
    Set m_wrkJet = CreateWorkspace("", "admin", "", dbUseJet)   ' Create Microsoft Jet Workspace object.
    Set m_DB = m_wrkJet.OpenDatabase(App.Path & "\Hotel.mdb")   ' Open Database object from saved Microsoft Jet database (Hotel.mdb)
    
    ' If we don't know the start date, we have to figure out the date of the first customer visit as well as the date of the last customer visit.
    If m_StartDate = "12:00:00 AM" Then         ' We have an unknown start time, pay all visits for this customer
        ' Build the SQL string
        SQL = "Select * From qryCustomerReservation Where " & _
          " LastName = '" & Trim(Name(0)) & "'" & _
          " and FirstName = '" & Trim(Name(1)) & "'"
        Set m_RS = m_DB.OpenRecordset(SQL, dbOpenSnapshot)
        If m_RS.RecordCount > 0 Then
            m_RS.MoveFirst
            m_StartDate = m_RS!Date             ' The date of the customer's first visit
            m_RS.MoveLast
            m_EndDate = m_RS!Date               ' The date of the customer's last visit
        End If
    End If
    
    ' Since we know for certain the start and end dates of the customer's visits that we're paying for,
    ' set all of those visits between startdate and enddate to "Paid"
    d = m_StartDate
    While d <= m_EndDate
        SQL = "Select * From qryCustomerReservation Where " & _
          " LastName = '" & Trim(Name(0)) & "'" & _
          " and FirstName = '" & Trim(Name(1)) & "'" & _
          " and Date = #" & Format(d, "mmm d yyyy") & "#"
        Set m_RS = m_DB.OpenRecordset(SQL, dbOpenDynaset)
        If m_RS.RecordCount > 0 Then
            m_RS.MoveFirst
            While Not m_RS.EOF
                m_RS.Edit
                m_RS!PaidFor = True             ' Set the "Paid" flag to True
                m_RS.Update
                m_RS.MoveNext
            Wend
        End If
        d = d + 1
    Wend
       
    ' Cleanup
    Set m_DB = Nothing
    Set m_wrkJet = Nothing
    Set m_RS = Nothing
End Sub

'*********************************************************************
' Add a column to a ListView control and set its attibutes.
' lvw is the listview to manipulate
' Caption is the header caption
' Width is the width of the column header in Twips
'
Public Sub AddColumn(lvw As ListView, Caption As String, Width As Single)
   Dim ch As ColumnHeader
   Set ch = lvw.ColumnHeaders.Add()             ' Create a new column.

   ' Set the new column's attributes.
   ch.Text = Caption
   ch.Width = Width
End Sub

'*********************************************************************
' A short procedure to give forms a nice teal shading.
'
Public Sub Dither(frm As Form)
   Dim intLoop As Integer                       ' Counter
   
   ' Set the pen parameters
   frm.DrawStyle = vbInsideSolid
   frm.DrawMode = vbCopyPen
   frm.ScaleMode = vbPixels
   frm.DrawWidth = 8
   frm.ScaleWidth = 256
   For intLoop = 0 To 255
      frm.Line (intLoop, 0)-(intLoop - 1, Screen.Height), RGB(0, intLoop, intLoop), B
   Next intLoop
End Sub

'*********************************************************************
' Add a list of the available paper sources for <PrinterName> to
' the combobox <cbo>
'
Public Sub EnumPrinterBins(PrinterName As String, cbo As ComboBox)
    Dim prn As Printer
    Dim hPrinter As Long                ' Handle of the selected printer
    Dim dwbins As Long                  ' The number of paperbins in the printer
    Dim i As Long                       ' counter
    Dim nameslist As String             ' The string returned with all the bin names
    Dim NameBin As String               ' The parsed bin name
    Dim numBin() As Integer             ' Used as part of the DeviceCapabilities API call
     
    cbo.Clear
    For Each prn In Printers
        ' Look through all the currently installed printers
        If prn.DeviceName = PrinterName Then
            ' We've found our printer -- open a handle to it
            If OpenPrinter(prn.DeviceName, hPrinter, 0) <> 0 Then
                ' Get the available bin numbers
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, ByVal vbNullString, 0)
                ReDim numBin(1 To dwbins)
                nameslist = String(24 * dwbins, 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINS, numBin(1), 0)
                dwbins = DeviceCapabilities(prn.DeviceName, prn.Port, _
                                        DC_BINNAMES, ByVal nameslist, 0)
                For i = 1 To dwbins
                    ' For each bin number, add its corresponding name
                    ' to our combo box
                    NameBin = Mid(nameslist, 24 * (i - 1) + 1, 24)
                    NameBin = Left(NameBin, InStr(1, NameBin, Chr(0)) - 1)
                    cbo.AddItem NameBin
                    cbo.ItemData(cbo.NewIndex) = numBin(i)
                Next i
                ' Close the printer
                Call ClosePrinter(hPrinter)
            Else
                ' OpenPrinter failed, so we can't retrieve information about it
                cbo.AddItem prn.DeviceName & "  <Unavailable>"
            End If
        End If
    Next prn
End Sub

