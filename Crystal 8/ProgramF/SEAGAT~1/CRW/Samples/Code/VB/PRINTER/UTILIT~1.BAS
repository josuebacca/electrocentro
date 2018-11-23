Attribute VB_Name = "Utilities"
'*********************************************************************
' Purpose: Provide a set of utility functions
'

Option Explicit

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



