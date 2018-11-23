Attribute VB_Name = "Utilities"
' *************************************************************
' Purpose:  Location of the global variables used throughout the
'           program, as well as a routine to make the forms look
'           more visually pleasing
'

Option Explicit

Public g_StartForm As Form                  ' Holds a reference to frmStartup
Public g_ReportName As String               ' The name of the report to be displayed in frmDisplay
Public g_RecordSource As String             ' Used to display the recordsource for the ADO data control in frmDataGrid
Public g_RecordSet As ADODB.Recordset       ' Used by ADO data control in frmDataGrid
Public g_FormCaption As String              ' Used to determine the caption displayed in frmDataGrid

' Constants used to determine the name of the Crystal Report to display
Public Const vbBalanceStatement As String = "Balance Statement"
Public Const vbIncomeStatement As String = "Income Statement"
Public Const vbMailLabels As String = "Mail Labels"
Public Const vbWorldSales As String = "WorldWide Sales"
Public Const vbProductCatalog As String = "Product Catalog"
Public Const vbEmployeeProfiles As String = "Employee Profiles"

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
    frm.ScaleWidth = 176
    For intLoop = 0 To 175
        frm.Line (intLoop, 0)-(intLoop - 1, Screen.Height), RGB(0, intLoop, intLoop), B
    Next intLoop
End Sub

