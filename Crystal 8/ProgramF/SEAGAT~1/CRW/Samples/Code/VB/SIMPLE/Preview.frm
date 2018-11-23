VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "CRViewer.dll"
Begin VB.Form frmPreview 
   Caption         =   "Simple Demo"
   ClientHeight    =   7410
   ClientLeft      =   1650
   ClientTop       =   1950
   ClientWidth     =   8775
   Icon            =   "Preview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8775
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
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
      EnableAnimationControl=   -1  'True
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
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: An introduction to using the Report Designer.  Some
'          of the basic methods used in manipulating and creating
'          reports are discussed in the code comments

Option Explicit

Dim m_Report As New CrystalReport1    ' Create a new instance of the Crystal Report
Dim m_RS As New ADOR.Recordset        ' Create and ADO record set

' *************************************************************
' You can also capture all kinds of events from the Smart Viewer
'
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    MsgBox "The Smart Viewer fires 27 different events!"
End Sub

' *************************************************************
' This report was created using the 'xtreme sample data' as it's data source,
' ie. the CrystalReport1.dsr remembers that information.  You can also
' use the SetDataSource method on the Database object to change the
' data source at runtime.  For example you can create a DAO/RDO/ADO record set
' at runtime and pass it to the report engine at runtime and the data
' in the report will reflect the data from the record set.  Obviously,
' the report field names and types must match the names and types in the
' record set.
'
Private Sub Form_Load()
    ' Open the recordset
    m_RS.Open "Select * from customer where country = 'USA'", "xtreme sample database"

    ' Pass this recordset to the report engine to use as the datasource
    ' If you comment out the following line the report will show all customer
    ' in the world because the report was created to show all customers and
    ' this information is stored in the DSR.
    m_Report.Database.SetDataSource m_RS

    ' The Smart Viewer (CRVIEWER) is the preview window for the report
    ' If you are don't wish to show the report to the user, ie. only want
    ' to print the report, then you don't need to use the viewer at all.
    CRViewer1.ReportSource = m_Report

    ' You have full access to all the report objects, so you may want
    ' to change the text in the title of the report
    ' Comment out the line below to see the text that is in the report
    ' definition.
    m_Report.Text6.SetText "You can change text objects through code!"

    '  The Smart viewer has a object model that allows you to modify the
    '  look and feel of the preview window at runtime.
    '  The ViewReport method will start the process of the report.
    CRViewer1.ViewReport
End Sub

' *************************************************************
' This code resizes the Smart Viewer when the form is resized
'
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub
