VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO Connection Methods Sample"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOLEDB 
      Caption         =   "Use &OLE DB Connection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   6825
      Width           =   2415
   End
   Begin VB.CommandButton cmdADO 
      Caption         =   "Use &ADO Connection"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   6840
      Width           =   2040
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About Sample"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6735
      TabIndex        =   2
      Top             =   6825
      Width           =   1590
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8565
      TabIndex        =   1
      Top             =   6825
      Width           =   1380
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6720
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9885
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
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate new methods of adding ADO connections to reports
'           (New to Crystal Reports 8.0)
'
'           There are two new methods that provide a convenient
'           way to add ADO data sources to a report:
'           1.) AddADOCommand takes as parameters an active ADO connection
'           and an ADO command object.  In this example, a new connection
'           is created to a database, and the resulting recordset is assigned
'           at runtime to the report
'           2.) AddOLEDBSource takes as parameters the name of an active
'           ADO connection and the name of a table that can be accessed
'           through that connection.  In this example, the connection set
'           up through the VB Data Environment is used as a data source.
'
'           Both methods can be used with either the VB Data Environment
'           or another data source created at runtime.
'           Notice in the Crystal Report design environment that there is
'           no data source attached to the report at design time.
'
Option Explicit

Dim m_Report As New CrystalReport1

' The ADO connection to the local database.
Dim cnn1 As ADODB.Connection
Dim datcmd1 As ADODB.Command

' *************************************************************
' Demonstrate the use of AddADOCommand by opening an ADO data command,
' adding the data source to the report, and then adding a field
' to the report that uses that data source.
'
Private Sub cmdADO_Click()
    Dim fld As FieldObject
    Dim strCnn As String
    
    ' Open the data connection
    Set cnn1 = New ADODB.Connection
    strCnn = "Provider=MSDASQL;Persist Security Info=False;Data Source=Xtreme Sample Database;Mode=Read"
    cnn1.Open strCnn

    ' Create a new instance of an ADO command object
    Set datcmd1 = New ADODB.Command
    Set datcmd1.ActiveConnection = cnn1
    datcmd1.CommandText = "Customer"
    datcmd1.CommandType = adCmdTable

    ' Add the datasource to the report
    m_Report.Database.AddADOCommand cnn1, datcmd1
    ' Add a new field object to the report and set the field object to use
    ' the new data source.
    Set fld = m_Report.Section3.AddFieldObject("{ado.Customer Name}", 0, 0)
    LoadReport
End Sub

' *************************************************************
' Demonstrate the use of AddOLEDBSource by opening an ADO data source,
' adding the data source to the report, and then adding a field
' to the report that uses that data source.  In this example, the
' OLEDB source in the VB Data Environment is used
'
Private Sub cmdOLEDB_Click()
    Dim fld As FieldObject
    
    ' Add the datasource to the report
    m_Report.Database.AddOLEDBSource DataEnvironment1.Connection1, "Customer"
    ' Add a new field object to the report and set the field object to use
    ' the new data source.
    Set fld = m_Report.Section3.AddFieldObject("{Customer.Customer Name}", 0, 0)
    LoadReport
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub LoadReport()
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    cmdOLEDB.Enabled = False
    cmdADO.Enabled = False
End Sub

' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

