VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Object Creation API Sample"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6915
      TabIndex        =   2
      Top             =   6855
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Default         =   -1  'True
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
      Left            =   8520
      TabIndex        =   1
      Top             =   6855
      Width           =   1275
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6735
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   9795
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
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
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate the use of the Report Object Creation API
'           (new to Crystal Reports 8.0).
'
'           The Creation API functions allow runtime creation of
'           report objects; including text, database field, unbound
'           field, chart, special, box, cross-tab, blob field, line,
'           picture, summary and subreport objects.
'
'           These report objects can either be added at runtime to
'           an existing report created in the VB designer, or (as
'           shown in this sample) a blank report can be created and
'           objects added to that blank report.  All properties that
'           are normally available for each object at runtime are
'           also available for these objects.
'
'           This sample shows how to create text, database field,
'           unbound field, summary, picture box and line objects.  For
'           a larger demonstration of advanced uses of the Creation API,
'           please see the Pro Athlete Salaries sample located in the
'           Crystal Reports 8.0 VB samples directory.
'
Option Explicit

Dim m_Proj As CRAXDRT.Application               ' Create a new COM instance of the Crystal Reports
                                                ' Run-time application
Dim m_Report As CRAXDRT.Report                  ' The dynamically generated report shell

Dim cnn1 As ADODB.Connection
Dim datcmd1 As ADODB.Command
' *************************************************************
' Add a text, line, box, and picture object to the header of the report
'
Private Sub AddHeaderObjects()
    Dim txtObj As TextObject
    
    m_Report.Sections(1).AddBoxObject 0, 0, 10000, 1000
    m_Report.Sections(2).AddPictureObject App.Path & "\logo.bmp", 0, 100
    m_Report.Sections(2).Height = m_Report.Sections(2).Height + 300
    Set txtObj = m_Report.Sections(1).AddTextObject("Mountain Bike Sales by Store", 1700, 0)
    ' Format the text object to look like a title
    With txtObj
        .Font.Size = 24
        .Font.Bold = True
        .TextColor = &H808000
        .Height = 800
        .Width = 6500
    End With
End Sub

' *************************************************************
' Add a database field, unbound field, and line object to the
' details section.  Also add a summary object that totals
' "Last Year's Sales".
'
Private Sub AddDetailObjects()
    Dim lnObj As LineObject
    Dim dbfldObj As FieldObject
    Dim UnboundfldObj As FieldObject
    Dim sumObj As FieldObject
    
    With m_Report.Sections(3)
        ' Add two new database field object to the report and set the field
        ' objects to use the new data source.
        .AddFieldObject "{ado.Customer Name}", 0, 0
        Set dbfldObj = .AddFieldObject("{ado.Last Year's Sales}", 6000, 0)
        ' Add an unbound field and a line object
        Set UnboundfldObj = .AddUnboundFieldObject(crStringField, 3000, 0)
        Set lnObj = .AddLineObject(2600, 0, 2600, 10)
        
        ' Format the various objects to be more visually appealing
        dbfldObj.TextColor = vbRed
        UnboundfldObj.SetUnboundFieldSource "{ado.City}"
        lnObj.LineThickness = 2
        lnObj.ExtendToBottomOfSection = True
    End With
    
    ' Add the summary object, and format it to align with the "Last Year's Sales"
    ' column on the report
    Set sumObj = m_Report.Sections(1).AddSummaryFieldObject(m_Report.Sections(3).ReportObjects(2).Field.Name, _
                                        crSTSum, 6000, 750)
    sumObj.Width = 2600
    sumObj.HorAlignment = crLeftAlign
End Sub

' *************************************************************
' Add a runtime datasource to the report
'
Private Sub AddDataSource()
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
    m_Report.PaperSize = crPaperA4Small
End Sub

' *************************************************************
' Create a new report, add a data source and runtime report
' objects, and then set the viewer to display this report
'
Private Sub InitReport()
    Set m_Proj = New CRAXDRT.Application
    Set m_Report = m_Proj.NewReport

    AddDataSource
    AddHeaderObjects
    AddDetailObjects
    
    ' Finally, view the report in the ActiveX Viewer
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ' Add the Data Source to the report
    InitReport
    
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

' *************************************************************
' Force the viewer to only display the report at 100%
'
Private Sub CRViewer1_ZoomLevelChanged(ByVal ZoomLevel As Integer)
    If ZoomLevel <> 100 Then CRViewer1.Zoom 100
End Sub

