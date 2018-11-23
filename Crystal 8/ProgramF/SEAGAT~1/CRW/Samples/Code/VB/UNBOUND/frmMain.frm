VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unbound Fields Sample"
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
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate one way of using Unbound Fields (new to
'           Crystal Reports 8.0).
'
'           In this sample, there are five unbound fields created
'           in the report.  Notice that there is no data source
'           added to the report at design time at all -- the data
'           source is created and added to the report at runtime.
'           Unbound fields can also be used as formulas, although
'           this sample does not demonstrate this use.
'
'           Unbound fields allow for great flexibility, since you can
'           change easily change the data source for an unbound field
'           at runtime.
'

Dim m_Report As New crUnboundFields

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ' Add the Data Source to the report
    m_Report.Database.AddOLEDBSource DataEnvironment1.Connection1, "Orders"
    
    ' Bind the five unbound fields to the newly added data source
    With m_Report
        .UnboundString1.SetUnboundFieldSource "{Orders.Ship Via}"
        .UnboundCurrency1.SetUnboundFieldSource "{Orders.Order Amount}"
        .UnboundBoolean1.SetUnboundFieldSource "{Orders.Shipped}"
        .UnboundNumber1.SetUnboundFieldSource "{Orders.Customer ID}"
        .UnboundDate1.SetUnboundFieldSource "{Orders.Order Date}"
    End With
    
    ' Do some runtime formatting of the number and date unbound fields
    ' to make the report look more appealing
    m_Report.UnboundNumber1.DecimalPlaces = 0
    With m_Report.UnboundDate1
        .HourType = crNoHour
        .MinuteType = crNoMinute
        .SecondType = crNumericNoSecond
        .AmString = ""
        .MonthType = crShortMonth
        .YearType = crLongYear
        .DateFirstSeparator = " "
        .DateSecondSeparator = ", "
    End With
 
    ' View the report
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
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

