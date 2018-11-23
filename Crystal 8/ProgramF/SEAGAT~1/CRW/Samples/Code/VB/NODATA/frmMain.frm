VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No Data in Report Event Sample"
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
   Begin VB.CommandButton cmdFilter 
      Caption         =   "&Apply Data Filter"
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
      Left            =   75
      TabIndex        =   3
      Top             =   6855
      Width           =   1935
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
' Purpose:  Demonstrate new "No Data" Event (New to Crystal Reports 8.0)
'
'           The "No Data" event is triggered when the data set
'           in the report is empty.  In this sample, a filter is
'           applied to the report that causes the report to have
'           no data.
'
'           A message box is then displayed which allows
'           the user the option to view the report without any data
'           or just to view a blank page
'
Option Explicit

Dim m_Report As New CrystalReport1

' *************************************************************
' Toggle a filter on the report that causes the report to have
' an empty data set.
'
Private Sub cmdFilter_Click()
    If cmdFilter.Caption = "&Apply Data Filter" Then
        ' Add the Filter
        m_Report.RecordSelectionFormula = "{Product.Product Name} = 'Super Xtreme Gloves'"
        cmdFilter.Caption = "&Remove Data Filter"
    Else
        ' Remove the Filter
        m_Report.RecordSelectionFormula = ""
        cmdFilter.Caption = "&Apply Data Filter"
    End If
    CRViewer1.Refresh
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
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

