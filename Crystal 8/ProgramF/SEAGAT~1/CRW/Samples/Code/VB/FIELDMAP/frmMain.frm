VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Field Mapping Sample"
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
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate how to manipulate Field Mapping (new to
'           Crystal Reports 8.0)
'
'           Field mapping allows the user to change the name of
'           a database field that a report field points to if
'           that field no longer exists or has been renamed in
'           the data source.
'
'           In this sample, the Report_FieldMapping event in the
'           code for CrystalReport1 is used to programmatically
'           change the mapping of the report fields to a different
'           database field.
'

Option Explicit

Dim m_Report As New CrystalReport1

' *************************************************************
' Load the Report in the viewer and enable fieldmapping
'
Private Sub Form_Load()
    Dim ReportDB As Database
    Set ReportDB = m_Report.Database
    
    ' Explanation of FieldMappingType:
    ' 0 - crAutoFieldMapping - Remove all report fields which do not have the same
    '                          name in the new database
    ' 1 - crPromptFieldMapping - Display a dialog box to get the user to manually
    '                          map the fields
    ' 2 - crEventFieldMapping - Cause the Report_FieldMapping event to be triggered
    '                          (see event code for CrystalReport1 in this sample)
    m_Report.FieldMappingType = crEventFieldMapping
                   
    ' Verifying the database or using SetLocation will trigger the Field Mapping event
    ReportDB.Verify
   
    ' Display our report in the viewer
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
End Sub

' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

