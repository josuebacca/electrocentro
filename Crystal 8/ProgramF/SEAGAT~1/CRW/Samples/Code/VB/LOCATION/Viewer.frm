VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmViewer 
   Caption         =   "Change Runtime OLE Location Sample"
   ClientHeight    =   8295
   ClientLeft      =   1710
   ClientTop       =   1935
   ClientWidth     =   10245
   Icon            =   "Viewer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   10245
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About Sample"
      Height          =   465
      Left            =   6750
      TabIndex        =   2
      Top             =   7770
      Width           =   1665
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   465
      Left            =   8535
      TabIndex        =   1
      Top             =   7770
      Width           =   1665
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7635
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10155
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
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Load the report into the viewer and then close the form
'          when the user is finished.  Please see the code for
'          crOLEReport for an explanation of how the sample
'          works.  This sample assumes that you have Microsoft Word
'          and Microsoft Excel installed.
'

Option Explicit

Dim m_Report As New crOLEReport

' *************************************************************
' Place the report into the viewer
'
Private Sub Form_Load()
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
End Sub

' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

