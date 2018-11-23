VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRViewer.dll"
Begin VB.Form frmPreview 
   Caption         =   "Inventory Report"
   ClientHeight    =   8220
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   9675
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   9675
   Begin VB.CommandButton cmdClose 
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
      Height          =   465
      Left            =   8145
      TabIndex        =   1
      Top             =   7725
      Width           =   1485
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   9645
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Display an inventory report in the Report Viewer
'

Option Explicit

Dim Report As New dsrInventory

' *************************************************************
' Initialize the form to display the inventory report
'
Private Sub Form_Load()
    CRViewer1.ReportSource = Report
    CRViewer1.ViewReport
End Sub

' *************************************************************
' When the form is resized, reposition the controls on the form
' to match the new form size
'
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    
    cmdClose.Top = ScaleHeight - 15 - cmdClose.Height
    cmdClose.Left = ScaleWidth - 15 - cmdClose.Width
    
    CRViewer1.Height = ScaleHeight - cmdClose.Height - 30
    CRViewer1.Width = ScaleWidth - 15
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

' *************************************************************
' Show the About Form
'
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
