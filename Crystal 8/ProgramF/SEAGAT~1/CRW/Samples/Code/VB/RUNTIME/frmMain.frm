VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5C9EA131-127A-11D1-BFB4-00A0C936E6F9}#1.0#0"; "cselexpt.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewer Runtime Options Sample"
   ClientHeight    =   7335
   ClientLeft      =   2100
   ClientTop       =   2250
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10020
   Begin VB.Frame Frame2 
      Caption         =   "Display Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   6735
      TabIndex        =   21
      Top             =   4800
      Width           =   3210
      Begin VB.CheckBox chkDisplayTabs 
         Caption         =   "&Tabs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   26
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkDisplayToolbar 
         Caption         =   "&Toolbar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   25
         Top             =   1590
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkDisplayGroupTree 
         Caption         =   "&Group Tree"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   24
         Top             =   930
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkDisplayBackgroundEdge 
         Caption         =   "&Background Edge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   23
         Top             =   285
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkDisplayBorder 
         Caption         =   "&Border"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   435
         TabIndex        =   22
         Top             =   615
         Value           =   1  'Checked
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enable Properties"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2430
      Left            =   75
      TabIndex        =   3
      Top             =   4800
      Width           =   6420
      Begin VB.CheckBox chkAnimationControl 
         Caption         =   "&Animation Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   20
         Top             =   360
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkExportButton 
         Caption         =   "&Export Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   1335
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkHelpButton 
         Caption         =   "&Help Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   18
         Top             =   1980
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkNavigationControls 
         Caption         =   "&Navigation Controls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   17
         Top             =   375
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkPopupMenu 
         Caption         =   "&Pop-up Menu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   16
         Top             =   705
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkSearchExpertButton 
         Caption         =   "&Search Expert Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkSelectExpertButton 
         Caption         =   "&Select Expert Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   14
         Top             =   690
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkEnableZoomControl 
         Caption         =   "&Zoom Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   13
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkEnableRefreshButton 
         Caption         =   "&Refresh Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   12
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkEnableToolbar 
         Caption         =   "&Toolbar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   11
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkEnableStopButton 
         Caption         =   "&Stop Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   10
         Top             =   1020
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkEnableSearchControl 
         Caption         =   "&Search Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   9
         Top             =   2010
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkEnableProgressControl 
         Caption         =   "Progress Con&trol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   8
         Top             =   1350
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkEnablePrintButton 
         Caption         =   "&Print Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   7
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkEnableGroupTree 
         Caption         =   "&Group Tree"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   1650
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkEnableCloseButton 
         Caption         =   "C&lose Button"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   690
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox chkEnableDrillDown 
         Caption         =   "&Drill Down"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   1005
         Value           =   1  'Checked
         Width           =   2010
      End
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
      Left            =   8565
      TabIndex        =   1
      Top             =   6825
      Width           =   1380
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9915
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
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin CSELEXPTLib.VSelExpert VSelExpert1 
      Height          =   270
      Left            =   6495
      TabIndex        =   27
      Top             =   6870
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   476
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate how to manipulate the Crystal Reports
'           ActiveX Viewer at Runtime.  All checkboxes on the
'           main form correspond to properties on the viewer
'           that can be changed.
'
'           In this sample, there are more than 21 viewer properties
'           that are manipulated.  Note that to enable the
'           Select Expert and Search Expert, a reference to the
'           Crystal Select Expert OLE control module must be added
'           and the Select Expert Control placed on the form.
'
'           The control is essentially invisible on the form and is
'           located immediately to the left of the "About Sample"
'           Command Button in this sample.
'

Dim m_Report As New CrystalReport1
Dim IgnoreEvent As Boolean             ' Used to prevent undesired events from
                                       ' triggering when toggling the Group Tree checkbox
Option Explicit

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    IgnoreEvent = False
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
' Check or Uncheck the Group Tree Checkbox when the Group Tree
' is toggled on the viewer
'
Private Sub CRViewer1_GroupTreeButtonClicked(ByVal Visible As Boolean)
    ' Ensure that the chkDisplayGroupTree event doesn't trigger
    IgnoreEvent = True
    chkDisplayGroupTree.Value = Abs(CRViewer1.DisplayGroupTree)
End Sub

' *************************************************************
' All of the following event handlers use the check boxes on the
' main form to enable and disable methods/properties on the
' Crystal Reports ActiveX Viewer.
'
Private Sub chkAnimationControl_Click()
    CRViewer1.EnableAnimationCtrl = Not CRViewer1.EnableAnimationCtrl
End Sub

Private Sub chkDisplayBackgroundEdge_Click()
    CRViewer1.DisplayBackgroundEdge = Not CRViewer1.DisplayBackgroundEdge
End Sub

Private Sub chkDisplayBorder_Click()
    CRViewer1.DisplayBorder = Not CRViewer1.DisplayBorder
End Sub

Private Sub chkDisplayGroupTree_Click()
    If IgnoreEvent = False Then _
        CRViewer1.DisplayGroupTree = Not CRViewer1.DisplayGroupTree
    IgnoreEvent = False
End Sub

Private Sub chkDisplayTabs_Click()
    CRViewer1.DisplayTabs = Not CRViewer1.DisplayTabs
End Sub

Private Sub chkDisplayToolbar_Click()
    If IgnoreEvent = False Then _
        CRViewer1.DisplayToolbar = Not CRViewer1.DisplayToolbar
    IgnoreEvent = False
End Sub

Private Sub chkEnableCloseButton_Click()
    CRViewer1.EnableCloseButton = Not CRViewer1.EnableCloseButton
End Sub

Private Sub chkEnableDrillDown_Click()
    CRViewer1.EnableDrillDown = Not CRViewer1.EnableDrillDown
End Sub

Private Sub chkEnableGroupTree_Click()
    CRViewer1.EnableGroupTree = Not CRViewer1.EnableGroupTree
End Sub

Private Sub chkEnablePrintButton_Click()
    CRViewer1.EnablePrintButton = Not CRViewer1.EnablePrintButton
End Sub

Private Sub chkEnableProgressControl_Click()
    CRViewer1.EnableProgressControl = Not CRViewer1.EnableProgressControl
End Sub

Private Sub chkEnableRefreshButton_Click()
    CRViewer1.EnableRefreshButton = Not CRViewer1.EnableRefreshButton
End Sub

Private Sub chkEnableSearchControl_Click()
    CRViewer1.EnableSearchControl = Not CRViewer1.EnableSearchControl
End Sub

Private Sub chkSearchExpertButton_Click()
    CRViewer1.EnableSearchExpertButton = Not CRViewer1.EnableSearchExpertButton
End Sub

Private Sub chkSelectExpertButton_Click()
    CRViewer1.EnableSelectExpertButton = Not CRViewer1.EnableSelectExpertButton
End Sub

Private Sub chkEnableStopButton_Click()
    CRViewer1.EnableStopButton = Not CRViewer1.EnableStopButton
End Sub

Private Sub chkEnableToolbar_Click()
    CRViewer1.EnableToolbar = Not CRViewer1.EnableToolbar
End Sub

Private Sub chkEnableZoomControl_Click()
    CRViewer1.EnableZoomControl = Not CRViewer1.EnableZoomControl
End Sub

Private Sub chkExportButton_Click()
    CRViewer1.EnableExportButton = Not CRViewer1.EnableExportButton
End Sub

Private Sub chkHelpButton_Click()
    CRViewer1.EnableHelpButton = Not CRViewer1.EnableHelpButton
End Sub

Private Sub chkNavigationControls_Click()
    CRViewer1.EnableNavigationControls = Not CRViewer1.EnableNavigationControls
End Sub

Private Sub chkPopupMenu_Click()
    CRViewer1.EnablePopupMenu = Not CRViewer1.EnablePopupMenu
End Sub


