VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmDisplay 
   Caption         =   "Display Form"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   465
      Left            =   5985
      TabIndex        =   2
      Top             =   7260
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   465
      Left            =   7830
      TabIndex        =   1
      Top             =   7260
      Width           =   1695
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewReport 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Display the Report selected by the user in the viewer
'          embedded in the window
'
Option Explicit

Public m_Report As Report                       ' The report to be embedded in the viewer

Private Sub Form_Load()
    Me.Caption = g_ReportName
    CRViewReport.DisplayGroupTree = False
    CRViewReport.EnableGroupTree = False
    ' Choose the report based on what has been set in frmStartup
    ' The constants are defined in Utilities.bas
    Select Case g_ReportName
    Case vbBalanceStatement
        Set m_Report = New CRBalanceStatement
    Case vbIncomeStatement
        Set m_Report = New CRIncomeStatement
    Case vbMailLabels
        Set m_Report = New CRMailLabels
    Case vbWorldSales
        Set m_Report = New CRWorldWideSales
    Case vbProductCatalog
        CRViewReport.DisplayGroupTree = True
        CRViewReport.EnableGroupTree = True
        Set m_Report = New CRProductCatalog
    Case vbEmployeeProfiles
        CRViewReport.DisplayGroupTree = True
        CRViewReport.EnableGroupTree = True
        Set m_Report = New CREmployeeProfiles
    End Select
                
    ' Display the report
    With CRViewReport
        .ReportSource = m_Report
        .ViewReport
        .Zoom 100
    End With
    Me.WindowState = vbMaximized
End Sub

' *************************************************************
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

' *************************************************************
' When the form is resized, reposition the controls on the form
' to match the new form size
'
Private Sub Form_Resize()
    cmdClose.Top = ScaleHeight - 15 - cmdClose.Height
    cmdClose.Left = ScaleWidth - 15 - cmdClose.Width
    
    cmdOptions.Top = cmdClose.Top
    cmdOptions.Left = cmdClose.Left - cmdOptions.Width - 60
    
    CRViewReport.Width = ScaleWidth - 15
    CRViewReport.Height = ScaleHeight - cmdClose.Height - 60
End Sub

' *************************************************************
Private Sub Form_Unload(Cancel As Integer)
    g_StartForm.Enabled = True
    g_StartForm.ZOrder 0
End Sub

' *************************************************************
' Show a few of the available options for the report
'
Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub

' *************************************************************
' A publicly exposed method that is called from frmOptions, this
' allows the user to dynamically change some of the viewer options
'
Public Sub SetOptions(frm As Form)
    With CRViewReport
        .EnableDrillDown = frm.chkWindowDrill
        .DisplayToolbar = frm.chkWindowToolbar
        .EnableGroupTree = frm.chkWindowGroupTree
        .EnableStopButton = frm.chkWindowCancel
        .EnableCloseButton = frm.chkWindowClose
        .EnableNavigationControls = frm.chkWindowNavigation
        .EnablePrintButton = frm.chkWindowPrint
        .EnableProgressControl = frm.chkWindowProgress
        .EnableRefreshButton = frm.chkWindowRefresh
        .EnableSearchControl = frm.chkWindowSearch
        .EnableZoomControl = frm.chkWindowZoom
    End With
End Sub

' *************************************************************
' A publicly exposed method that is called from frmOptions, this
' allows the user to dynamically retrieve some of the viewer options

Public Sub GetOptions(frm As Form)
    With CRViewReport
        frm.chkWindowCancel.Value = Abs(CInt(.EnableStopButton))
        frm.chkWindowClose = Abs(CInt(.EnableCloseButton))
        frm.chkWindowDrill = Abs(CInt(.EnableDrillDown))
        frm.chkWindowToolbar = Abs(CInt(.DisplayToolbar))
        frm.chkWindowGroupTree = Abs(CInt(.EnableGroupTree))
        frm.chkWindowNavigation = Abs(CInt(.EnableNavigationControls))
        frm.chkWindowPrint = Abs(CInt(.EnablePrintButton))
        frm.chkWindowProgress = Abs(CInt(.EnableProgressControl))
        frm.chkWindowRefresh = Abs(CInt(.EnableRefreshButton))
        frm.chkWindowSearch = Abs(CInt(.EnableSearchControl))
        frm.chkWindowZoom = Abs(CInt(.EnableZoomControl))
    End With
End Sub

' *************************************************************
' Force the viewer to
Private Sub CRViewReport_ZoomLevelChanged(ByVal ZoomLevel As Integer)
    If ZoomLevel <> 100 Then CRViewReport.Zoom 100
End Sub

