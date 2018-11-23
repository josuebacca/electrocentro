VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Variables Sample"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
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
      TabIndex        =   4
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
   Begin VB.Label Label1 
      Caption         =   "First Digits of Each Green Highlighted Value:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   7185
      Width           =   3990
   End
   Begin VB.Label lblFirstDigits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   4140
      TabIndex        =   5
      Top             =   7185
      Width           =   2565
   End
   Begin VB.Label lblTotal 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3165
      TabIndex        =   3
      Top             =   6870
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Total of Blue Highlighted Values:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   6870
      Width           =   3030
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate how to use Report Variables (new to Crystal Reports 8.0)
'           See the code CRReportVariables for a detailed explanation
'           of the sample

Option Explicit

Dim m_Report As New CRReportVariables

Dim m_Connection As ADODB.Connection
Dim adoRS As ADODB.Recordset

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
    Dim SQL As String
    Dim strConnect As String
   
    ' Create and bind the ADO Recordset object
    Set m_Connection = New ADODB.Connection
    Set adoRS = New ADODB.Recordset

    ' Open the connection
    strConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\ReportVariables.mdb;Mode=Read"
    m_Connection.Open strConnect
    SQL = "Select * from tblSampleData"
    adoRS.Open SQL, m_Connection, adOpenDynamic, adLockBatchOptimistic
    m_Report.Database.SetDataSource adoRS
    
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

