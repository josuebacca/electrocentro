VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Options"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   4395
      TabIndex        =   31
      Top             =   3855
      Width           =   1365
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   2925
      TabIndex        =   30
      Top             =   3855
      Width           =   1365
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3765
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   6641
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Viewer"
      TabPicture(0)   =   "frmOptions.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkWindowZoom"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkWindowRefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkWindowNavigation"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkWindowToolbar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkWindowCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkWindowSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkWindowProgress"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkWindowPrint"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkWindowGroupTree"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkWindowClose"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkWindowDrill"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Tables"
      TabPicture(1)   =   "frmOptions.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTablesPassword"
      Tab(1).Control(1)=   "lblTablesUser"
      Tab(1).Control(2)=   "lblTablesDatabase"
      Tab(1).Control(3)=   "lblTablesServer"
      Tab(1).Control(4)=   "lblTablesDLL"
      Tab(1).Control(5)=   "lblTablesLocation"
      Tab(1).Control(6)=   "lblTablesDesc"
      Tab(1).Control(7)=   "lblTablesName"
      Tab(1).Control(8)=   "txtTablesPassword"
      Tab(1).Control(9)=   "txtTablesUser"
      Tab(1).Control(10)=   "txtTablesDatabase"
      Tab(1).Control(11)=   "txtTablesServer"
      Tab(1).Control(12)=   "txtTablesDLL"
      Tab(1).Control(13)=   "txtTablesLocation"
      Tab(1).Control(14)=   "txtTablesDesc"
      Tab(1).Control(15)=   "txtTablesName"
      Tab(1).Control(16)=   "cmdTablesTest"
      Tab(1).Control(17)=   "lstTables"
      Tab(1).ControlCount=   18
      Begin VB.ListBox lstTables 
         Height          =   1230
         Left            =   -74865
         TabIndex        =   21
         Top             =   525
         Width           =   1965
      End
      Begin VB.CommandButton cmdTablesTest 
         Caption         =   "Test &Connectivity"
         Height          =   390
         Left            =   -74865
         TabIndex        =   20
         Top             =   2130
         Width           =   1965
      End
      Begin VB.TextBox txtTablesName 
         Height          =   300
         Left            =   -71370
         TabIndex        =   19
         Top             =   480
         Width           =   2010
      End
      Begin VB.TextBox txtTablesDesc 
         Height          =   300
         Left            =   -71370
         TabIndex        =   18
         Top             =   885
         Width           =   2010
      End
      Begin VB.TextBox txtTablesLocation 
         Height          =   300
         Left            =   -71355
         TabIndex        =   17
         Top             =   1290
         Width           =   2010
      End
      Begin VB.TextBox txtTablesDLL 
         Height          =   300
         Left            =   -71370
         TabIndex        =   16
         Top             =   1680
         Width           =   2010
      End
      Begin VB.TextBox txtTablesServer 
         Height          =   300
         Left            =   -71370
         TabIndex        =   15
         Top             =   2085
         Width           =   2010
      End
      Begin VB.TextBox txtTablesDatabase 
         Height          =   300
         Left            =   -71370
         TabIndex        =   14
         Top             =   2490
         Width           =   2010
      End
      Begin VB.TextBox txtTablesUser 
         Height          =   300
         Left            =   -71370
         TabIndex        =   13
         Top             =   2895
         Width           =   2010
      End
      Begin VB.TextBox txtTablesPassword 
         Height          =   300
         Left            =   -71370
         TabIndex        =   12
         Top             =   3315
         Width           =   2010
      End
      Begin VB.CheckBox chkWindowDrill 
         Caption         =   "&Drill Down"
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   1098
         Width           =   2025
      End
      Begin VB.CheckBox chkWindowClose 
         Caption         =   "C&lose Button"
         Height          =   195
         Left            =   3105
         TabIndex        =   10
         Top             =   1110
         Width           =   2025
      End
      Begin VB.CheckBox chkWindowGroupTree 
         Caption         =   "&Group Tree"
         Height          =   195
         Left            =   450
         TabIndex        =   9
         Top             =   1626
         Width           =   2025
      End
      Begin VB.CheckBox chkWindowPrint 
         Caption         =   "&Print Button"
         Height          =   195
         Left            =   450
         TabIndex        =   8
         Top             =   2154
         Width           =   2025
      End
      Begin VB.CheckBox chkWindowProgress 
         Caption         =   "Progress Con&trol"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   2682
         Width           =   2025
      End
      Begin VB.CheckBox chkWindowSearch 
         Caption         =   "&Search Button"
         Height          =   195
         Left            =   450
         TabIndex        =   6
         Top             =   3210
         Width           =   1365
      End
      Begin VB.CheckBox chkWindowCancel 
         Caption         =   "&Cancel Button"
         Height          =   195
         Left            =   3105
         TabIndex        =   5
         Top             =   585
         Width           =   1800
      End
      Begin VB.CheckBox chkWindowToolbar 
         Caption         =   "&Toolbar"
         Height          =   195
         Left            =   450
         TabIndex        =   4
         Top             =   570
         Width           =   1800
      End
      Begin VB.CheckBox chkWindowNavigation 
         Caption         =   "&Navigation Controls"
         Height          =   195
         Left            =   3105
         TabIndex        =   3
         Top             =   1620
         Width           =   1800
      End
      Begin VB.CheckBox chkWindowRefresh 
         Caption         =   "&Refresh Button"
         Height          =   195
         Left            =   3105
         TabIndex        =   2
         Top             =   2145
         Width           =   1800
      End
      Begin VB.CheckBox chkWindowZoom 
         Caption         =   "&Zoom Controls"
         Height          =   195
         Left            =   3105
         TabIndex        =   1
         Top             =   2670
         Width           =   1800
      End
      Begin VB.Label lblTablesName 
         Caption         =   "Name:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   29
         Top             =   555
         Width           =   1305
      End
      Begin VB.Label lblTablesDesc 
         Caption         =   "Description:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   28
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblTablesLocation 
         Caption         =   "Location:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   27
         Top             =   1365
         Width           =   1305
      End
      Begin VB.Label lblTablesDLL 
         Caption         =   "DLL Name:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   26
         Top             =   1755
         Width           =   1305
      End
      Begin VB.Label lblTablesServer 
         Caption         =   "Server Name:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   25
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label lblTablesDatabase 
         Caption         =   "Database Name:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   24
         Top             =   2565
         Width           =   1305
      End
      Begin VB.Label lblTablesUser 
         Caption         =   "User ID:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   23
         Top             =   2970
         Width           =   1305
      End
      Begin VB.Label lblTablesPassword 
         Caption         =   "Password:"
         Height          =   270
         Left            =   -72750
         TabIndex        =   22
         Top             =   3375
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Provide a form for demonstrating one way to change
'           and display a few of the properties available from
'           the shown report in frmDisplay
'
Option Explicit

Private m_crwTable As DatabaseTable         ' Used to retrieve database info about the report

' *************************************************************
' Apply the user's changes without closing the form
'
Private Sub cmdApply_Click()
    frmDisplay.SetOptions Me
    cmdApply.Enabled = False
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

' *************************************************************
' Test that the database is currently accessible
'
Private Sub cmdTablesTest_Click()
    If Not m_crwTable Is Nothing Then
        If m_crwTable.TestConnectivity Then
            MsgBox "Connected!", vbInformation
        Else
            MsgBox "Not connected!", vbInformation
        End If
    End If
End Sub

' *************************************************************
' Initialize the data in the form, retrieving information about
' each table in the report as well as some of the viewer options
'
Private Sub Form_Load()
    Dim intTable As Integer
        
    frmDisplay.GetOptions Me
    
    With frmDisplay.m_Report.Database.Tables
        For intTable = 1 To .Count
            lstTables.AddItem .Item(intTable).Name
        Next intTable
    End With
    If lstTables.ListCount > 0 Then lstTables.ListIndex = 0
    cmdApply.Enabled = False
End Sub

' *************************************************************
' Display the database information about the currently selected
' table
'
Private Sub lstTables_Click()
    Set m_crwTable = frmDisplay.m_Report.Database.Tables.Item(lstTables.ListIndex + 1)
    With m_crwTable
        Me.txtTablesDLL.Text = .DllName
        Me.txtTablesServer.Text = .LogOnServerName
        Me.txtTablesDatabase.Text = .LogOnDatabaseName
        Me.txtTablesUser.Text = .LogOnUserID
        Me.txtTablesName.Text = .Name
        Me.txtTablesDesc.Text = .DecriptiveName
        Me.txtTablesLocation.Text = .Location
    End With
End Sub

' *************************************************************
' Enable the "Apply" Button when then user changes something
'
Private Sub chkWindowDrill_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowGroupTree_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowToolbar_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowCancel_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowClose_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowNavigation_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowPrint_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowProgress_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowRefresh_Click()
    cmdApply.Enabled = True
End Sub

' *************************************************************
Private Sub chkWindowZoom_Click()
    cmdApply.Enabled = True
End Sub

