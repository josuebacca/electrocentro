VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmployee 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee"
   ClientHeight    =   4575
   ClientLeft      =   330
   ClientTop       =   780
   ClientWidth     =   7725
   ForeColor       =   &H00404040&
   Icon            =   "Employee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Employee"
   Begin VB.TextBox txtDummy 
      DataField       =   "Position"
      DataSource      =   "AdoData1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3780
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "ODBC;DSN=Xtreme sample database"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employee"
      Top             =   4185
      Width           =   7455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5535
      TabIndex        =   25
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "&Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4455
      TabIndex        =   24
      Tag             =   "&Grid"
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3375
      TabIndex        =   23
      Tag             =   "&Update"
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   22
      Tag             =   "&Refresh"
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   21
      Tag             =   "&Delete"
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   20
      Tag             =   "&Add"
      Top             =   3870
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notes"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2880
      Width           =   4305
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Extension"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2280
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Home Phone"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1960
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hire Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   1640
      Width           =   1680
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Birth Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MMMM d, yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   1680
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Position"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1000
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "First Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   5
      Top             =   680
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Last Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   3
      Top             =   360
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Employee ID"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   40
      Width           =   1680
   End
   Begin MSAdodcLib.Adodc AdoData1 
      Height          =   330
      Left            =   6330
      Top             =   3870
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Xtreme sample database"
      OLEDBString     =   "DSN=Xtreme sample database"
      OLEDBFile       =   ""
      DataSourceName  =   "Xtreme sample data"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Employee"
      Caption         =   "Data"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Tag             =   "Notes:"
      Top             =   2625
      Width           =   1815
   End
   Begin VB.OLE oleField 
      DataField       =   "Photo"
      DataSource      =   "Data1"
      Height          =   3375
      Left            =   4650
      OLETypeAllowed  =   1  'Embedded
      SizeMode        =   3  'Zoom
      TabIndex        =   17
      Top             =   360
      Width           =   2910
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Photo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4650
      TabIndex        =   16
      Tag             =   "Photo:"
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Extension:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Tag             =   "Extension:"
      Top             =   2295
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Tag             =   "Home Phone:"
      Top             =   1980
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Hire Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Tag             =   "Hire Date:"
      Top             =   1665
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Tag             =   "Birth Date:"
      Top             =   1335
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Tag             =   "Position:"
      Top             =   1020
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Tag             =   "First Name:"
      Top             =   705
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "Last Name:"
      Top             =   375
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Employee ID:"
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Handle the adding/deleting of Employees.  This form,
'          along with frmProductType, is a bit different because
'          it has to support OLE objects in the database.  ADO
'          doesn't seem to do this very well yet, so this particular
'          form uses both an ADO and DAO Data Control
'
Option Explicit

' *************************************************************
' Add a new Employee
'
Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub

' *************************************************************
' Delete an Employee
'
Private Sub cmdDelete_Click()
    ' Ensure that we have records to work with
    If Data1.Recordset.RecordCount = 0 Then Exit Sub
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
End Sub

' *************************************************************
' Refresh the Recordset -- only really needed for when multiple
' users are connected to the same data source.
'
Private Sub cmdRefresh_Click()
    Data1.Refresh
End Sub

' *************************************************************
' Update the user's changes (no validation, though)
'
Private Sub cmdUpdate_Click()
    Data1.Recordset.Edit
    Data1.Recordset.Update
End Sub

' *************************************************************
' Since the Data Grid in frmDataGrid expects an ADO data source,
' send it that data source and not the DAO data source
'
Private Sub cmdGrid_Click()
    g_RecordSource = AdoData1.RecordSource
    Set g_RecordSet = AdoData1.Recordset
    frmDataGrid.Show vbModal
End Sub

' *************************************************************
' Handle errors and allow the user to continue
'
Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    MsgBox "Error: " & Error$(DataErr)
    Response = 0
End Sub

' *************************************************************
' When the user changes records, update the caption on the data
' control
'
Private Sub Data1_Reposition()
    Data1.Caption = Data1.Recordset.AbsolutePosition + 1
End Sub

' *************************************************************
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Initialize the Data control to the correct database (Xtreme.mdb),
' which must be located in the same directory as this application
'
Private Sub Form_Load()
    g_FormCaption = Me.Caption
    Dither Me
End Sub

' *************************************************************
' When the form is about to close, we've done all our work
' Make the sure the mouse pointer provides the correct feedback
'
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    g_StartForm.Enabled = True
    g_StartForm.ZOrder 0
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

' *************************************************************
' Put  to get data into an empty ole control
' and have it saved back to the table

Private Sub oleField_DblClick()
    oleField.InsertObjDlg
End Sub

