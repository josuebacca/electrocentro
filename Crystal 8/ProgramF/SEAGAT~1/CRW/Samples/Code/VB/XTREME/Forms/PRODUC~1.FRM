VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProductType 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Type"
   ClientHeight    =   2895
   ClientLeft      =   390
   ClientTop       =   1755
   ClientWidth     =   7650
   Icon            =   "ProductType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Product Type"
   Begin VB.TextBox txtDummy 
      DataField       =   "Product Type ID"
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
      Left            =   4020
      TabIndex        =   14
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
      Height          =   375
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "[Product Type]"
      Top             =   2505
      Width           =   7470
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
      Left            =   5520
      TabIndex        =   13
      Top             =   2175
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
      Left            =   4440
      TabIndex        =   12
      Tag             =   "&Grid"
      Top             =   2175
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
      Left            =   3360
      TabIndex        =   11
      Tag             =   "&Update"
      Top             =   2175
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
      TabIndex        =   10
      Tag             =   "&Refresh"
      Top             =   2175
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
      TabIndex        =   9
      Tag             =   "&Delete"
      Top             =   2175
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
      TabIndex        =   8
      Tag             =   "&Add"
      Top             =   2175
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Description"
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
      Height          =   1155
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   915
      Width           =   4440
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Product Type Name"
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   2880
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Product Type ID"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   40
      Width           =   1440
   End
   Begin MSAdodcLib.Adodc AdoData1 
      Height          =   330
      Left            =   6525
      Top             =   2175
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      RecordSource    =   "[Product Type]"
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
   Begin VB.OLE oleField 
      DataField       =   "Picture"
      DataSource      =   "Data1"
      Height          =   1710
      Left            =   4665
      OLETypeAllowed  =   1  'Embedded
      SizeMode        =   3  'Zoom
      TabIndex        =   7
      Top             =   360
      Width           =   2880
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture:"
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
      Left            =   4665
      TabIndex        =   6
      Tag             =   "Picture:"
      Top             =   60
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Tag             =   "Description:"
      Top             =   690
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type Name:"
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
      Tag             =   "Product Type Name:"
      Top             =   375
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type ID:"
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
      Tag             =   "Product Type ID:"
      Top             =   60
      Width           =   1560
   End
End
Attribute VB_Name = "frmProductType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Handle the adding/deleting of Product Types.  This form,
'          along with frmProductType, is a bit different because
'          it has to support OLE objects in the database.  ADO
'          doesn't seem to do this very well yet, so this particular
'          form uses both an ADO and DAO Data Control
'
Option Explicit

' *************************************************************
' Add a new Product Type
'
Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub

' *************************************************************
' Delete an Product Type
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
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Put  to get data into an empty ole control
' and have it saved back to the table

Private Sub oleField_DblClick()
    oleField.InsertObjDlg
End Sub


