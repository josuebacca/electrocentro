VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDataGrid 
   ClientHeight    =   4080
   ClientLeft      =   615
   ClientTop       =   795
   ClientWidth     =   7230
   Icon            =   "DataGrid.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleMode       =   0  'User
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7230
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7230
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Tag             =   "&Close"
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "R&efresh"
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Tag             =   "&Refresh"
         Top             =   0
         Width           =   1800
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   405
      Left            =   45
      Top             =   3585
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   714
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   1
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   3165
      Left            =   -15
      TabIndex        =   3
      Top             =   345
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5583
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Called from any one of the data entry forms, this data
'          grid is intended to make data entry easier, and also show
'          the clear advantages of turning this data into a report
'          instead of just using a grid - the data is easier to read and
'          export to other formats, graphics can be added easily.
'          Seagate Crystal Reports provides the flexiblitiy and power
'          to meet your data reporting needs.
'

Option Explicit

Dim SortDesc() As Boolean                   ' Records whether each column was last sorted
                                            ' Ascending or Descending
Dim m_RecordSource As String                ' The record source SQL string for Data1

' *************************************************************
' Initialize the Data grid to the same datasource as the form that
' is displaying this form
'
Private Sub Form_Load()
    On Error GoTo LoadErr                   ' Handle errors
    
    Me.Caption = g_FormCaption
    Data1.RecordSource = g_RecordSource
    Set Data1.Recordset = g_RecordSet
    m_RecordSource = Data1.RecordSource
    Data1.Refresh
        
    Set grdDataGrid.DataSource = Data1      ' Set our data source
    grdDataGrid.Refresh
    ReDim SortDesc(grdDataGrid.Columns.Count) ' Get the number of columns
                                              ' for later use in sorting
    Exit Sub
    
LoadErr:                                    ' Gracefully exit if any problems occur
    MsgBox "Error:" & Err & " " & Err.Description, vbExclamation
    Unload Me
End Sub

' *************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

' *************************************************************
' Make sure that the data grid covers all the window except for
' the top row of buttons
'
Private Sub Form_Resize()
    On Error Resume Next                ' Don't display errors in this procedure
    If Me.WindowState <> 1 Then _
        grdDataGrid.Height = Me.Height - (425 + picButtons.Height)
End Sub

' *************************************************************
' Confirm the user's decision to permenantly delete the row
'
Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    If MsgBox("Delete Current Row?", vbYesNo + vbQuestion) <> vbYes Then _
        Cancel = True
End Sub

' *************************************************************
' Confirm the user's decision before committing the changes
'
Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)
    If MsgBox("Commit changes?", vbYesNo + vbQuestion) <> vbYes Then _
        Cancel = True
End Sub

' *************************************************************
' Refresh the data set
'
Private Sub cmdRefresh_Click()
    Data1.Refresh
End Sub

' *************************************************************
' Allow the user the ability to sort the data by clicking on the
' column headers.  The last sorting action (whether Ascending or
' Descending) is remembered for each column.  Therefore, the columns
' will alternate ascending and descending sort order on an individual
' base irrespective of what the other columns are sorted like.
'
Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SortDir As String
    
    If SortDesc(ColIndex) Then
        SortDir = " DESC "
    Else
        SortDir = " ASC "
    End If
    
    ' Create the sort string
    Data1.RecordSource = m_RecordSource & " Order By [" & _
        grdDataGrid.Columns(ColIndex).DataField & "]" & SortDir
    Data1.Refresh
    SortDesc(ColIndex) = Not SortDesc(ColIndex)
End Sub

