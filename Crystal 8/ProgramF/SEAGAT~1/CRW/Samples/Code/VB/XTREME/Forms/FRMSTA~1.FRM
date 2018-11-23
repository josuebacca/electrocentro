VERSION 5.00
Begin VB.Form frmStartup 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xtreme Mountain Bikes"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   510
      Left            =   4110
      TabIndex        =   5
      Top             =   5730
      Width           =   1815
   End
   Begin VB.CommandButton cmdLaunchReportView 
      Caption         =   "Show &Report"
      Height          =   510
      Left            =   4095
      TabIndex        =   3
      Top             =   4995
      Width           =   1815
   End
   Begin VB.CommandButton cmdLaunchDataView 
      Caption         =   "Show &Data View"
      Height          =   510
      Left            =   1005
      TabIndex        =   2
      Top             =   4995
      Width           =   1815
   End
   Begin VB.ListBox lstReportView 
      Height          =   2220
      Left            =   3270
      TabIndex        =   1
      Top             =   2685
      Width           =   2625
   End
   Begin VB.ListBox lstDataView 
      Height          =   3180
      Left            =   195
      TabIndex        =   0
      Top             =   1725
      Width           =   2625
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A demonstration of Seagate Crystal Reports and Visual Basic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   285
      TabIndex        =   6
      Top             =   810
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   90
      X2              =   5895
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   5880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Xtreme Mountain Bikes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   15
      TabIndex        =   4
      Top             =   120
      Width           =   5955
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   3030
      Picture         =   "frmStartup.frx":030A
      Top             =   630
      Width           =   2700
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  The first screen the user sees, this form provides
'           a navigation tool to move among all the various forms
'           and reports

Option Explicit

' These constants all correspond to a form that can be displayed.
Private Const vbEmpty As String = ""
Private Const vbCustomers As String = "Customers"
Private Const vbOrders As String = "Orders"
Private Const vbOrderDetails As String = "Order Details"
Private Const vbProducts As String = "Products"
Private Const vbProductTypes As String = "Product Types"
Private Const vbEmployees As String = "Employees"
Private Const vbEmployeeAddresses As String = "Employee Addresses"
Private Const vbStorePurchases As String = "Store Purchases"
Private Const vbFinancialData As String = "Financial Data"

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

' *************************************************************
' Initialize the form
'
Private Sub Form_Load()
    Set g_StartForm = Me
    InitListboxes
    Dither Me
End Sub

' *************************************************************
' Add items to the listbox corresponding to the available data
' forms that can be used to modify and report from data in Xtreme.mdb
'
Private Sub InitListboxes()
    ' Data Entry Forms
    With lstDataView
        .AddItem vbCustomers
        .AddItem vbEmpty
        .AddItem vbOrders
        .AddItem vbOrderDetails
        .AddItem vbEmpty
        .AddItem vbProducts
        .AddItem vbProductTypes
        .AddItem vbEmpty
        .AddItem vbEmployees
        .AddItem vbEmployeeAddresses
        .AddItem vbEmpty
        .AddItem vbStorePurchases
        .AddItem vbFinancialData
        .ListIndex = 0
    End With
    
    ' Reporting forms
    With lstReportView
        .AddItem vbProductCatalog
        .AddItem vbEmpty
        .AddItem vbMailLabels
        .AddItem vbEmpty
        .AddItem vbEmployeeProfiles
        .AddItem vbEmpty
        .AddItem vbWorldSales
        .AddItem vbIncomeStatement
        .AddItem vbBalanceStatement
        .ListIndex = 0
    End With
End Sub

' *************************************************************
' Only allow the user to select non-empty listitems
'
Private Sub lstDataView_Click()
    With lstDataView
        If .ListIndex = -1 Then Exit Sub
        If .List(.ListIndex) = vbEmpty Then .Selected(.ListIndex - 1) = True
    End With
    cmdLaunchDataView.Default = True
End Sub

' *************************************************************
' Open the selected data form
'
Private Sub lstDataView_DblClick()
    cmdLaunchDataView_Click
End Sub

' *************************************************************
' Only allow the user to select non-empty listitems
'
Private Sub lstReportView_Click()
    With lstReportView
        If .ListIndex = -1 Then Exit Sub
        If .List(.ListIndex) = vbEmpty Then .Selected(.ListIndex - 1) = True
    End With
    cmdLaunchReportView.Default = True
End Sub

' *************************************************************
' Open the selected report form
'
Private Sub lstReportView_DblClick()
    cmdLaunchReportView_Click
End Sub

' *************************************************************
' Open the selected data form
'
Private Sub cmdLaunchDataView_Click()
    If lstDataView.ListIndex = -1 Then Exit Sub
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    Select Case lstDataView.List(lstDataView.ListIndex)
    Case vbCustomers
        frmCustomer.Show vbModeless
    Case vbOrders
        frmOrders.Show vbModeless
    Case vbOrderDetails
        frmOrdersDetail.Show vbModeless
    Case vbProducts
        frmProduct.Show vbModeless
    Case vbProductTypes
        frmProductType.Show vbModeless
    Case vbEmployees
        frmEmployee.Show vbModeless
    Case vbEmployeeAddresses
        frmEmployeeAddresses.Show vbModeless
    Case vbStorePurchases
        frmPurchases.Show vbModeless
    Case vbFinancialData
        frmFinancials.Show vbModeless
    Case Else
        Me.Enabled = True
        Screen.MousePointer = vbDefault
    End Select
End Sub

' *************************************************************
' Open the selected report form
'
Private Sub cmdLaunchReportView_Click()
    If lstReportView.ListIndex = -1 Then Exit Sub
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    g_ReportName = lstReportView.List(lstReportView.ListIndex)
    frmDisplay.Show vbModal
End Sub

