VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmPreview 
   Caption         =   "Load Picture Demo"
   ClientHeight    =   8505
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   10860
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8505
   ScaleWidth      =   10860
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
      Left            =   9360
      TabIndex        =   1
      Top             =   8025
      Width           =   1485
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10845
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
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Display an invoice in the Report Viewer. This demo
'          shows one way that pictures can be added to a
'          report (see the dsrLoadPicture code)
'

Option Explicit

Dim adoRs As ADODB.Recordset
Dim Report As New dsrLoadPicture


' *************************************************************
' Generate the data source for the report, and use a parameter
' to filter the data source
'
Private Sub Form_Load()
    Dim MyToday As String           ' This is a user defined date
    Dim SQL As String
    'All of the Record Selection Formula is integratated within the SQL
    '-----------------------------------------------
    MyToday = "#4/20/97#"
    SQL = "SELECT" & _
        " Customer.`Customer Name`, Customer.`Address1`, Customer.`Address2`, Customer.`City`, Customer.`Region`, Customer.`Country`, Customer.`Postal Code`," & _
        " Orders.`Order ID`, Orders.`Order Amount`, Orders.`Order Date`, Orders.`Ship Date`, Orders.`Shipped`, Orders.`Payment Received`,  Customer.`Last Year's Sales`" & _
    " From" & _
        " Customer INNER JOIN Orders ON" & _
            " Customer.`Customer ID` = Orders.`Customer ID`" & _
    " Where" & _
        " Orders.`Ship Date` <= " & MyToday & " AND" & _
        " Orders.`Order Amount` > 0. AND" & _
        " Orders.`Shipped` = 1 AND" & _
        " Orders.`Payment Received` = 0"
    
    'Add this to the SQL Statement to view only statements that are overdue:  and {@due date}<{@My Today}
    Set adoRs = CreateObject("adodb.recordset")
    
    ' Code for creating the recordset.  Uses the ODBC datasource name
    ' Xtreme Sample Data.  Make sure that this DSN is properly set up.
    adoRs.Open SQL$, "database=Xtreme.mdb;uid=;pwd=;dsn=Xtreme Sample Database", adOpenKeyset, adLockBatchOptimistic
    Report.Database.SetDataSource adoRs
    
    '//////////////////////////////////////////////////////////////////////////////////////
    'Using a Crystal Parameter in Record Selection formula the data set can be reduced further.
    'The parameter is set up in the designer and can be directly populated at runtime or
    'it can be left to prompt the user for a value.
    'The following code passes an employee name to the parameter field.
    'The parameter field can be viewed in the field view of the designer
    'To view the Selection Formula right click on a blank area in the designer and
    'select Report | Select Expert... and click on the Show Formula button
    
    Report.ParameterFields(1).SetCurrentValue "Anne Dodsworth"
    ' Set the viewer's report source to the Report object
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

