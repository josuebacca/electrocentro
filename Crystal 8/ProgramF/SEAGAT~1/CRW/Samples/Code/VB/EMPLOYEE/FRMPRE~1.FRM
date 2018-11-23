VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmPreview 
   Caption         =   "Preview Employee Profiles"
   ClientHeight    =   7785
   ClientLeft      =   1650
   ClientTop       =   1830
   ClientWidth     =   10395
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10395
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   8925
      TabIndex        =   1
      Top             =   7320
      Width           =   1425
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10365
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
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Display the Employee Profile report
'

Option Explicit

Dim m_Report As New dsrEmployeeProfiles
Dim adoRs As ADODB.Recordset

' *************************************************************
' Initialize the data set for the report and add the data set at
' run-time

Private Sub Form_Load()
    Dim SQL As String
    
    SQL = "SELECT" & _
        " Employee.`Last Name`, Employee.`First Name`, Employee.`Position`, Employee.`Birth Date`, Employee.`Home Phone`, Employee.`Extension`, Employee.`Photo`, Employee.`Notes`, " & _
        " `Employee Addresses`.`Address1`, `Employee Addresses`.`Address2`, `Employee Addresses`.`City`, `Employee Addresses`.`Region`, `Employee Addresses`.`Country`, `Employee Addresses`.`Postal Code` " & _
    " From" & _
        " `Employee` INNER JOIN `Employee Addresses` ON " & _
        " Employee.`Employee ID` = `Employee Addresses`.`Employee ID` "

    ' Create and bind the ADO Recordset object
    Set adoRs = CreateObject("adodb.recordset")

    ' Code for creating the recordset.  Uses the ODBC datasource name
    ' Xtreme Sample Data.  Make sure that this DSN is properly set up.
    adoRs.Open SQL, "database=Xtreme.mdb;uid=;pwd=;dsn=Xtreme Sample Database", adOpenKeyset, adLockBatchOptimistic

    m_Report.Database.SetDataSource adoRs
    CRViewer1.ReportSource = m_Report
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


