VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5C9EA131-127A-11D1-BFB4-00A0C936E6F9}#1.0#0"; "cselexpt.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search and Select Experts Sample"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select Expert"
      Default         =   -1  'True
      Height          =   405
      Left            =   1875
      TabIndex        =   5
      Top             =   6870
      Width           =   1650
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search Expert"
      Height          =   405
      Left            =   75
      TabIndex        =   3
      Top             =   6870
      Width           =   1650
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
      Height          =   6720
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9885
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
   Begin CSELEXPTLib.VSelExpert VSelExpert1 
      Height          =   390
      Left            =   5835
      TabIndex        =   4
      Top             =   6855
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   688
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate how to programmatically control the
'           Search and Select Wizards (New to Crystal Reports 8.0)
'
'           The Search and Select Wizards allow advanced filtering
'           and highlighting of data displayed in the ActiveX viewer.
'           In this sample, two different methods of using the
'           Search and Select Wizards are demonstrated.
'
'           1.) Two command buttons on the main form display a search
'               or select wizard with predefined fields and ranges
'               already set for the user.  The only displayed fields
'               and ranges are those that have been added at runtime
'
'           2.) The Search and Select Expert buttons on the top right
'               part of the viewer are over-ridden, allowing the
'               programmer to specify default criteria to search for.
'               Alternatively, the default search and select experts
'               can also be displayed
'

Option Explicit

Dim m_Report As New CrystalReport1

' *************************************************************
' Add two criteria to the Search or Select Expert:
'  1.) Product Type Name = "Gloves"
'  2.) Units in Stock >= 300
' The Function returns True if the user clicked OK and False
' if the user cancels the search or select expert dialog
'
Private Function Show_SearchAndSelectExpert() As Boolean
    Dim SelectField As New CSELEXPTLib.IDBFIELDDEF
    Dim SelectRange As New CSELEXPTLib.IDBFIELDRANGE
    
    ' Initialize the Search Expert control
    With VSelExpert1
        .ClearFields
        .ClearRanges
        .AddField
        .AddField
        .AddRange
        .AddRange
    End With
    
    ' Add the first criteria
    ' Each criteria requires a field and a matching range
    ' Set the properties for the first criteria's field
    Set SelectField = VSelExpert1.GetField(0)
    With SelectField
        .FieldType = crStringField
        .FieldName = "Product Type.Product Type Name"
        .AddValue "Product Type.Product Type Name"
    End With
    
    ' Set the properties for the first criteria's range
    Set SelectRange = VSelExpert1.GetRange(0)
    With SelectRange
        Set SelectField = .SetRangeFieldByName("Product Type.Product Type Name")
        .Operation = crSelectEqual
        .AddValue "Gloves"
    End With
    
    ' Add the second criteria
    ' Set the properties for the second criteria's field
    Set SelectField = VSelExpert1.GetField(1)
    With SelectField
        .FieldType = crInt32sField
        .FieldName = "Purchases.Units in Stock"
        .AddValue "Purchases.Units in Stock"
    End With
    
    Set SelectRange = VSelExpert1.GetRange(1)
    ' Set the properties for the second criteria's range
    With SelectRange
        Set SelectField = .SetRangeFieldByName("Purchases.Units in Stock")
        .Operation = crSelectGreaterThan
        .Inclusive = True
        .AddValue "300"          ' Every parameter used with AddValue must be passed
                                 ' in as a string
    End With
    
    ' Show the Selection Expert
    On Error GoTo HandleProblem
    Dim RetVal As Long
    RetVal = VSelExpert1.Show
    If RetVal = 0 Then
        Show_SearchAndSelectExpert = False
    Else
        Show_SearchAndSelectExpert = True
    End If
    Exit Function
    
HandleProblem:
    Show_SearchAndSelectExpert = False
End Function

' *************************************************************
' Apply a search filter to the data set displayed in the viewer
'
Private Sub cmdSearch_Click()
    m_Report.RecordSelectionFormula = ""
    ' GetSelectionString(True) would return a valid SQL string
    ' that could be applied to a data set
    If Show_SearchAndSelectExpert Then _
        CRViewer1.SearchByFormula VSelExpert1.GetSelectionString(False)
End Sub

' *************************************************************
' Select and highlight values in the viewer data set
'
Private Sub cmdSelect_Click()
    ' GetSelectionString(True) would return a valid SQL string
    ' that could be applied to a data set
    If Show_SearchAndSelectExpert Then
        m_Report.RecordSelectionFormula = VSelExpert1.GetSelectionString(False)
        CRViewer1.Refresh
    End If
End Sub

' *************************************************************
' To use your own selection criteria when the user clicks on the
' selection expert button on the viewer, set UseDefault = False
'
Private Sub CRViewer1_SelectionFormulaButtonClicked(selctionFormula As String, UseDefault As Boolean)
    UseDefault = False
    ' Add in our own predefined selection formula
    selctionFormula = "{Product.Product Name} IN " & _
                      " [ 'Active Oudoors Crochet Glove', 'InFlux Crochet Glove' ]"
End Sub

' *************************************************************
' UseDefault is set to true by default; seting UseDefault to
' False would allow you to override the default search expert
' and display your own search expert
'
Private Sub CRViewer1_SearchExpertButtonClicked(UseDefault As Boolean)
    UseDefault = True
End Sub

' *************************************************************
' Display the Selection Formula that will be applied to the
' data in the viewer.
'
Private Sub CRViewer1_SelectionFormulaBuilt(ByVal selctionFormula As String, UseDefault As Boolean)
    MsgBox "Your search formula is: " & selctionFormula, vbInformation, "Search Expert Formula"
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
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

