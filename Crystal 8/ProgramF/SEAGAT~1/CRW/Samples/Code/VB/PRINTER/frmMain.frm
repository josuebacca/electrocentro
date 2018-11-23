VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   Caption         =   "Printer Options Demo"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1545
      Left            =   -15
      TabIndex        =   1
      Top             =   6675
      Width           =   9555
      Begin VB.CommandButton cmdPrinterSettings 
         Caption         =   "&Printer Setup Method"
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
         Left            =   2625
         TabIndex        =   13
         Top             =   1020
         Width           =   2310
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
         Height          =   465
         Left            =   7800
         TabIndex        =   12
         Top             =   1020
         Width           =   1695
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
         Height          =   465
         Left            =   5415
         TabIndex        =   11
         Top             =   1020
         Width           =   1890
      End
      Begin VB.ComboBox cboPaperSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7305
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   585
         Width           =   2190
      End
      Begin VB.ComboBox cboPaperOrientation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   585
         Width           =   2190
      End
      Begin VB.ComboBox cboPrinterDuplex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2535
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   585
         Width           =   2190
      End
      Begin VB.ComboBox cboPaperSource 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   585
         Width           =   2190
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Size:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7320
         TabIndex        =   10
         Top             =   330
         Width           =   1020
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Orientation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4935
         TabIndex        =   9
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duplexing:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2535
         TabIndex        =   8
         Top             =   330
         Width           =   915
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Source:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label lblPrinterName 
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   0
         Width           =   5505
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose:  Demonstrate how to programmatically change the report
'           printer settings at runtime.
'
'           There are four new properties of the report that allow
'           you to easily get and set PaperSize, PaperOrientation,
'           PaperSource, and PrinterDuplex for the report printer options.
'
'           There is also a new method called PrinterSetup that provides
'           a Windows standard printer setup window to allow the user
'           to change the printer properties directly at runtime.
'
'           The two sets of methods are independent of each other.  Changing
'           the Windows standard printer setup will not alter the report
'           printer settings and vice-versa.  These new properties and
'           method allow for much more runtime control over how the report
'           is printed.
'

Option Explicit

Dim m_Report As New CRPrinterSettings

' *************************************************************
' Changes the size of the paper for the report.
' Note that your printer may not accept all the available paper sizes.
'
Private Sub cboPaperSize_Click()
    CRPrinterSettings.PaperSize = cboPaperSize.ItemData(cboPaperSize.ListIndex)    ' Set the papersize option
    If Me.Visible Then CRViewer1.Refresh
End Sub

' *************************************************************
' Changes the paper orientation for the displayed report.
'
Private Sub cboPaperOrientation_Click()
    CRPrinterSettings.PaperOrientation = cboPaperOrientation.ItemData(cboPaperOrientation.ListIndex)
    If Me.Visible Then CRViewer1.Refresh
End Sub

' *************************************************************
' Changes the paper bin source for the displayed report.
' To enumerate the printer bins available for your printer, see EnumPrinterBins
' in Utilities.bas
' Note that your printer may override this setting to accommodate the papersize setting.
'
Private Sub cboPaperSource_Click()
    CRPrinterSettings.PaperSource = cboPaperSource.ItemData(cboPaperSource.ListIndex)
End Sub

' *************************************************************
' Changes the printer duplex setting for the displayed report.
'
Private Sub cboPrinterDuplex_Click()
    CRPrinterSettings.PrinterDuplex = cboPrinterDuplex.ItemData(cboPrinterDuplex.ListIndex)
End Sub

' *************************************************************
' Call the Printer Setup dialog.  This dialog does not reflect
' changes that you may have made via the PaperSource, PrinterDuplex
' and PaperSize methods, since this method changes the **Printer Settings**,
' not the **Report Printer Settings**.  The two sets of methods are
' independent and are intended for use in different situations.
'
Private Sub cmdPrinterSettings_Click()
    CRPrinterSettings.PrinterSetup Me.hWnd
End Sub

' *************************************************************
' Load the report into the viewer and initialize the printer settings
'
Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    ShowPrinterSource
    ShowPrinterDuplex
    ShowPaperOrientation
    ShowPaperSize
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    fraOptions.Top = ScaleHeight - fraOptions.Height
    fraOptions.Left = ScaleWidth - fraOptions.Width
    CRViewer1.Height = ScaleHeight - fraOptions.Height - 40
    CRViewer1.Width = ScaleWidth
End Sub

' *************************************************************
' Display the list of available printer bins in the cboPaperSource
' combo box.
'
Private Sub ShowPrinterSource()
    Dim i As Integer                            ' Counter
    Dim PaperSource As Integer
    
    lblPrinterName.Caption = "Report Printer Settings for: " & CRPrinterSettings.PrinterName       ' Display the printer name
    EnumPrinterBins CRPrinterSettings.PrinterName, cboPaperSource
    With cboPaperSource
        PaperSource = CRPrinterSettings.PaperSource    ' Get the report's paper source
        ' Cycle through the combo box and select the correct currently selected type of papersource in the report
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperSource Then .ListIndex = i
        Next i
    End With
End Sub

' *************************************************************
' Display the list of available printer duplexing types in the
' cboPrinterDuplex combo box.
'
Private Sub ShowPrinterDuplex()
    Dim i As Integer                            ' Counter
    Dim PrinterDuplex As CRPrinterDuplexType
    
    ' Addcbo is a helper function to make the code cleaner
    ' Addcbo format:   <combo name>, <item caption>, <.itemdata(.listindex) to assign>
    Addcbo cboPrinterDuplex, "Simplex", crPRDPSimplex
    Addcbo cboPrinterDuplex, "Horizontal", crPRDPHorizontal
    Addcbo cboPrinterDuplex, "Vertical", crPRDPVertical
    PrinterDuplex = CRPrinterSettings.PrinterDuplex    ' Get the report's printer duplex setting
    ' Cycle through the combo box and select the correct currently selected type of printer duplexing in the report
    With cboPrinterDuplex
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PrinterDuplex Then .ListIndex = i
        Next i
    End With

End Sub

' *************************************************************
' Display the list of available paper orientations in the
' cboPaperOrientation combo box.
'
Private Sub ShowPaperOrientation()
    Dim i As Integer                            ' Counter
    Dim PaperOrientation As CRPaperOrientation
    
    Addcbo cboPaperOrientation, "Portrait", crPortrait
    Addcbo cboPaperOrientation, "Landscape", crLandscape
    PaperOrientation = CRPrinterSettings.PaperOrientation  ' Get the report's paper orientation setting
    ' Cycle through the combo box and select the correct currently selected type of paper orientation in the report
    With cboPaperOrientation
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperOrientation Then .ListIndex = i
        Next i
    End With
End Sub

' *************************************************************
' Enumerate all the available paper sizes for the report
' *** Note that GetPaperSource, GetPaperSize, GetPrinterDuplex, GetPaperOrientation will not retrieve
' accurate settings unless the printer settings have been saved in the report or the properties have
' been set some place in code.
'
Private Sub ShowPaperSize()
    Dim i As Integer                            ' Counter
    Dim PaperSize As CRPaperSize
       
    ' Add the large number of supported paper sizes to the cboPaperSize combobox
    Addcbo cboPaperSize, "Default", crDefaultPaperSize
    Addcbo cboPaperSize, "Letter", crPaperLetter
    Addcbo cboPaperSize, "Small Letter", crPaperLetterSmall
    Addcbo cboPaperSize, "Legal", crPaperLegal
    Addcbo cboPaperSize, "10x14", crPaper10x14
    Addcbo cboPaperSize, "11x17", crPaper11x17
    Addcbo cboPaperSize, "A3", crPaperA3
    Addcbo cboPaperSize, "A4", crPaperA4
    Addcbo cboPaperSize, "A4 Small", crPaperA4Small
    Addcbo cboPaperSize, "A5", crPaperA5
    Addcbo cboPaperSize, "B4", crPaperB4
    Addcbo cboPaperSize, "B5", crPaperB5
    Addcbo cboPaperSize, "C Sheet", crPaperCsheet
    Addcbo cboPaperSize, "D Sheet", crPaperDsheet
    Addcbo cboPaperSize, "Envelope 9", crPaperEnvelope9
    Addcbo cboPaperSize, "Envelope 10", crPaperEnvelope10
    Addcbo cboPaperSize, "Envelope 11", crPaperEnvelope11
    Addcbo cboPaperSize, "Envelope 12", crPaperEnvelope12
    Addcbo cboPaperSize, "Envelope 14", crPaperEnvelope14
    Addcbo cboPaperSize, "Envelope B4", crPaperEnvelopeB4
    Addcbo cboPaperSize, "Envelope B5", crPaperEnvelopeB5
    Addcbo cboPaperSize, "Envelope B6", crPaperEnvelopeB6
    Addcbo cboPaperSize, "Envelope C3", crPaperEnvelopeC3
    Addcbo cboPaperSize, "Envelope C4", crPaperEnvelopeC4
    Addcbo cboPaperSize, "Envelope C5", crPaperEnvelopeC5
    Addcbo cboPaperSize, "Envelope C6", crPaperEnvelopeC6
    Addcbo cboPaperSize, "Envelope C65", crPaperEnvelopeC65
    Addcbo cboPaperSize, "Envelope DL", crPaperEnvelopeDL
    Addcbo cboPaperSize, "Envelope Italy", crPaperEnvelopeItaly
    Addcbo cboPaperSize, "Envelope Monarch", crPaperEnvelopeMonarch
    Addcbo cboPaperSize, "Envelope Personal", crPaperEnvelopePersonal
    Addcbo cboPaperSize, "E Sheet", crPaperEsheet
    Addcbo cboPaperSize, "Executive", crPaperExecutive
    Addcbo cboPaperSize, "Fanfold Legal German", crPaperFanfoldLegalGerman
    Addcbo cboPaperSize, "Fanfold Standard German", crPaperFanfoldStdGerman
    Addcbo cboPaperSize, "Fanfold US", crPaperFanfoldUS
    Addcbo cboPaperSize, "Folio", crPaperFolio
    Addcbo cboPaperSize, "Ledger", crPaperLedger
    Addcbo cboPaperSize, "Note", crPaperNote
    Addcbo cboPaperSize, "Quarto", crPaperQuarto
    Addcbo cboPaperSize, "Statement", crPaperStatement
    Addcbo cboPaperSize, "Tabloid", crPaperTabloid
    PaperSize = CRPrinterSettings.PaperSize    ' Get the report's paper size setting
    ' Cycle through the combo box and select the correct currently selected type of paper size in the report
    With cboPaperSize
        For i = 0 To .ListCount - 1
            If .ItemData(i) = PaperSize Then .ListIndex = i
        Next i
    End With
End Sub

' *************************************************************
' A small helper function for the ShowPrinterOption functions that
' helps reduce the amount of code to write
'   Addcbo format:   <combo name to add item to>, <item caption>, <.itemdata(.listindex) to assign>
Private Sub Addcbo(cbo As ComboBox, Name As String, index As Integer)
    cbo.AddItem Name                        ' Add the name of the item to the combo box
    cbo.ItemData(cbo.NewIndex) = index      ' Set the .itemdata(.listindex) for later retrieval
End Sub


' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

