VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewer"
   ClientHeight    =   7650
   ClientLeft      =   2190
   ClientTop       =   2715
   ClientWidth     =   10350
   Icon            =   "Viewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10350
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2250
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      Begin VB.ComboBox cboZoom 
         Height          =   315
         Left            =   4650
         TabIndex        =   4
         Top             =   41
         Width           =   840
      End
      Begin VB.TextBox txtCurPage 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Text            =   "0"
         Top             =   41
         Width           =   375
      End
      Begin VB.ComboBox cboSearch 
         Height          =   315
         Left            =   5630
         TabIndex        =   2
         Top             =   41
         Width           =   1150
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6855
      Left            =   45
      TabIndex        =   5
      Top             =   435
      Width           =   10260
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
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
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
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7365
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLst 
      Left            =   7860
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuCustomToolbar 
         Caption         =   "Display Custom Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuseperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCRToolbar 
         Caption         =   "Display Crystal Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDisplayGroupTree 
         Caption         =   "Display Group Tree"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Demonstrate how the Viewer can be controlled using external
'          controls
'

Option Explicit

' Module Constants
Const TOP_VIEW As Integer = 1 ' const used to determine Parent view

' These constants are used for setting the buttons in the toolbar
Const CLOSE_BUT As Integer = 1
Const FIRSTPAGE_BUT As Integer = 2
Const PREVPAGE_BUT As Integer = 3
Const NEXTPAGE_BUT As Integer = 5
Const LASTPAGE_BUT As Integer = 6
Const PRINT_BUT As Integer = 8
Const REFRESH_BUT As Integer = 9
Const GROUPTREE_BUT As Integer = 10
Const SEARCH_BUT As Integer = 12
    
Dim m_Report As New dsrInventory

' *************************************************************
' Enables/Disables toolbar depending upon the parameter State
'
Private Sub mEnableCloseButton(State As Boolean)
    Dim CloseButton As Button
    
    Set CloseButton = Toolbar.Buttons(CLOSE_BUT)
    CloseButton.Enabled = State
    
    Set CloseButton = Nothing
End Sub

' *************************************************************
' Calls Smart Viewer Zoom method to set the zoom percentage according
' to the user's choice
'
Private Sub cboZoom_Click()
    Dim strZmPcnt As String
    
    strZmPcnt = mstrGetZoomPercentage
    CRViewer1.Zoom (CInt(strZmPcnt))
End Sub

' *************************************************************
' When the Group tree is toggled on the Smart Viewer toolbar,
' the custom toolbar needs to be notified of the change.
'
Private Sub CRViewer1_GroupTreeButtonClicked(ByVal Visible As Boolean)
    Call mSetGroupTree
End Sub

' *************************************************************
' Disables Close button on custom toolbar if New Active View
' is the top most view
'
Private Sub CRViewer1_ViewChanged(ByVal oldViewIndex As Long, ByVal newViewIndex As Long)
    If newViewIndex = (TOP_VIEW - 1) Then
        mEnableCloseButton False
    Else
        mEnableCloseButton True
    End If
End Sub

' *************************************************************
' Displays the zoom percentage in dropdown box
'
Private Sub mSetZoomLevelOnControl(strZoomLevel As String)
        Dim indx As Integer
        
        Select Case strZoomLevel
        Case "1"  'Page Width
            cboZoom.Text = "51%"
        Case "2"  'Whole Page
            cboZoom.Text = "52%"
        Case Else
            For indx = 0 To cboZoom.ListCount
                If strZoomLevel = Left(cboZoom.List(indx), Len(strZoomLevel)) Then
                    cboZoom.Text = cboZoom.List(indx)
                    Exit Sub
                End If
            Next indx
    End Select
End Sub

' *************************************************************
' The zoom level was changed by Smart Viewer toolbar, so notify
' the custom toolbar of the change
'
Private Sub CRViewer1_ZoomLevelChanged(ByVal ZoomLevel As Integer)
    mSetZoomLevelOnControl CStr(ZoomLevel)
End Sub

' *************************************************************
' Get the current page of the viewer when the viewer has finished
' generating the page
'
Private Function migetCurrentPage() As Integer
    While CRViewer1.IsBusy
        DoEvents
    Wend
    migetCurrentPage = CRViewer1.GetCurrentPageNumber
End Function

' *************************************************************
' Load the Inventory Report into the Smart Viewer
'
Private Sub Form_Load()
    On Error GoTo Form_Load_err

    Screen.MousePointer = vbHourglass
    Call mCreateToolbar
    
    ' Set the report source
    CRViewer1.ReportSource = m_Report
    CRViewer1.ViewReport
    
    StatusBar.SimpleText = m_Report.ReportTitle
    txtCurPage = migetCurrentPage

    Call Form_Resize
    Screen.MousePointer = vbDefault
    Exit Sub

' Let the user know about any errors that might have occurred.
Form_Load_err:
    MsgBox "Error: " + CStr(Err) + Chr(10) + Chr(13) + _
            Error(Err), , "Form Load"
End Sub

' *************************************************************
' Load zoom percentages into combo box
'
Private Sub mInsertZoomPercentages()
    cboZoom.Clear
    With cboZoom
        .AddItem "400%", 0
        .AddItem "300%", 1
        .AddItem "200%", 2
        .AddItem "150%", 3
        .AddItem "100%", 4
        .AddItem "75%", 5
        .AddItem "50%", 6
        .AddItem "25%", 7
        .AddItem "Page Width", 8
        .AddItem "Whole Page", 9
    End With
    cboZoom.Text = cboZoom.List(4)  ' Set list to 100%
End Sub

' *************************************************************
' Reposition the controls in the form to compensate for the new
' form size.
'
Private Sub Form_Resize()
    Dim iTop As Integer
    Dim iAdjustment As Integer
    
    If Toolbar.Visible Then
        iTop = Toolbar.Height
        iAdjustment = Toolbar.Height + StatusBar.Height
    Else
        iTop = 0
        iAdjustment = StatusBar.Height
    End If
    
    Debug.Assert Me.Height > iAdjustment
    
    CRViewer1.Top = iTop
    CRViewer1.Left = 0
    CRViewer1.Height = Me.Height - iAdjustment
    CRViewer1.Width = Me.Width
End Sub

' *************************************************************
' Show info about this demo
'
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
' Toggle the Smart Viewer toolbar
'
Private Sub mnuCRToolbar_Click()
    CRViewer1.DisplayToolbar = Not CRViewer1.DisplayToolbar
    mnuCRToolbar.Checked = Not mnuCRToolbar.Checked
End Sub

' *************************************************************
' Toggle the Custom toolbar
'
Private Sub mnuCustomToolbar_Click()
    Toolbar.Visible = Not Toolbar.Visible
    mnuCustomToolbar.Checked = Not mnuCustomToolbar.Checked
    Call Form_Resize
End Sub

' *************************************************************
' Toggle the group tree
'
Private Sub mnuDisplayGroupTree_Click()
    CRViewer1.DisplayGroupTree = Not CRViewer1.DisplayGroupTree
    mnuDisplayGroupTree.Checked = Not mnuDisplayGroupTree.Checked
End Sub

' *************************************************************
' Checks active view to see if it is the top index view.  If the
' view is the top view (1) then the close view button on toolbar
' is disabled.
'
Private Sub mCheckActiveView()
    If CRViewer1.ActiveViewIndex = TOP_VIEW Then mEnableCloseButton False
End Sub

' *************************************************************
' Create the toolbar with all the appropriate icons.
'
Private Sub mCreateToolbar()
    Dim setButton As Button
    Dim ImageList As ListImage
    Dim strIconPath As String
    
    On Error GoTo mCreateToolbar_err
    
    ' Set path to icon directory
    strIconPath = App.Path + "\icons\"
    
    ' Set icon sizing and create list
    ImgLst.ImageHeight = 16
    ImgLst.ImageWidth = 16
    With ImgLst.ListImages
        Set ImageList = .Add(, "close", LoadPicture(strIconPath + "w95mbx01.ico"))
        Set ImageList = .Add(, "print", LoadPicture(strIconPath + "printfld.ico"))
        Set ImageList = .Add(, "refresh", LoadPicture(strIconPath + "refresh.ico"))
        Set ImageList = .Add(, "search", LoadPicture(strIconPath + "binoculr.ico"))
        Set ImageList = .Add(, "firstpage", LoadPicture(strIconPath + "arw03lt.ico"))
        Set ImageList = .Add(, "prevpage", LoadPicture(strIconPath + "arw04lt.ico"))
        Set ImageList = .Add(, "nextpage", LoadPicture(strIconPath + "arw04rt.ico"))
        Set ImageList = .Add(, "lastpage", LoadPicture(strIconPath + "arw03rt.ico"))
        Set ImageList = .Add(, "grouptree", LoadPicture(strIconPath + "graph14.ico"))
    End With
    
    ' Bind toolbar to imagelist and set buttons on toolbar which require icons
    Set Toolbar.ImageList = ImgLst
    Toolbar.ButtonHeight = ImgLst.ImageHeight
    Toolbar.ButtonWidth = ImgLst.ImageWidth
    
    ' Set an icon for each button on the toolbar
    Set setButton = Toolbar.Buttons(CLOSE_BUT)
        setButton.Image = "close"
        setButton.ToolTipText = "Close Current View"
    
    Set setButton = Toolbar.Buttons(FIRSTPAGE_BUT)
        setButton.Image = "firstpage"
        setButton.ToolTipText = "Go to First Page"
    
    Set setButton = Toolbar.Buttons(PREVPAGE_BUT)
        setButton.Image = "prevpage"
        setButton.ToolTipText = "Go to Previous Page"
    
    Set setButton = Toolbar.Buttons(NEXTPAGE_BUT)
        setButton.Image = "nextpage"
        setButton.ToolTipText = "Go to Next Page"
    
    Set setButton = Toolbar.Buttons(LASTPAGE_BUT)
        setButton.Image = "lastpage"
        setButton.ToolTipText = "Go to Last Page"
    
    Set setButton = Toolbar.Buttons(PRINT_BUT)
        setButton.Image = "print"
        setButton.ToolTipText = "Print Report"
    
    Set setButton = Toolbar.Buttons(REFRESH_BUT)
        setButton.Image = "refresh"
        setButton.ToolTipText = "Refresh"
        
    Set setButton = Toolbar.Buttons(SEARCH_BUT)
        setButton.Image = "search"
        setButton.ToolTipText = "Search Text"
    
    Set setButton = Toolbar.Buttons(GROUPTREE_BUT)
        setButton.Image = "grouptree"
        setButton.ToolTipText = "Toggle Group Tree"
        If CRViewer1.DisplayGroupTree Then
            setButton.Value = tbrPressed
        Else
            setButton.Value = tbrUnpressed
        End If
        
        mInsertZoomPercentages 'insert zoom percentages into combo box
    Exit Sub

' Handle toolbar errors
mCreateToolbar_err:
    MsgBox "Error: " + CStr(Err) + Chr(10) + Chr(13) + _
            Error(Err), , "Creating Toolbar"
End Sub

' *************************************************************
' Determine which toolbar button was clicked and then execute
' that action
'
Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
    Case CLOSE_BUT
        ' Closes active view
        If CRViewer1.ActiveViewIndex > 1 Then
            CRViewer1.CloseView (CRViewer1.ActiveViewIndex)
        End If
        mCheckActiveView
    Case FIRSTPAGE_BUT
        CRViewer1.ShowFirstPage
        txtCurPage = migetCurrentPage
    Case PREVPAGE_BUT
        CRViewer1.ShowPreviousPage
        txtCurPage = migetCurrentPage
    Case NEXTPAGE_BUT
        CRViewer1.ShowNextPage
        txtCurPage = migetCurrentPage
    Case LASTPAGE_BUT
        CRViewer1.ShowLastPage
        txtCurPage = migetCurrentPage
    Case PRINT_BUT
        CRViewer1.PrintReport
    Case REFRESH_BUT
        CRViewer1.Refresh
    Case SEARCH_BUT
        Call mSearchForText
    Case GROUPTREE_BUT
        CRViewer1.DisplayGroupTree = Not CRViewer1.DisplayGroupTree
        Call mSetGroupTree
    End Select
End Sub

' *************************************************************
' Search for text in the report
'
Private Sub mSearchForText()
    If cboSearch.Text = "" Then
        MsgBox "Search Text not specified", vbOKOnly, "Search Text"
    Else
        CRViewer1.SearchForText (cboSearch.Text)
    End If
End Sub

' *************************************************************
' Set the GroupTree button to be the same as the "Display Group Tree"
' option in the viewer
'
Private Sub mSetGroupTree()
    Dim GroupTreeButton As Button
    
    Set GroupTreeButton = Toolbar.Buttons(GROUPTREE_BUT)
    If CRViewer1.DisplayGroupTree Then
        GroupTreeButton.Value = tbrPressed
    Else
        GroupTreeButton.Value = tbrUnpressed
    End If
    
    Set GroupTreeButton = Nothing
End Sub

' *************************************************************
' Return a string that the viewer can use to set the zoom-in value
'
Private Function mstrGetZoomPercentage() As String
    Dim ipercentpos As String
    
    ipercentpos = InStr(1, cboZoom.Text, "%")
    If ipercentpos <> 0 Then
        mstrGetZoomPercentage = Left(cboZoom.Text, ipercentpos - 1) ' Returns a numeric string
    ElseIf cboZoom.Text = "Page Width" Then
        mstrGetZoomPercentage = "1"
    ElseIf cboZoom.Text = "Whole Page" Then
        mstrGetZoomPercentage = "2"
    End If
End Function
