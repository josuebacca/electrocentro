VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelections 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pro Athlete Salary Report"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelections.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About the Sample"
      Height          =   510
      Left            =   90
      TabIndex        =   31
      Top             =   7515
      Width           =   2025
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "&Year"
      Height          =   195
      Left            =   5835
      TabIndex        =   23
      Top             =   7020
      Width           =   195
   End
   Begin VB.CheckBox chkSalary 
      Caption         =   "&Salary"
      Height          =   195
      Left            =   5835
      TabIndex        =   22
      Top             =   6570
      Width           =   195
   End
   Begin VB.CheckBox chkTeam 
      Caption         =   "&Team"
      Height          =   195
      Left            =   5835
      TabIndex        =   21
      Top             =   6135
      Width           =   195
   End
   Begin VB.CheckBox chkPlayer 
      Caption         =   "&Player"
      Height          =   195
      Left            =   5835
      TabIndex        =   20
      Top             =   5685
      Width           =   195
   End
   Begin VB.OptionButton optAll 
      BackColor       =   &H00000000&
      Height          =   240
      Left            =   5835
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   2775
      Width           =   240
   End
   Begin VB.OptionButton optHighest 
      BackColor       =   &H00000000&
      Caption         =   "&Highest Paid"
      Height          =   240
      Left            =   5835
      TabIndex        =   11
      Top             =   3225
      Width           =   240
   End
   Begin VB.OptionButton optLowest 
      BackColor       =   &H00000000&
      Caption         =   "&Lowest Paid"
      Height          =   240
      Left            =   5835
      TabIndex        =   10
      Top             =   3705
      Width           =   240
   End
   Begin VB.OptionButton optSalaryHigher 
      BackColor       =   &H00000000&
      Caption         =   "Salary > $"
      Height          =   240
      Left            =   5835
      TabIndex        =   9
      Top             =   4170
      Width           =   240
   End
   Begin VB.OptionButton optSalaryLower 
      BackColor       =   &H00000000&
      Caption         =   "Salary < $"
      Height          =   240
      Left            =   5835
      TabIndex        =   8
      Top             =   4635
      Width           =   240
   End
   Begin VB.TextBox txtSalaryHigher 
      Height          =   375
      Left            =   7185
      TabIndex        =   7
      Text            =   "txtPaidHigher"
      Top             =   4080
      Width           =   1515
   End
   Begin VB.TextBox txtHigh 
      Height          =   360
      Left            =   7470
      TabIndex        =   6
      Text            =   "txtHigh"
      Top             =   3120
      Width           =   1230
   End
   Begin VB.TextBox txtLow 
      Height          =   360
      Left            =   7470
      TabIndex        =   5
      Text            =   "txtLow"
      Top             =   3600
      Width           =   1230
   End
   Begin VB.TextBox txtSalaryLower 
      Height          =   375
      Left            =   7185
      TabIndex        =   4
      Text            =   "txtPaidLower"
      Top             =   4605
      Width           =   1515
   End
   Begin VB.CheckBox chkGroupByTeam 
      Height          =   195
      Left            =   5835
      TabIndex        =   3
      Top             =   2400
      Width           =   195
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show Report"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4695
      TabIndex        =   2
      Top             =   7560
      Width           =   2025
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Exit"
      Height          =   510
      Left            =   6810
      TabIndex        =   1
      Top             =   7560
      Width           =   2025
   End
   Begin MSComctlLib.TreeView tvwSelect 
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   1890
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   9710
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   6
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fields to Include"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   5730
      TabIndex        =   30
      Top             =   5250
      Width           =   2055
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   29
      Top             =   7005
      Width           =   1380
   End
   Begin VB.Label lblSalary 
      BackStyle       =   0  'Transparent
      Caption         =   "     Sala&ry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   28
      Top             =   6555
      Width           =   1110
   End
   Begin VB.Label lblTeam 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   27
      Top             =   6105
      Width           =   1380
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   26
      Top             =   5655
      Width           =   1380
   End
   Begin VB.Label lblSalaryLower 
      BackStyle       =   0  'Transparent
      Caption         =   "     Sa&lary < $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5865
      TabIndex        =   25
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Label lblSalaryHigher 
      BackStyle       =   0  'Transparent
      Caption         =   "     S&alary > $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5865
      TabIndex        =   24
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Image imgLeagueLogo 
      Height          =   2010
      Index           =   1
      Left            =   7785
      Picture         =   "frmSelections.frx":030A
      Top             =   5325
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   3300
      Index           =   1
      Left            =   5655
      Picture         =   "frmSelections.frx":2AE4
      Top             =   4110
      Width           =   3120
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Leagues, Teams or Players"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   19
      Top             =   1635
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Professional Athlete Salary Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   1485
      TabIndex        =   18
      Top             =   90
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player Report Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   5730
      TabIndex        =   17
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Label lblGroupByTeam 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Group By Year and Team"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5865
      TabIndex        =   16
      Top             =   2370
      Width           =   2925
   End
   Begin VB.Label lblAll 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Show All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   15
      Top             =   2775
      Width           =   1380
   End
   Begin VB.Label lblHighest 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Highest Paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   14
      Top             =   3210
      Width           =   1785
   End
   Begin VB.Label lblLowest 
      BackStyle       =   0  'Transparent
      Caption         =   "     &Lowest Paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5865
      TabIndex        =   13
      Top             =   3675
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3300
      Index           =   0
      Left            =   5655
      Picture         =   "frmSelections.frx":24366
      Top             =   1875
      Width           =   3120
   End
   Begin VB.Image imgLeagueLogo 
      Height          =   1335
      Index           =   3
      Left            =   90
      Picture         =   "frmSelections.frx":45BE8
      Top             =   165
      Width           =   1245
   End
   Begin VB.Image imgLeagueLogo 
      Height          =   1335
      Index           =   2
      Left            =   7515
      Picture         =   "frmSelections.frx":47D5E
      Top             =   165
      Width           =   1290
   End
   Begin VB.Image imgLeagueLogo 
      Height          =   735
      Index           =   0
      Left            =   3510
      Picture         =   "frmSelections.frx":4A038
      Top             =   750
      Width           =   1485
   End
End
Attribute VB_Name = "frmSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************
' Purpose: Provide an interface that allows the user to compare
'          the salaries of any players from any teams from any
'          leagues.
'

Option Explicit

Const numLeagues = 7                    ' The number of leagues that are available in the database

Dim m_Filter(numLeagues) As clsFilter   ' Used to determine which leagues, teams and players need
                                        ' to be excluded and included in the report

' *************************************************************
' Save the user's settings into a format that can be easily read
' later when generating the report in frmMain
'
Private Sub SaveSettings()
    Dim firstLevelNode As Node
    Dim secondLevelNode As Node
    Dim teamNode As Node
    Dim LeagueCount As Integer
    Dim i As Integer
        
    ' Clear all the clsFilters (the class used to store the user's choices about which
    ' leagues, teams and players to include and exclude
    For i = 0 To numLeagues
        Set m_Filter(i) = New clsFilter
    Next i
    g_LeagueAbbrev = ""
    
    ' Now save the user's settings to dynamically generate the report
    g_GroupByTeam = chkGroupByTeam
    g_TeamName = chkTeam.Value
    g_PlayerName = chkPlayer.Value
    g_Salary = chkSalary.Value
    g_YearName = chkYear.Value
    
    ' Root Level
    LeagueCount = 0
    Set firstLevelNode = tvwSelect.Nodes(1)
    While Not (firstLevelNode Is Nothing)
        ' Table Name
        Set secondLevelNode = firstLevelNode.Child
        While Not (secondLevelNode Is Nothing)
            LeagueCount = LeagueCount + 1
            ' Check if we need to include this table
            If NeedTable(secondLevelNode) > 0 Then
                m_Filter(LeagueCount).League = secondLevelNode.Tag
                g_LeagueAbbrev = g_LeagueAbbrev + secondLevelNode.Tag
                ' Do we need to specially exclude any teams?
                Set teamNode = secondLevelNode.Child
                While Not (teamNode Is Nothing)
                    If NeedTeam(teamNode) = 0 Then
                        m_Filter(LeagueCount).AddTeam teamNode.Text
                    ElseIf teamNode.Child.Text <> "1" Then
                        ' Do we need to specially exclude any players?
                        ExcludePlayers m_Filter(LeagueCount), teamNode
                    End If
                    Set teamNode = teamNode.Next
                Wend
            End If
            Set secondLevelNode = secondLevelNode.Next
        Wend
        Set firstLevelNode = firstLevelNode.Next
    Wend
End Sub

' *************************************************************
' Check if any individual players need to be excluded. If an
' individual player is to be excluded, then add this player
' name to the applicable clsFilter.AddPlayer
'
Private Sub ExcludePlayers(filter As clsFilter, nodx As Node)
    Dim curNode As Node
    
    Set curNode = nodx.Child
    While Not (curNode Is Nothing)
        If curNode.Checked = False Then filter.AddPlayer curNode.Text
        Set curNode = curNode.Next
    Wend
End Sub

' *************************************************************
' Recursively check if the treeview node for the team or any node for the players
' on the team is checked.  If zero is returned, then the team needs to be
' excluded in the SQL statement.  A non-zero value means that this team
' needs to be included in the data set for the report
'
Private Function NeedTeam(nodx As Node) As Long
    Dim curNode As Node
    Dim IncludeTeam As Long
    
    IncludeTeam = 0
    If nodx.Checked Then
        IncludeTeam = IncludeTeam + 1
    Else
        Set curNode = nodx.Child
        While Not (curNode Is Nothing)
            IncludeTeam = IncludeTeam + NeedTeam(curNode)
            Set curNode = curNode.Next
        Wend
    End If
    NeedTeam = IncludeTeam
End Function

' *************************************************************
' Recursively check through all the child nodes of this league node
' If the return value is zero, then the table doesn't need to be included
' If the return value is non-zero, then the table needs to be included
' in the SQL statement for the report.
'
Private Function NeedTable(nodx As Node) As Long
    Dim curNode As Node
    Dim includeTable As Long
    
    includeTable = 0
    If nodx.Checked Then
        includeTable = includeTable + 1
    Else
        Set curNode = nodx.Child
        While Not (curNode Is Nothing)
            includeTable = includeTable + NeedTable(curNode)
            Set curNode = curNode.Next
        Wend
    End If
    NeedTable = includeTable
End Function

' *************************************************************
' Now that the user's preferences have been saved in SaveSettings(),
' generate the SQL statement that will be used in the report.
' Also save the other relevant settings, such as the number of leagues
' that are being included and whether the user wants ascending or
' descending sorting.  This information is used to rearrange the
' report headings and details in frmMain
'
Private Sub GenerateSQLStatement()
    Dim s As String
    Dim strTemp As String
    Dim i As Integer, j As Long
    
    ' Add parameter to only select the topN players from each league
    If optHighest.Value = True Then
        g_SortOrder = "DESC"
        g_TopCount = CLng(txtHigh.Text)
        g_TypeofCount = vbTopN
    ' Add parameter to only select the bottomN players from each league
    ElseIf optLowest.Value = True Then
        g_TopCount = CLng(txtLow.Text)
        g_SortOrder = "ASC"
        g_TypeofCount = vbTopN
    ' Add parameter to only select those players that make more than $<txtSalaryHigher>
    ElseIf optSalaryHigher.Value = True Then
        g_TopCount = CLng(txtSalaryHigher.Text)
        g_SortOrder = "DESC"
        g_TypeofCount = vbCompareSalary
    ' Add parameter to only select those players that make less than $<txtSalaryLower>
    ElseIf optSalaryLower.Value = True Then
        g_TopCount = CLng(txtSalaryLower.Text)
        g_SortOrder = "ASC"
        g_TypeofCount = vbCompareSalary
    ' Default to "Show all players" (optAll.Value = True)
    Else
        g_TopCount = 0
        g_SortOrder = "DESC"
        g_TypeofCount = vbTopN
    End If
    
    ' Begin generating the SQL statement that will be applied to the PlayerSalaries.mdb database
    strTemp = ""
    ' Select the tables that will be needed in the report
    For i = 0 To numLeagues
        If m_Filter(i).League <> "" Then
            If strTemp <> "" Then strTemp = strTemp & " UNION "
            strTemp = strTemp & "SELECT Team, Salary, Year, Player FROM " & m_Filter(i).League
            
            ' Exclude individual teams
            s = ""
            If m_Filter(i).GetTeams.Count > 0 Then
                For j = 1 To m_Filter(i).GetTeams.Count
                    If s <> "" Then s = s & " AND "
                    s = s & "Team <> '" & m_Filter(i).GetTeams(j) & "'"
                Next j
                strTemp = strTemp & " WHERE (" & s & ")"
            End If
            
            ' Exclude individual players
            s = ""
            If m_Filter(i).GetPlayers.Count > 0 Then
                For j = 1 To m_Filter(i).GetPlayers.Count
                    If s <> "" Then s = s & " AND "
                    s = s & "Player <> '" & m_Filter(i).GetPlayers(j) & "'"
                Next j
                If Right(strTemp, 1) = ")" Then
                    strTemp = strTemp & " AND " & s
                Else
                    strTemp = strTemp & " WHERE " & s
                End If
            End If
        End If
    Next i
    
    ' Order the dataset Ascending or Descending
    strTemp = strTemp & " Order by Salary " & g_SortOrder
    g_SQLString = strTemp
End Sub

' *************************************************************
' Ensure that there is only valid numeric data entered into any
' of the textboxes.
'
Private Function EnsureValidData() As Boolean
    Dim bCheck As Boolean
    
    On Error GoTo Finish
    ' Assume that the data is not OK
    bCheck = False
    ' For each textbox, only check the text if it is enabled, and only
    ' permit positive integers or long values
    If txtLow.Enabled Then
        If CInt(txtLow.Text) = txtLow.Text And _
           CInt(txtLow.Text) > 0 Then bCheck = True
    ElseIf txtHigh.Enabled Then
       If CInt(txtHigh.Text) = txtHigh.Text And _
          CInt(txtHigh.Text) > 0 Then bCheck = True
    ElseIf txtSalaryHigher.Enabled Then
       If CLng(txtSalaryHigher.Text) = txtSalaryHigher.Text And _
          CLng(txtSalaryHigher.Text) > 0 Then bCheck = True
    ElseIf txtSalaryLower.Enabled Then
       If CLng(txtSalaryLower.Text) = txtSalaryLower.Text And _
          CLng(txtSalaryLower.Text) > 0 Then bCheck = True
    Else
        bCheck = True
    End If

    EnsureValidData = bCheck
    If bCheck Then Exit Function
    
Finish:
    ' Let the user know there is a problem that needs to fixed
    EnsureValidData = False
    MsgBox "Please enter positive integer values into the text boxes.", vbExclamation, "Correction Required"
End Function

' *************************************************************
' Ensure that at least one node in the treeview has been selected
'
Private Function EnsureNodeSelected() As Boolean
    Dim bCheck As Boolean
    Dim nodx As Node
    
    ' Assume false
    bCheck = False
    For Each nodx In tvwSelect.Nodes
        If nodx.Checked = True Then bCheck = True
    Next
        
    ' Let the user know there is a problem that needs to fixed
    If bCheck = False Then _
        MsgBox "Please select a League, Year, Team or Player to use as a data source.", vbInformation, "Correction Required"
    
    EnsureNodeSelected = bCheck
End Function

' *************************************************************
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdDone_Click()
    Unload Me
End Sub

' *************************************************************
' Save the user's settings, generate the SQL statement and if
' all goes well, show the report
'
Private Sub cmdShow_Click()
    If Not EnsureValidData Then Exit Sub
    If Not EnsureNodeSelected Then Exit Sub
    Screen.MousePointer = vbHourglass
    SaveSettings
    GenerateSQLStatement
    Set m_StartForm = Me
    Me.Enabled = False
    frmMain.Show vbModeless
End Sub

' *************************************************************
' Initialize the form
'
Private Sub Form_Load()
    Dither Me
    InitVisualOptions
    InitTreeview
End Sub

' *************************************************************
' Set the initial default settings for the checkboxes, radio buttons
' and text boxes
'
Private Sub InitVisualOptions()
    chkGroupByTeam.Value = 1
    optAll.Value = True
    txtHigh.Enabled = False
    txtLow.Enabled = False
    txtLow.Text = "100"
    txtHigh.Text = "100"
    txtSalaryHigher.Text = "1,000,000"
    txtSalaryLower.Text = "1,000,000"
    chkPlayer.Value = 1
    chkTeam.Value = 1
    chkSalary.Value = 1
    chkYear.Value = 1
End Sub

' *************************************************************
' Initialize the treeview.  There is one node for each table
' in the PlayerSalaries.mdb database. AddNodes is a helper function
' that adds all the teams in a league as child nodes to the league
' node.  Note that the nodes for the players are only added to the
' treeview when they need to be viewed.  This on demand treeview node
' creation greatly decreases the amount of time it takes to populate
' the treeview.
'
Private Sub InitTreeview()
    Dim Connection As New ADODB.Connection              ' Use ADO for the connection
    Dim RS As New ADODB.Recordset
    Dim SortKey As String
    
    ' Open the connection
    Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\PlayerSalaries.mdb;Mode=Read"
    Connection.Open
    
    SortKey = " Order By Team, Player"
    ' Get all the records from our data tables
    RS.Open "Select * From NBA9697" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "NBA", "1996/1997", "NBA9697", True
    RS.Close
    
    RS.Open "Select * From NBA9798" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "NBA", "1997/1998", "NBA9798", False
    RS.Close
    
    RS.Open "Select * From NHL9697" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "NHL", "1996/1997", "NHL9697", True
    RS.Close
    
    RS.Open "Select * From NHL9798" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "NHL", "1997/1998", "NHL9798", False
    RS.Close
    
    RS.Open "Select * From NFL1997" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "NFL", "1997", "NFL1997", True
    RS.Close
    
    RS.Open "Select * From MLB1998" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "MLB", "1998", "MLB1998", True
    RS.Close
    
    RS.Open "Select * From MLB1999" + SortKey, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    AddNodes RS, "MLB", "1999", "MLB1999", False
    RS.Close
    
    ' Cleanup
    Connection.Close
    Set Connection = Nothing
    Set RS = Nothing
End Sub

' *************************************************************
' This helper function for InitTreeview takes a recordset and adds
' a group of nodes in the format:
'   Root Level = <Name of League>
'   1st Child = <Year>
'   2nd Child = <Teams>
'   3rd Child = <Players>
' The 3rd level nodes are not added until the user chooses to expand
' the team node to view and individually select players
'
Private Sub AddNodes(RS As ADODB.Recordset, League As String, _
                     Year As String, TableName As String, NewLeague As Boolean)
    Dim nodx As Node
    Dim s As String
    Dim UniquePrefix As String
    
    ' Add a new root node if this is a new league
    If NewLeague = True Then
        Set nodx = tvwSelect.Nodes.Add(, , League, League)
        nodx.Tag = "Root" & League
    End If
    
    ' Add a child node that represents the year for the salary information
    UniquePrefix = TableName
    Set nodx = tvwSelect.Nodes.Add(League, tvwChild, UniquePrefix, Year)
    nodx.Tag = TableName
    nodx.EnsureVisible
    s = ""
    While Not RS.EOF
        ' Team
        If RS!Team <> s Then
            Set nodx = tvwSelect.Nodes.Add(UniquePrefix, tvwChild, UniquePrefix + RS!Team, RS!Team)
            ' Add a dummy node to make the "Plus" Sign appear
            Set nodx = tvwSelect.Nodes.Add(UniquePrefix + RS!Team, tvwChild, UniquePrefix + RS!Team + "1", "1")
            s = RS!Team
        End If
        RS.MoveNext
    Wend
End Sub

' *************************************************************
' Make the transparent label seem like it is attached to the
' checkbox immediately to its left
'
Private Sub lblGroupByTeam_Click()
    chkGroupByTeam.Value = Abs(Not (-1 * chkGroupByTeam.Value))
End Sub

' *************************************************************
Private Sub lblPlayer_Click()
    chkPlayer.Value = Abs(Not (-1 * chkPlayer.Value))
End Sub

' *************************************************************
Private Sub lblTeam_Click()
    chkTeam.Value = Abs(Not (-1 * chkTeam.Value))
End Sub

' *************************************************************
Private Sub lblSalary_Click()
    chkSalary.Value = Abs(Not (-1 * chkSalary.Value))
End Sub

' *************************************************************
Private Sub lblYear_Click()
    chkYear.Value = Abs(Not (-1 * chkYear.Value))
End Sub

' *************************************************************
' Don't allow the treeview to collapse to the root nodes -- the
' second level nodes must always be visible
'
Private Sub tvwSelect_Collapse(ByVal Node As MSComctlLib.Node)
    If Left(Node.Tag, 4) = "Root" Then Node.Expanded = True
End Sub

' *************************************************************
' Expand the selected node on-demand.  Since nodes are only
' loaded as the viewer wants to view them, the program can load
' much quicker.  This user interface enhancement is done only
' to make the program appear quicker.
'
Private Sub tvwSelect_Expand(ByVal Node As MSComctlLib.Node)
    Dim Connection As New ADODB.Connection              ' Use ADO for the connection
    Dim RS As New ADODB.Recordset
    Dim SQL As String
    Dim NodeLevel As Integer
    Dim RootNode As Node
    Dim nodx As Node
    Dim Team As String
    
    ' Implement "On-demand" node expansion
    ' Open the connection
    Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\PlayerSalaries.mdb;Mode=Read"
    Connection.Open
    
    NodeLevel = 0
    ' Traverse to the Root node
    Set RootNode = Node
    While Not (RootNode.Parent Is Nothing)
        NodeLevel = NodeLevel + 1
        Set RootNode = RootNode.Parent
    Wend
    
    ' Check if this node has already been expanded
    If NodeLevel <> 2 Or Node.Children > 1 Then GoTo Cleanup
    
    ' Remove Dummy node if it exists
    If Node.Children = 1 Then
        tvwSelect.Nodes.Remove Node.Child.Index
    End If
    
    ' Load the players for this team from the database
    Team = Node.Text
    SQL = "Select * From " & Node.Parent.Tag & " Where Team = '" & Team & "'" & _
          " Order By Team, Player"
    RS.Open SQL, Connection, adOpenDynamic, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set nodx = tvwSelect.Nodes.Add(Node.Parent.Tag + Team, tvwChild, RootNode.Tag + Team + RS!Player, RS!Player)
        nodx.Checked = nodx.Parent.Checked
        RS.MoveNext
    Wend
    RS.Close

Cleanup:
    Set Connection = Nothing
    Set RS = Nothing
End Sub

' *************************************************************
' When a parent node is checked/unchecked, recursively check/uncheck
' all child nodes to be equivalent to the parent
'
Private Sub tvwSelect_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim curNode As Node
    Dim CheckedNow As Boolean
        
    CheckedNow = Node.Checked
    
    If Node.Children > 0 Then
        ' Recursively check/uncheck all the node's children
        Set curNode = Node.Child
        While Not (curNode Is Nothing)
            curNode.Checked = CheckedNow
            tvwSelect_NodeCheck curNode     ' Recursive Call
            Set curNode = curNode.Next
        Wend
    End If
End Sub

' *************************************************************
' Invert the check box when the user clicks on a node
'
Private Sub tvwSelect_NodeClick(ByVal Node As MSComctlLib.Node)
    Node.Checked = Not Node.Checked
End Sub

' *************************************************************
' User interface niceties
'
Private Sub optAll_Click()
    txtSalaryLower.Enabled = False
    txtSalaryHigher.Enabled = False
    txtHigh.Enabled = False
    txtLow.Enabled = False
End Sub

' *************************************************************
Private Sub optHighest_Click()
    txtSalaryLower.Enabled = False
    txtSalaryHigher.Enabled = False
    txtHigh.Enabled = True
    txtLow.Enabled = False
End Sub

' *************************************************************
Private Sub optLowest_Click()
    txtSalaryLower.Enabled = False
    txtSalaryHigher.Enabled = False
    txtHigh.Enabled = False
    txtLow.Enabled = True
End Sub

' *************************************************************
Private Sub optSalaryLower_Click()
    txtSalaryLower.Enabled = True
    txtSalaryHigher.Enabled = False
    txtHigh.Enabled = False
    txtLow.Enabled = False
End Sub

' *************************************************************
Private Sub optSalaryHigher_Click()
    txtSalaryLower.Enabled = False
    txtSalaryHigher.Enabled = True
    txtHigh.Enabled = False
    txtLow.Enabled = False
End Sub

' *************************************************************
' Make clicking the label equivalent to clicking on the radio button
'
Private Sub lblHighest_Click()
    optHighest.Value = Not optHighest.Value
End Sub

' *************************************************************
Private Sub lblLowest_Click()
    optLowest.Value = Not optLowest.Value
End Sub

' *************************************************************
Private Sub lblSalaryHigher_Click()
    optSalaryHigher.Value = Not optSalaryHigher.Value
End Sub

' *************************************************************
Private Sub lblSalaryLower_Click()
    optSalaryLower.Value = Not optSalaryLower.Value
End Sub

