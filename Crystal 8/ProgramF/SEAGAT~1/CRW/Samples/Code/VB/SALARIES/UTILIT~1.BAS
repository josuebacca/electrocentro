Attribute VB_Name = "Utilities"
' *************************************************************
' Purpose: Contains the global variables that are shared between
'          forms.  Also contains common procedures shared by
'          the forms in the applicaiton

Option Explicit

' Global Variables shared between frmMain and frmSelections
Public g_LeagueAbbrev As String         ' Used to save names of the leagues to be included in the report
Public g_SQLString As String            ' The SQL string applied to the data source for the report
Public g_GroupByTeam As Boolean         ' Set to the value of lblGroupByTeam in frmMain
Public g_SortOrder As String            ' "ASC" or "DESC"
Public g_TopCount As Long               ' The number of players to display - 0 means everyone
Public g_TypeofCount As Integer         ' Set to either vbTopN or vbCompareSalary

Public Const vbTopN = 1                 ' Applied to g_TypeofCount when the user selects Highest or Lowest Paid n players
Public Const vbCompareSalary = 2        ' Applied to g_TypeofCount when the user selects Higher or Lower than $n

Public g_TeamName As Boolean            ' Show the team name  True/False
Public g_PlayerName As Boolean          ' Show the player name  True/False
Public g_Salary As Boolean              ' Show the player's salary  True/False
Public g_YearName As Boolean            ' Show the year the player was paid the displayed salary  True/False

Public m_StartForm As Form              ' A reference holder for frmMain (the startup form)

'*********************************************************************
' A short procedure to give forms a nice teal shading.
'
Public Sub Dither(frm As Form)
    Dim intLoop As Integer                       ' Counter
   
    ' Set the pen parameters
    frm.DrawStyle = vbInsideSolid
    frm.DrawMode = vbCopyPen
    frm.ScaleMode = vbPixels
    frm.DrawWidth = 8
    frm.ScaleWidth = 176
    ' Shade the form
    For intLoop = 0 To 175
        frm.Line (intLoop, 0)-(intLoop - 1, Screen.Height), RGB(0, intLoop, intLoop), B
    Next intLoop
End Sub

