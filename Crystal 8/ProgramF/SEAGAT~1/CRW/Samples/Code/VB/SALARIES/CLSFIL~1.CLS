VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
' A small class designed to allow individual leagues, teams and
' players to be individual excluded and excluded from the report
' This class essentially saves the user's preferences for multiple use
' and easy access in frmSelections
'

Option Explicit

Dim Players As New Collection           ' The players to be excluded from the report
Dim Teams As New Collection             ' The teams to be excluded from the report
Dim m_LeagueName As String              ' The league name that the players and teams are selected from

' *************************************************************
' Return the name of the league that this class instance is referring to
'
Property Get League() As String
    League = m_LeagueName
End Property

' *************************************************************
' Set the name of the league that this class instance is referring to
'
Property Let League(Name As String)
    m_LeagueName = Name
End Property

' *************************************************************
' Add a player to the "Exclude players" collection
'
Public Sub AddPlayer(Player As String)
    Players.Add Player
End Sub

' *************************************************************
' Return the Excluded players collection
'
Public Function GetPlayers() As Collection
    Set GetPlayers = Players
End Function

' *************************************************************
' Add a team to the excluded teams collection
'
Public Sub AddTeam(Team As String)
    Teams.Add Team
End Sub

' *************************************************************
' Return the excluded teams collection
'
Public Function GetTeams() As Collection
    Set GetTeams = Teams
End Function

' *************************************************************
' Free any memory we might have been using
'
Private Sub Class_Terminate()
    Set Players = Nothing
    Set Teams = Nothing
End Sub
