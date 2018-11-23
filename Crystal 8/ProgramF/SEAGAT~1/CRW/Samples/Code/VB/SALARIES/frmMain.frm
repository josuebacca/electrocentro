VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pro Athlete Salary Report"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Height          =   450
      Left            =   8040
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
   End
   Begin CRVIEWERLibCtl.CRViewer crViewer 
      Height          =   7770
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   9585
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
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
      EnableExportButton=   -1  'True
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
' Purpose: Display a report generated entirely at run-time using the
'          new (for version 8.0) Crystal Reports Creation API functions
'          This form takes the selections made in frmSelections and translates
'          them into a report that is placed in the Smart Viewer.  The report
'          uses the Creation API to manipulate the report into one of the hundreds
'          of possible layouts for the Pro Athlete Salary Report.
'

Option Explicit

Dim m_Proj As CRAXDRT.Application               ' Create a new COM instance of the Crystal Reports
                                                ' Run-time application
Dim WithEvents m_Report As CRAXDRT.Report       ' The dynamically generated report
Attribute m_Report.VB_VarHelpID = -1

Dim m_Connection As ADODB.Connection            ' Use ADO for the connection.  The connection should maintain state as
                                                ' long as frmMain doesn't go out of scope
Dim m_RS As ADODB.Recordset                     ' The Recordset used in AddDataSource

Dim AlreadyNoData As Integer                    ' Used to ensure the "No Data" message only appears once
Dim m_maxLogoHeight As Long                     ' The height of the tallest league logo
Dim m_Commmand As ADODB.Command

' *************************************************************
' Don't show the report header
'
Private Sub DoReportHeader()
    m_Report.Sections(1).Suppress = True
End Sub

' *************************************************************
' Display the appropriate logos in the page header.  A logo will
' only appear if the user has chosen players from that particular
' league.  Note how the header is formatted differently if one, two,
' three or four leagues are chosen.
'
Private Sub DoPageHeader()
    Dim Logo As OLEObject
    Dim PlayerSalariesTitle As TextObject
    Dim obj As Object           ' Used when moving the report objects around
    Dim obj2 As Object          ' Used when moving the report objects around
    Dim LocX As Long            ' The X-offset for the objects
    
    ' Section 2 is the page header
    ' For each section, only add the logo if the user has one or more players
    ' chosen from that league. Also find out the height of the tallest logo.
    With m_Report.Sections(2)
        LocX = 0
        m_maxLogoHeight = 0
        If InStr(g_LeagueAbbrev, "NBA") > 0 Then
            ' Add the object to the report
            Set Logo = .AddPictureObject(App.Path & "\res\NBA.bmp", LocX, 0)
            m_maxLogoHeight = Logo.Height
            LocX = LocX + 1400
        End If
        
        If InStr(g_LeagueAbbrev, "NHL") > 0 Then
            Set Logo = .AddPictureObject(App.Path & "\res\NHL.bmp", LocX, 0)
            If Logo.Height > m_maxLogoHeight Then m_maxLogoHeight = Logo.Height
            LocX = LocX + 1500
        End If
        
        If InStr(g_LeagueAbbrev, "MLB") > 0 Then
            Set Logo = .AddPictureObject(App.Path & "\res\MLB.bmp", LocX, 0)
            If Logo.Height > m_maxLogoHeight Then m_maxLogoHeight = Logo.Height
            LocX = LocX + 2000
        End If
        
        If InStr(g_LeagueAbbrev, "NFL") > 0 Then
            Set Logo = .AddPictureObject(App.Path & "\res\NFL.bmp", LocX, 0)
            If Logo.Height > m_maxLogoHeight Then m_maxLogoHeight = Logo.Height
        End If
        
        Set PlayerSalariesTitle = .AddTextObject("Pro Athlete Salary Report", 10, 0)
        PlayerSalariesTitle.Name = "Title"
    End With
    
    ' Now move the logos and text around
    With PlayerSalariesTitle
        .Width = 5000
        .Height = 600
        .Font.Size = 22
        .Font.Bold = True
        .Top = 300
        .Left = 1500
        
        ' Depending on the number of objects in the page header, reformat the page header
        ' title and logo
        Select Case m_Report.Sections(2).ReportObjects.Count
        Case 1:
            ' Nothing extra to do
        Case 2:
            ' Nothing extra to do
        Case 3:
            ' Reformat the Player Salaries Title to be located between the two logos
            .Font.Size = 20
            .Width = 3500
            .Height = 1000
            .Left = 1500
            .HorAlignment = crHorCenterAlign
            Set obj = m_Report.Sections(2).ReportObjects(2)
            obj.Left = 5250
        Case 4:
            ' Reformat the Player Salaries Title to be located beneath the two logos
            Set obj = m_Report.Sections(2).ReportObjects(1)
            Set obj2 = m_Report.Sections(2).ReportObjects(2)
            obj2.Left = obj.Left + obj.Width + 1500
            Set obj2 = m_Report.Sections(2).ReportObjects(3)
            obj2.Left = 5250
            .Top = 1400
        Case 5:
            .Top = 1400
        End Select
    End With

End Sub

' *************************************************************
' Add the group header.  The group headers are only added if the
' user has clicked on the "Group by Year and Team" check box in
' frmMain
'
Private Sub DoGroupHeader(SalaryFormula As FormulaFieldDefinition, _
                          TeamFormula As FormulaFieldDefinition, _
                          YearFormula As FormulaFieldDefinition)
    ' All the various objects that can be added to the header
    Dim PlayerTitle As TextObject
    Dim SalaryTitle As TextObject
    Dim TeamTitle As TextObject
    Dim PayrollTitle As TextObject
    Dim YearTitle As TextObject
    Dim TeamSummary As FieldObject
    Dim TeamGroupObj As FieldObject
    Dim YearHeading As FieldObject
    Dim SectionNum As Integer
    Dim vOffset As Long
    Dim groupNum As Integer     ' The number of groups can be 0, 1 or 2
    Dim LocX As Long            ' The X-offset for the objects
    Dim obj As Object
    Dim i As Integer
    
    groupNum = 0
    SectionNum = 2
    
    ' Add the group by Team
    If g_GroupByTeam Then
        ' Add a hidden group by Year to make the grouping look better.  This "invisible"
        ' grouping will not occur if the user is not displaying the League year.
        If g_YearName = True Then
            m_Report.AddGroup groupNum, YearFormula, crGCAnyValue, crAscendingOrder
            SectionNum = SectionNum + 1
            groupNum = groupNum + 1
            m_Report.Sections(SectionNum).Height = 1
        End If
        ' Add the "Group by Team" option
        If g_TeamName = True Then
            m_Report.AddGroup groupNum, TeamFormula, crGCAnyValue, crAscendingOrder
            SectionNum = SectionNum + 1
        End If
        ' Since we have groups to view, turn on the viewer tree
        crViewer.EnableGroupTree = True
        vOffset = 770
    Else
        ' Prevent the Field Headers from overlapping
        If m_maxLogoHeight = 2085 Then  ' NBA Logo Height
            vOffset = m_maxLogoHeight + 50
        Else
            vOffset = m_maxLogoHeight + 550
        End If
        crViewer.EnableGroupTree = False
    End If
    
    ' Since we don't know what fields are being added until runtime, we use an X offset variable
    ' to align the selected objects correctly.
    LocX = 10
    If g_PlayerName Then
        Set PlayerTitle = m_Report.Sections(SectionNum).AddTextObject("Player", LocX, vOffset)
        LocX = LocX + 2400
    End If
    If g_TeamName Then
        Set TeamTitle = m_Report.Sections(SectionNum).AddTextObject("Team", LocX, vOffset)
        LocX = LocX + 2500
    End If
    If g_Salary Then
        Set SalaryTitle = m_Report.Sections(SectionNum).AddTextObject("Salary", LocX, vOffset)
        LocX = LocX + 1900
    End If
    
    If g_GroupByTeam Then
        If g_TeamName = True Then
            ' If we're grouping by Team, add a feature that shows the total displayed payroll for
            ' each team.
            Set TeamGroupObj = m_Report.Sections(SectionNum).AddFieldObject(TeamFormula, 10, 0)
            Set TeamSummary = m_Report.Sections(SectionNum).AddSummaryFieldObject(SalaryFormula, crSTSum, 0, 450)
            Set PayrollTitle = m_Report.Sections(SectionNum).AddTextObject("Total Payroll Shown:", 0, 450)
            TeamSummary.Left = 2350
            TeamSummary.CurrencySymbolType = crCSTFloatingSymbol
            TeamSummary.DecimalPlaces = 0
            TeamGroupObj.TextColor = vbBlue
            TeamGroupObj.HasDropShadow = True
        End If
        If g_YearName = True Then
            Set YearHeading = m_Report.Sections(SectionNum).AddFieldObject(YearFormula, 0, 0)
            YearHeading.Left = 4500
            YearHeading.HorAlignment = crRightAlign
        End If
    Else
        If g_YearName = True Then Set YearTitle = m_Report.Sections(SectionNum).AddTextObject("Year", LocX, vOffset)
    End If

    ' Format all the text objects except for the title to look more appealing.
    ' Also don't try to format OLE objects (obj.Kind = 6)
    For Each obj In m_Report.Sections(SectionNum).ReportObjects
        If obj.Kind <> 6 And obj.Name <> "Title" Then
            obj.Width = 2000
            obj.Height = 350
            obj.Font.Name = "Arial"
            obj.Font.Size = 13
            If obj.Kind = crTextObject And obj.Left <> 0 Then
                ' Underline and change the color of the field headings.  All field headings
                ' have a left property set to a non-zero value.
                obj.TextColor = &H808000
                obj.Font.Underline = True
            Else
                ' Bold everything else
                obj.Font.Bold = True
            End If
        End If
    Next
    
    ' Increase the visibility of the Team name and payroll
    If g_TeamName = True And g_GroupByTeam = True Then
        TeamGroupObj.Width = 3700
        TeamGroupObj.Font.Size = 15
        PayrollTitle.Width = 3000
    End If
  
End Sub

' *************************************************************
' Add the database fields to the report.  This particular method
' creates unbound fields and binds those fields to the database at
' runtime.  Unbound fields are formula fields in disguise, so the
' techniques for manipulating unbound fields also apply to formulas
'
Private Sub DoDetails(SalaryFormula As FormulaFieldDefinition, _
                      TeamFormula As FormulaFieldDefinition, _
                      YearFormula As FormulaFieldDefinition)

    Dim PlayerFormula As FormulaFieldDefinition
    Dim PlayerObj As FieldObject
    Dim SalaryObj As FieldObject
    Dim TeamObj As FieldObject
    Dim YearObj As FieldObject
    Dim obj As Object
    Dim LocX As Long            ' The X-offset for the objects
    
    LocX = 10
    ' Add the formulas to the formula fields collection
    ' A field must be recurring if you want to group on a field or
    ' automatically bind it to a data source.
    ' Putting "WhileReadingRecords" into the formula text makes the formula
    ' into a recurring field.
    Const Recur = "WhileReadingRecords;" & vbNewLine
    Set PlayerFormula = m_Report.FormulaFields.Add("Player", Recur & "Space(10)")
    Set TeamFormula = m_Report.FormulaFields.Add("Team", Recur & "Space(10)")
    Set SalaryFormula = m_Report.FormulaFields.Add("Salary", Recur & "$0.0")
    Set YearFormula = m_Report.FormulaFields.Add("Year", Recur & "Space(10)")
    
    ' Add the fields only if the user wants them
    ' The player's name
    If g_PlayerName = True Then
        Set PlayerObj = m_Report.Sections(3).AddFieldObject(PlayerFormula, LocX, 0)
        PlayerObj.Name = "Player"
        PlayerObj.Width = 2600
        LocX = LocX + 2100
    End If
    
    ' Add the name of the team that the player belongs to
    If g_TeamName = True Then
        Set TeamObj = m_Report.Sections(3).AddFieldObject(TeamFormula, LocX, 0)
        TeamObj.Name = "Team"
        TeamObj.Width = 2800
        LocX = LocX + 2000
    End If
    
    ' Add the player's salary for that year
    If g_Salary = True Then
        Set SalaryObj = m_Report.Sections(3).AddFieldObject(SalaryFormula, LocX, 0)
        SalaryObj.Name = "Salary"
        SalaryObj.Width = 1500
        SalaryObj.CurrencySymbolType = crCSTFloatingSymbol
        SalaryObj.DecimalPlaces = 0
        LocX = LocX + 1600
    End If
    
    ' Add the year that the player made this amount of money
    If g_YearName = True Then
        Set YearObj = m_Report.Sections(3).AddFieldObject(YearFormula, LocX, 0)
        YearObj.Name = "Year"
        YearObj.Width = 2000
        YearObj.HorAlignment = crRightAlign
        If g_GroupByTeam Then YearObj.TextColor = vbWhite
    End If
    
    ' Format all the fields in the details section (Section 3) to a more
    ' attractive format
    For Each obj In m_Report.Sections(3).ReportObjects
        obj.Font.Name = "Arial"
        obj.Font.Size = 11
        obj.Height = 250
    Next
    
    ' Sort the fields ascending or descending, depending on the user's wishes
    ' Default to Descending order, unless the user wants to see those players
    ' that made less than a certain amount of money.
    If g_SortOrder = "ASC" Then
        m_Report.RecordSortFields.Add SalaryFormula, crAscendingOrder
    Else
        m_Report.RecordSortFields.Add SalaryFormula, crDescendingOrder
    End If
End Sub

' *************************************************************
' Add a footer to the bottom of each page.  The section number for the
' page varies according to how many grouping fields were added.
'
Private Sub DoFooters()
    Dim PageNumber As FieldObject
    Dim ln As LineObject
    Dim curSection As Integer
    Dim lineRight As Long
    Dim PageX As Long
        
    curSection = 5
    ' The page width for the report is set to a different value if there are no
    ' grouping options in the report
    If g_GroupByTeam Then
        If g_TeamName And g_YearName Then curSection = curSection + 1
        lineRight = 6500
        PageX = 2700
        m_Report.Sections(curSection).NewPageAfter = True
    Else
        lineRight = 8500
        PageX = 4000
    End If
                
    ' Add a line directly above the page number
    m_Report.Sections(curSection).Height = 700
    Set ln = m_Report.Sections(curSection).AddLineObject(10, 100, lineRight, 100)
    ln.LineColor = vbBlue
    ln.LineThickness = 4
    
    ' Add the page number to the footer
    Set PageNumber = m_Report.Sections(curSection).AddSpecialVarFieldObject(crSVTPageNofM, PageX, 0)
    PageNumber.Top = 200
    PageNumber.Font.Name = "Arial"
End Sub

' *************************************************************
' Change the paper size of the report to better center the report
' fields and group headers.
'
Private Sub DoReportSizing()
    If g_GroupByTeam Then
        m_Report.PaperSize = crPaperEnvelope14
    Else
        m_Report.PaperSize = crPaperEnvelopeB5
    End If
    m_Report.RightMargin = m_Report.LeftMargin
    m_Report.TopMargin = 100
End Sub

' *************************************************************
' Add an ADO data source to the report.  The PlayerSalaries.mdb
' database must be located in the same directory as this application
' The ADO data set is generated at run-time and does not persist after
' the report viewing is completed.
'
Private Sub AddDataSource()
    Dim TableName As String

    If g_SQLString = "" Then Exit Sub       ' Ensure we had an SQL string to work with
    
    Set m_Connection = New ADODB.Connection
    Set m_RS = New ADODB.Recordset
    ' Open the connection
    m_Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\PlayerSalaries.mdb;Mode=Read"
    m_Connection.Open
    
    ' Get all the records that the user wants -- this runtime dataset is assigned to the report
    m_RS.Open g_SQLString, m_Connection, adOpenDynamic, adLockPessimistic, adCmdText
    
    ' Add the record set to the report.
    m_Report.Database.AddADOCommand m_RS.ActiveConnection, m_RS.ActiveCommand
    
    ' Automatically bind the unbound fields to the open data set.  The name of the
    ' unbound field must be the same (case insensitive) as the database field.  The
    ' name of the unbound field can easily be changed at runtime to allow binding to
    ' a different database field.  Alternatively, the FieldName.SetUnboundFieldSource "Database Field Name"
    ' function could be used to manually bind the unbound field.
    m_Report.AutoSetUnboundFieldSource crBMTName
    
    If g_TopCount > 0 Then
        ' Add the "Top N" or "Greater/Less" than functionality
        TableName = m_Report.Database.Tables(1).Name
        If g_TypeofCount = vbTopN Then
            m_RS.MoveFirst
            m_RS.Move g_TopCount - 1
            If Not m_RS.EOF Then
                If g_SortOrder = "DESC" Then
                    m_Report.RecordSelectionFormula = "{" & TableName & ".Salary} >= $" & CStr(m_RS!Salary)
                Else
                    m_Report.RecordSelectionFormula = "{" & TableName & ".Salary} <= $" & CStr(m_RS!Salary)
                End If
            End If
        ElseIf g_TypeofCount = vbCompareSalary Then
            If g_SortOrder = "DESC" Then
                m_Report.RecordSelectionFormula = "{" & TableName & ".Salary} >= $" & CStr(g_TopCount)
            Else
                m_Report.RecordSelectionFormula = "{" & TableName & ".Salary} <= $" & CStr(g_TopCount)
            End If
        End If
    End If
    
    ' Finally, view the report in the Smart Viewer!
    crViewer.ViewReport
End Sub

' *************************************************************
' Initialize a few form variables and set the report generation
' in motion.
'
Private Sub Form_Load()
    AlreadyNoData = 0
    Screen.MousePointer = vbHourglass
    InitReport
End Sub

' *************************************************************
' The "central control station" for the report generation.
'
Private Sub InitReport()
    ' The Salary, Team and Year formulas are shared between the details
    ' and group header - declare them here.
    Dim SalaryFormula As FormulaFieldDefinition
    Dim TeamFormula As FormulaFieldDefinition
    Dim YearFormula As FormulaFieldDefinition
    
    ' Create a new report and set the viewer to point to this report
    Set m_Proj = New CRAXDRT.Application
    Set m_Report = m_Proj.NewReport
    crViewer.ReportSource = m_Report

    m_Report.Sections(1).Suppress = True
    DoReportSizing
    DoReportHeader
    DoPageHeader
    DoDetails SalaryFormula, TeamFormula, YearFormula
    DoGroupHeader SalaryFormula, TeamFormula, YearFormula
    DoFooters

    AddDataSource
    Screen.MousePointer = vbDefault
End Sub

' *************************************************************
' Cleanup any leftover residue from the report creation
'
Private Sub Form_Unload(Cancel As Integer)
    If Not (m_Report Is Nothing) Then Set m_Report = Nothing
    If Not (m_Proj Is Nothing) Then Set m_Proj = Nothing
    If Not (m_Connection Is Nothing) Then Set m_Connection = Nothing
    If Not (m_RS Is Nothing) Then
        m_RS.Close
        Set m_RS = Nothing
    End If
    m_StartForm.Enabled = True
End Sub

' *************************************************************
' If there are no records that match the user's criteria (for example,
' Players earning more than $100,000,000 per year), then let the user know.
'
Private Sub m_Report_NoData(pCancel As Boolean)
    If AlreadyNoData = 0 Then MsgBox "There are no players that match your criteria!", vbExclamation, "Pro Athlete Salary Report"
    AlreadyNoData = AlreadyNoData + 1
    pCancel = True
End Sub

' *************************************************************
Private Sub cmdExit_Click()
    Unload Me
End Sub

' *************************************************************
' Force the viewer to always display at 100%
'
Private Sub crViewer_ZoomLevelChanged(ByVal ZoomLevel As Integer)
    If ZoomLevel <> 100 Then crViewer.Zoom 100
End Sub

