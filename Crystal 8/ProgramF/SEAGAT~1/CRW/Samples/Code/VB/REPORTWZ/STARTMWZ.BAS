Attribute VB_Name = "StartModule"
' *************************************************************
' Purpose:  Demonstrate the use of the Crystal Report Wizard
'           (new to Crystal Reports 8.0).
'
'           The Report Wizard is a very powerful new feature that allows
'           the user to create an entirely new report at runtime
'           and then view or save this report.  The report wizard
'           can be used as a starting point for a custom report wizard,
'           allowing the programmer to set their own defaults for
'           report creation.
'
'           Try experimenting with the wizard by adding database fields,
'           text objects, unbound fields, graphs and formulas to the report
'           before calling the wizard.
'

Option Explicit

Dim m_App As New CRAXDRT.Application
Dim m_Report As New CRAXDRT.Report
Dim m_Wizard As New CRStandardWizard

' *************************************************************
' Start the Crystal Reports Wizard using data based on the
' "Xtreme Sample Database"
'
Sub Main()
    Dim tblCustomer As DatabaseTable
    Dim tblOrders As DatabaseTable
    Dim fldDef1 As DatabaseFieldDefinition
    Dim fldDef2 As DatabaseFieldDefinition
    Dim CustomerIDLink As TableLink
    
    ' Create a blank report
    Set m_Report = m_App.NewReport
    
    With m_Report.Database
        ' Add a datasource to the report.  The wizard will include any data
        ' source that has been added to the report
        .AddOLEDBSource DataEnvironment1.XtremeConnection, "Customer"
        .AddOLEDBSource DataEnvironment1.XtremeConnection, "Orders"
        
        ' Note that if you are adding multiple tables to a report and you want
        ' the tables to be linked, you must add the links before calling the
        ' Report Creation wizard.  There is currently no visual linking done
        ' in the wizard.
        ' Get references to the database tables
        Set tblCustomer = .Tables(1)
        Set tblOrders = .Tables(2)
        ' Get references to each table's Customer ID field
        Set fldDef1 = .Tables(1).Fields(1)
        Set fldDef2 = .Tables(2).Fields(3)
        
        ' Add a link between the two tables
        Set CustomerIDLink = .Links.Add(tblCustomer, tblOrders, _
                                        fldDef1, fldDef2, crJTEqual, _
                                        crLTLookupParallel, False, True)
    End With
    ' Give the report wizard a starting report definition to work with.
    ' If you give it a report that includes other report objects such as fields
    ' or text objects, the report wizard will incorporate those pre-existing
    ' objects into the final report.  In this case, we're giving the wizard
    ' a blank report with some database objects to work with.
    Set m_Wizard.CrystalReport = m_Report
    m_Wizard.DisplayReportWizard
End Sub

