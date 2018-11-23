<%
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
' Creating a new report using unbound fields
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
                                                                      
'
' CONCEPT:
'                                                                     
' This application is designed to demonstrate how to create a new report in an ASP
' application.  This ASP uses the concept of "Unbound Fields" where you can add a field
' to a report at runtime.

'===================================================================================
'Create the Crystal Reports Objects
'===================================================================================
'
'You will notice that the Crystal Reports objects are scoped as session variables.
'This is because the page on demand processing is performed by a prewritten
'ASP page called "rptserver.asp".  In order to allow rptserver.asp easy access 
'to the Crystal Report objects, we scope them as session variables.  That way
'any ASP page running in this session, including rptserver.asp, can use them.

reportname = Request.Form("ReportName")
GroupField = "{" & Request.Form("GroupField") & "}"
SortDirection = Request.Form("SortDirection")
FirstField = Request.Form("FirstField")
SecondField = Request.Form("SecondField")
ThirdField = Request.Form("ThirdField")

'This line creates a string variable called reportname that we will use to pass
'the Crystal Report filename (.rpt file) to the OpenReport method.
'To re-use this code for your application you would change the name of the report
'so as to reference your report file.

' CREATE THE APPLICATION OBJECT                                                                     
If Not IsObject (session("oApp")) Then                              
Set session("oApp") = Server.CreateObject("CrystalRuntime.Application")
End If                                                                

'This "if/end if" structure is used to create the Crystal Reports Application
'object only once per session.  Creating the application object - session("oApp")
'loads the Crystal Report Design Component automation server (craxdrt32.dll) into memory.
'
'We create it as a session variable in order to use it for the duration of the
'ASP session.  This is to elimainate the overhead of loading and unloading the
'craxdrt32.dll in and out of memory.  Once the application object is created in
'memory for this session, you can run many reports without having to recreate it.
                                                                      
' CREATE THE REPORT OBJECT                                     
'                                                                     
'The Report object is created by calling the Application object's OpenReport method.

Path = Request.ServerVariables("PATH_TRANSLATED")                     
While (Right(Path, 1) <> "\" And Len(Path) <> 0)                      
iLen = Len(Path) - 1                                                  
Path = Left(Path, iLen)                                               
Wend                                                                  
                                                                      
'This "While/Wend" loop is used to determine the physical path (eg: C:\) to the 
'Crystal Report file by translating the URL virtual path (eg: http://Domain/Dir)                                                                        

'OPEN THE REPORT (but destroy any previous one first)                                                     

If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if

Set session("oRpt") = session("oApp").NewReport

'This line uses the "PATH" and "reportname" variables to reference the Crystal
'Report file, and open it up for processing.
'
'Notice that we do not create the report object only once.  This is because
'within an ASP session, you may want to process more than one report.  The
'rptserver.asp component will only process a report object named session("oRpt").
'Therefor, if you wish to process more than one report in an ASP session, you
'must open that report by creating a new session("oRpt") object.

session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False

'These lines disable the Error reporting mechanism included the built into the
'Crystal Report Design Component automation server (craxdrt32.dll).
'This is done for two reasons:
'
'1.  The print engine is executed on the Web Server, so any error messages
'    will be displayed there.  If an error is reported on the web server, the
'    print engine will stop processing and you application will "hang".
'
'2.  This ASP page and rptserver.asp have some error handling logic desinged
'    to trap any non-fatal errors (such as failed database connectivity) and
'    display them to the client browser.
'
'**IMPORTANT**  Even though we disable the extended error messaging of the engine
'fatal errors can cause an error dialog to be displayed on the Web Server machine.
'For this reason we reccomend that you set the "Allow Service to Interact with Desktop"
'option on the "World Wide Web Publishing" service (IIS service).  That way if your ASP
'application freezes you will be able to view the error dialog (if one is displayed).

'====================================================================================
' Add Databases to the Report
'====================================================================================

Set Database = session("oRpt").Database
'Instantiate the Database Collection

Set Tables = Database.Tables
'Instantiate the Tables Collection

Call Tables.Add(Path + "..\..\..\Databases\xtreme.mdb", "Customer")
'This line adds the Customer Table from the Xtreme.mdb database file to the report.
'since the database is in a subdirectory 3 directories higher than this asp page
'we use the path and the DOS command to change up a directory

Set FirstTable = Tables.Item(1)
'Instantiate the First Table in the Tables Collection (only using one table in the report)
Set Fields = FirstTable.Fields
'Instantiate the Fields Collection for the table

For i = 1 to Fields.Count
	result = strcomp(Fields.Item(cint(i)).Name,cstr(GroupField))
	If cint(result) = 0 then
		Set GroupFieldIs = Fields.Item(cint(i))
	end if
next
'This loop iterates through each field in the table until the grouping field is found.
'We compare the field's name to the name of the field entered in the HTML page.
'Once we match the selected field with a field from the database we instantiate that field
'as an object that we pass (as the second parameter) to the AddGroup method


Select Case cstr(SortDirection)
	Case "Ascending"  intsortdirection = 0 
	Case "Descending" intsortdirection = 1
	Case "Original Order" intsortdirection = 2
End Select
'This case statement compares the sort direction entered by the user

session("oRpt").AddGroup 0, GroupFieldIs, 14, intsortdirection
'Add the group to the report using the GroupFieldIs object as the grouped field.

Select Case FirstField 
	Case "Customer.Customer ID"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 7, 720, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(1)
		FieldObject.SetUnboundFieldSource "{Customer.Customer ID}"
	Case "Customer.Customer Name"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 12, 720, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(1)
		FieldObject.SetUnboundFieldSource "{Customer.Customer Name}"
	Case "Customer.Last Year's Sales"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 8, 720, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(1)
		FieldObject.SetUnboundFieldSource "{Customer.Last Year's Sales}"
End Select

Select Case SecondField 
	Case "Customer.Customer ID"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 7, 2880, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(2)
		FieldObject.SetUnboundFieldSource "{Customer.Customer ID}"
	Case "Customer.Customer Name"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 12, 2880, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(2)
		FieldObject.SetUnboundFieldSource "{Customer.Customer Name}"
	Case "Customer.Last Year's Sales"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 8, 2880, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(2)
		FieldObject.SetUnboundFieldSource "{Customer.Last Year's Sales}"
End Select

Select Case ThirdField
	Case "Customer.Customer ID"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 7, 5040, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(3)
		FieldObject.SetUnboundFieldSource "{Customer.Customer ID}"
	Case "Customer.Customer Name"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 12, 5040, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(3)
		FieldObject.SetUnboundFieldSource "{Customer.Customer Name}"
	Case "Customer.Last Year's Sales"
		session("oRpt").Sections.Item(3).AddUnboundFieldObject 8, 5040, 0
		Set FieldObject = session("oRpt").Sections.Item(3).ReportObjects.Item(3)
		FieldObject.SetUnboundFieldSource "{Customer.Last Year's Sales}"
End Select
'Each Case statement compares the values for fields entered by the user.  Once we know which
'field is to be added, in which order, we then call the AddUnboundField method to add an
'an unbound field object for the field.  An unbound field is basically a "container" field
'that we can give the value of a database field.  Once the unbound field has been added to
'the report we then create a field object from the unbound field and finnaly assign
'the unbpund field a database field value
'selected.  Once the field object has been created

session("oRpt").Save Path & "\newreports\" & reportname

'====================================================================================
' Retrieve the Records and Create the "Page on Demand" Engine Object
'====================================================================================

On Error Resume Next                                                  
session("oRpt").ReadRecords                                           
If Err.Number <> 0 Then                                               
  Response.Write "Read Records Didn't Go"
Else

  If IsObject(session("oPageEngine")) Then                              
  	set session("oPageEngine") = nothing
  End If
set session("oPageEngine") = session("oRpt").PageEngine
End If                                                                 

' INSTANTIATE THE CRYSTAL REPORTS SMART VIEWER
'
'When using the Crystal Reports automation server in an ASP environment, we use
'the same page on demand "Smart Viewers" used with the Crystal Web Report Server.
'The are four Crystal Reports Smart Viewers:
'
'1.  ActiveX Smart Viewer
'2.  Java Smart Viewer
'3.  HTML Frame Smart Viewer
'4.  HTML Page Smart Viewer
'
'The Smart Viewer that you use will based on the browser's display capablities.
'For Example, you would not want to instantiate the Java viewer if the browser
'did not support Java applets.  For purposes on this demo, we have chosen to
'define a viewer.  You can through code determine the support capabilities of
'the requesting browser.  However that functionality is inherent in the Crystal
'Reports automation server and is beyond the scope of this demonstration app.
'
'We have chosen to leverage the server side include functionality of ASP
'for simplicity sake.  So you can use the SmartViewer*.asp files to instantiate
'the smart viewer that you wish to send to the browser.  Simply replace the line
'below with the Smart Viewer asp file you wish to use.
'
'The choices are SmartViewerActiveX.asp, SmartViewerJave.asp,
'SmartViewerHTMLFrame.asp, and SmartViewerHTMLPAge.asp.
'Note that to use this include you must have the appropriate .asp file in the 
'same virtual directory as the main ASP page.
'
'*NOTE* For SmartViewerHTMLFrame and SmartViewerHTMLPage, you must also have
'the files framepage.asp and toolbar.asp in your virtual directory.

viewer = Request.Form("Viewer")

'This line collects the value passed for the viewer to be used, and stores
'it in the "viewer" variable.

If cstr(viewer) = "ActiveX" then
%>
<!-- #include file="SmartViewerActiveX.asp" -->
<%
ElseIf cstr(viewer) = "Netscape Plug-in" then
%>
<!-- #include file="ActiveXPluginViewer.asp" -->
<%
ElseIf cstr(viewer) = "Java using Browser JVM" then
%>
<!-- #include file="SmartViewerJava.asp" -->
<%
ElseIf cstr(viewer) = "Java using Java Plug-in" then
%>
<!-- #include file="JavaPluginViewer.asp" -->
<%
ElseIf cstr(viewer) = "HTML Frame" then
	Response.Redirect("htmstart.asp")
Else
	Response.Redirect("rptserver.asp")
End If
'The above If/Then/Else structure is designed to test the value of the "viewer" varaible
'and based on that value, send down the appropriate Crystal Smart Viewer.
%>