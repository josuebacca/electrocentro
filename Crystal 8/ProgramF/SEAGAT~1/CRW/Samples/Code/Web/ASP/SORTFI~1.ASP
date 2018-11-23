<%@ LANGUAGE="VBSCRIPT" %>

<%
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
' Changing Sort Criteria
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
                                                                      
'
' CONCEPT:
'                                                                     
' This application is designed to demonstrate how to add a sort criteria at runtime
' and modify an existing sort criteria.  The demo report used in this application
' has a subreport in it.  The subreport has a Sort field already in it, while the
' main report does not.  We will therefor add a sort field to the main report
' and modify the sort field of the subreport.

'===================================================================================
'Create the Crystal Reports Objects
'===================================================================================
'
'You will notice that the Crystal Reports objects are scoped as session variables.
'This is because the page on demand processing is performed by a prewritten
'ASP page called "rptserver.asp".  In order to allow rptserver.asp easy access 
'to the Crystal Report objects, we scope them as session variables.  That way
'any ASP page running in this session, including rptserver.asp, can use them.

reportname = "SortMain.rpt"

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

Set session("oRpt") = session("oApp").OpenReport(path & reportname, 1)

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
' Add a Sort Field to the Main Report
'====================================================================================
sortdirection = Request.Form("direction")
sortfield = Request.Form("field")

Set Database = session("oRpt").Database
'Instantiate the Database Collection

Set Tables = Database.Tables
'Instantiate the Tables Collection
Set FirstTable = Tables.Item(1)
'Instantiate the First Table in the Tables Collection (only using one table in the report)
Set Fields = FirstTable.Fields
'Instantiate the Fields Collection for the table
For i = 1 to Fields.Count
	result = strcomp(Fields.Item(cint(i)).Name,cstr(sortfield))
	If cint(result) = 0 then
		Set SortFieldIs = Fields.Item(cint(i))
	end if
next
'This loop compares the name of each field in the table to the sort field selected in
'the preceding HTML page.  If the field name match the Sort Field Object is set

Set session("Sortfields") = session("oRpt").RecordSortFields
'We start by instantiating a Sort Fields collection.  The session("SortFields")
'object represents all the sort fields in the report.

Select Case cstr(sortdirection) 
	Case "Ascending"  intsortdirection = 0 
	Case "Descending" intsortdirection = 1
	Case "Original Order" intsortdirection = 2
End Select
'Get the sort direction constant for the direction set in the preceding HTML page

session("SortFields").Add SortFieldIs, cint(intsortdirection)
'Since the main report does not have a sort field, we must use the add() method
'to sort the main report's records.  Both sortdirection and sortfield are variables
'that contain the values associated with the sort direction (ascending, descending or
'original order) and the database field to base the sort.  We collect these values
'from the calling HTML page.

'====================================================================================
' Modify Sort Field in the Subreport
'====================================================================================
subsortdirection = Request.Form("SubDirection")
'Collect the value for the subreport's sort direction, from the calling HTML page.

'First we must open the subreport, in order to have access to its Collections,
'Objects, Properties and Methods.  To do this we use the OpenReport method.

Set session("oSubRpt") = session("oRpt").OpenSubreport("SortSub.rpt")
'Now we have a report object that represents the subreport

Set session("SubSortfields") = session("oSubRpt").RecordSortFields
'Again, we start by instantiating a Sort Fields collection.  The session("SubSortFields")
'object represents all the sort fields in the subreport.

set sortfield1 = session("SubSortfields").Item(1)
'This instantiates an Object representing the first (and only) sort field in the subreport

Select Case cstr(subsortdirection) 
	Case "Ascending"  intsubsortdirection = 0 
	Case "Descending" intsubsortdirection = 1
	Case "Original Order" intsubsortdirection = 2
End Select

sortfield1.SortDirection = cint(intsubsortdirection)
'We now set the sort direction for the subreport based on the user's input

'====================================================================================
' Retrieve the Records and Create the "Page on Demand" Engine Object
'====================================================================================

On Error Resume Next                                                  
session("oRpt").ReadRecords                                           
If Err.Number <> 0 Then                                               
  Response.Write "An Error has occured on the server in attempting to access the data source"
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
