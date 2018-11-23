<%@ LANGUAGE="VBSCRIPT" %>

<%
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
' Reporting directly off an ADO Recordset                                       
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
                                                                      
'
' CONCEPT:
'                                                                     
' This application is designed to demonstrate how to report on the
' contents of an ADO recordset.  We will first construct the ADO
' Connection and Recordset objects, then populate the recordset
' by passing an SQL statement to the database.  We will then
' construct the Crystal Reports objects, and point the report
' to the ADO recordset.  Finnaly we will send the Crystal Reports
' Smart Viewer to the client to display the report pages.


'  Step 1:  Create The ADO Connection and Recordset

'An ADO Database Connection is a connection layer to allow database access
'from applications such as Active Server Pages to your existing ODBC Data Source
'(DSN).  For the purposes of this example application we will use an ODBC System
'Datasource Name (System DSN) called "Xtreme Sample Data" which references the
'Crystal Reports demo Access database called Xtreme.mdb.

'Create the ADO Database Connection:

Set oConn = Server.CreateObject("ADODB.Connection")

'This line creates an ADO Connection Object named oConn.  We will
'use this oConn ADO connection object to connect to the ODBC DSN

'To use the oConn ADO connection we must first open it:

oConn.Open("Xtreme Sample Database")

'This line opens the connection to our ODBC datasource.  In this case our
'ODBC datasource points to an Access database file called Xtreme.mdb.

'Now we must create a Record Set object:

set session("oRs") = Server.CreateObject("ADODB.Recordset")

'The above line will create a session variable named session("oRs").  This variable
'will be our ADO Recordset and contain the data returned by an SQL "Select" statement

'Defining and populating the recordset:

session("oRs").ActiveConnection = oConn
'defines the ADO Connection Object the recordset will use

session("oRs").Open "SELECT [Product ID], [Product Name] FROM Product"

'populate the recordset by passing an SQL statement to ODBC, in this case
'we will be reporting two fields from the "Product" table of Xtreme.mdb

'===================================================================================
'Create the Crystal Reports Objects
'===================================================================================
'
'You will notice that the Crystal Reports objects are scoped as session variables.
'This is because the page on demand processing is performed by a prewritten
'ASP page called "rptserver.asp".  In order to allow rptserver.asp easy access 
'to the Crystal Report objects, we scope them as session variables.  That way
'any ASP page running in this session, including rptserver.asp, can use them.

reportname = "ADORecordset.rpt"

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

'======================================================================================
'======================================================================================
 
'Now we must tell the report to report off of the data in the ADO recordset:

'To base a report on data from a dynamically generated ADO recordset, we must
'build the report based on the data structure of the recordset we will create.
'Then at runtime, we tell the report to report off of the data in the ADO Record set.
'The report is currently created against a database structure file (ADORecordset.ttx)
'This ttx file contains the structure of the recordset, and not the actual data.

'A Crystal Report is completely dependant on the structure of the dataset the report
'will use.  Therefor it is important that your database structure (ttx file) or DSN
'reflects EXACTLY the data that is contained in the ADO recordset at runtime.

session("oRpt").DiscardSavedData
set Database = session("oRpt").Database
'Instantiates a database collection which references the database(s) used in the report.

set Tables = Database.Tables
'Instantiates a Tables collection which references the Tables of the Database object.

set Table1 = Tables.Item(1)
'Instantiates a table object which references the first table used in the report.
'In this case this table object currently refers to the ADORecordset.ttx file.

Table1.SetPrivateData 3, session("oRs") 

'The "SetPrivateData" line tells the report that it datasource is now the recordset
'Now the report will display the data contained in the session("oRs") record set.
'If your report contained a subreport that was based off this or a different reordset
'you must follow the same steps above only referencing the subreport object.
'
'
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