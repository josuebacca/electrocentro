<%@ LANGUAGE="VBSCRIPT" %>

<%
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
' Exporting a Crystal Report to a different file format
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
                                                                      
'
' CONCEPT:
'                                                                     
' This application is designed to demonstrate how to export a Crystal Report
' to a different file format, then point the browser to the new file.

'===================================================================================
'Create the Crystal Reports Objects
'===================================================================================
'
'You will notice that the Crystal Reports objects are scoped as session variables.
'This is because the page on demand processing is performed by a prewritten
'ASP page called "rptserver.asp".  In order to allow rptserver.asp easy access 
'to the Crystal Report objects, we scope them as session variables.  That way
'any ASP page running in this session, including rptserver.asp, can use them.

reportname = "ExportReport.rpt"

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
' Exporting the Crystal Report
'====================================================================================

session("filename") = Request.Form("filename")
session("filetype") = Request.Form("filetype")
'These lines collect the values passed from the calling HTML page for the export
'type and filename

Set session("ExportOptions") = Session("oRpt").ExportOptions
'First we create an export options collection which allows us access
'to the exporting properties of the automation server.

session("ExportFileName") = Path + "exports\" + cstr(session("filename"))
'The filename we pass to Crystal Reports must include a path.


Select Case cstr(session("filetype")) 
	Case "Crystal Report" ExportType = 1 
	Case "Microsoft Word" ExportType = 14
	Case "Microsoft Excel" ExportType = 21
	Case "Microsoft Excel (tabular)" ExportType = 22
	Case "HTML" ExportType = 24
	Case "Paginated Text" ExportType = 10
	Case "Rich Text Format" ExportType = 4
	Case "Text" ExportType = 8
	Case "Tab Seperated Values" ExportType = 6
	Case "Comma Seperated Values" ExportType = 5
	Case "Adobe PDF" ExportType = 31
End Select
'This Select/Case structure uses the value passed for the export format
'type and sets the ExportType variable to the appropriate "Format Type"
'integer value.  For the complete list of Format, and Destinatation types
'you can search for "ExportOptions" in the Developer's Help.

If cint(ExportType) <> 24 then
    session("ExportOptions").DiskFileName = session("ExportFileName")
Else
    session("ExportOptions").HTMLFileName = session("ExportFileName")
End If

If cint(ExportType) = 10 then
    session("ExportOptions").NumberofLinesPerPage = 50
End if
'If the export format is paginated text we must set the number of
'records that appear per page.

'We must test to see if the report is to be exported to HTML or another
'file type.  If the report is exported to HTML, we must use the HTML File
'name instead of Disk file name.

session("ExportOptions").FormatType = cint(ExportType)
'This line sets the file type that the report will be exported to.  We
'use the value that we collected from the calling HTML page.

session("ExportOptions").DestinationType = 1
'This line sets the destination of the exported file.  1 means that
'we are writing the file to disk.

Session("oRpt").Export False

session("ExportVirtualDirectory") = "/scrsamples/Active%20Server%20Pages/exports/"
'The ExportVirtualDirectory variable is used to reference the a web directory
'so that we can point the browser to the newly exported file.

'====================================================================================
' Point the browser to the newly exported file
' (but only if the file is previewable by the browser)
'====================================================================================

Select Case session("ExportOptions").FormatType
	Case 14
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 21
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 22
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 24
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 10
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 4
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case 8
	    Response.Redirect session("ExportVirtualDirectory") & cstr(session("filename"))
	Case Else
	    Response.Redirect("Exported.asp")
End Select
%>