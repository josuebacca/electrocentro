<%@ LANGUAGE="VBSCRIPT" %>

<%
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
' Crystal Report Parameter Fields
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
                                                                      
'
' CONCEPT:
'                                                                     
' This application is designed to demonstrate how to pass a value to a Crystal
' Report parameter field at runtime.  The parameter field has been placed in the
' "Report Header" section to serve as the title of the report.

'===================================================================================
'Create the Crystal Reports Objects
'===================================================================================
'
'You will notice that the Crystal Reports objects are scoped as session variables.
'This is because the page on demand processing is performed by a prewritten
'ASP page called "rptserver.asp".  In order to allow rptserver.asp easy access 
'to the Crystal Report objects, we scope them as session variables.  That way
'any ASP page running in this session, including rptserver.asp, can use them.

reportname = "ParameterField.rpt"

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

session("oRpt").DiscardSavedData

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
' Insert a value for the Crystal Report Parameter Field
'====================================================================================
                                                                     
set session("ParamCollection") = Session("oRpt").Parameterfields              
'This line creates an object to represent the collection of parameter
'fields that are contained in the report.                                                                      
																	  
set Param1 =  session("ParamCollection").Item(1)
'This line creates an object to reference the first parameter in the 
'report.  You can also use the parameter name in the Item() statement.

ParamValue = Request.Form("ParamValue")
'This line creates a temporary variable to store the value to pass to
'the paraemter field.

Call Param1.SetCurrentValue (CStr(ParamValue), 12)

'This line calls the SetCurrentValue method of the parameter field
'object to give the parameter field the value of the ParamValue variable.
'
'The first parameter passed to the SetCurrentValue method is the value
'that you want to pass to the Crystal Reports parameter field.  The
'second parameter is the datatype of the value you are passing to the 
'Crystal Parameter field.

'Since we are passing a VBScript variable as the value for the parameter
'we must cast the variable as the proper datatype before passing it to
'the report.  This is because VBScript variables are not datatyped unless
'cast.

'Below is a list of Crystal Data types, the VBScript cast functions and
'data type qualifier parameter to pass to the SetCurrentValue method:
     
'    Param Field Type    VBScript Call      Type Number
'
'    NumberField             CDbl                7
'    CurrencyField           CDbl                8
'    Boolean                 CBool               9
'    StringField             CStr                12
                                                                      
'*Note* that you can not use the CDate() function to cast a variable as
'a VB Date type, then qualify that variable as type number 10 to pass
'the value as a Crystal Date.  This is because a VB Date and a Crystal
'Date type are not the same.  The best solution is to pass the date value
'to the Crystal Parameter field as a string type and convert the strng
'to a date in Crystal Reports using a Crystal Formula Field.

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