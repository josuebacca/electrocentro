<%
' Get QS variables
dim reportPath
rpttoview = request.querystring("rpt")
viewer = request.querystring("init")

'build full path for report

rpttoview = MID(request.ServerVariables("PATH_TRANSLATED"), 1, (LEN(request.ServerVariables("PATH_TRANSLATED"))-41)) & "\Reports\General Business\" & rpttoview

' build path to MDB



If Not IsObject ( session ("oApp")) Then
Set session ("oApp") = Server.CreateObject("CrystalRuntime.Application")
End If

' Turn off all Error Message dialogs
'Set oGlobalOptions = Session ("oApp").Options
session ("oApp").SetMorePrintEngineErrorMessages(False)

' Open the report
If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if

Set session("oRpt") = session("oApp").OpenReport(rpttoview,1)

' Turn off sepecific report error messages
'Set oRptOptions = Session("oRpt").Options
session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False

' Opening the page engine will cause the data to be read
Set session("oPageEngine") = session("oRpt").PageEngine
if (UCase(Request.ServerVariables("HTTPS")) = "ON") then
	reportPath = "https://"
else
	reportPath = "http://"
end if
reportPath = reportPath + Request.ServerVariables("SERVER_NAME") + ":" + Request.ServerVariables("SERVER_PORT") + "/scrsamples/Web Component Server/rptserver.asp"
' Now decide what viewer to create
Select Case viewer

	Case "java"
%>
<html>
<head>
<title>Seagate Java Viewer using Browser's JVM</title>
</head>
<body bgcolor=C6C6C6>
<SCRIPT LANGUAGE="JavaScript"><!--
 	var _ns3 = false;
 	var _ns4 = false;
 	//--></SCRIPT>
 	<COMMENT><SCRIPT LANGUAGE="JavaScript1.1"><!--
 	var _info = navigator.userAgent;
 	var _ns3 = (navigator.appName.indexOf("Netscape") >= 0 && _info.indexOf("Win16") < 0 && _info.indexOf("Mozilla/3") >= 0);
 	var _ns4 = (navigator.appName.indexOf("Netscape") >= 0 && _info.indexOf("Win16") < 0 && _info.indexOf("Mozilla/4") >= 0 );
 	//--></SCRIPT></COMMENT>
 		<SCRIPT LANGUAGE="JavaScript"><!--
 			if(_ns3==true)
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=100%  archive="/viewer/JavaViewer/ReportViewer.zip">' );
 			else if (_ns4 == true)
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=100%  archive="/viewer/JavaViewer/ReportViewer.jar">' );
 			else
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=100%  >' );
 		//--></SCRIPT>

 		<param name=ReportName value="rptserver.asp">
		<param name=Language value="es">
		<param name=HasGroupTree value=true>
		<param name=ShowGroupTree value=true>
		<param name=HasRefreshButton value=true>
		<param name=HasPrintButton value=true>
		<param name=HasExportButton value=true>
 		<param name=cabbase value="/viewer/JavaViewer/ReportViewer.cab">
		</applet>

</body>
</html>

<%
	Case "actx"
%>

<HTML>
<HEAD>
<TITLE>Seagate ActiveX Viewer</TITLE>
</HEAD>
<BODY BGCOLOR=C6C6C6 LANGUAGE=VBScript ONLOAD="Page_Initialize">

<OBJECT ID="CRViewer"
	CLASSID="CLSID:C4847596-972C-11D0-9567-00A0C9273C2A"
	WIDTH=100% HEIGHT=95%
	CODEBASE="/viewer/activeXViewer/activexviewer.cab#Version=8,0,0,314">
<PARAM NAME="EnableRefreshButton" VALUE=1>
<PARAM NAME="EnableGroupTree" VALUE=1>
<PARAM NAME="DisplayGroupTree" VALUE=1>
<PARAM NAME="EnablePrintButton" VALUE=1>
<PARAM NAME="EnableExportButton" VALUE=1>
<PARAM NAME="EnableDrillDown" VALUE=1>
<PARAM NAME="EnableSearchControl" VALUE=1>
<PARAM NAME="EnableAnimationControl" VALUE=1>
<PARAM NAME="EnableZoomControl" VALUE=1>
</OBJECT>

<SCRIPT LANGUAGE="VBScript">
<!--
Sub Page_Initialize
	On Error Resume Next
	Dim webBroker
	Set webBroker = CreateObject("WebReportBroker.WebReportBroker")
	if ScriptEngineMajorVersion < 2 then
		window.alert "IE 3.02 users on NT4 need to get the latest version of VBScript or install IE 4.01 SP1. IE 3.02 users on Win95 need DCOM95 and latest version of VBScript, or install IE 4.01 SP1. These files are available at Microsoft's web site."
		CRViewer.ReportName = Location.Protocol + "//" + Location.Host + "/scrsamples/Web Component Server/rptserver.asp"
	else
		Dim webSource
		Set webSource = CreateObject("WebReportSource.WebReportSource")
		webSource.ReportSource = webBroker
		webSource.URL = Location.Protocol + "//" + Location.Host + "/scrsamples/Web Component Server/rptserver.asp"
		webSource.PromptOnRefresh = True
		CRViewer.ReportSource = webSource
	end if
	CRViewer.ViewReport
End Sub
-->
</SCRIPT>

</BODY>
</HTML>
<%
	Case "nav_plugin"
%>

<html>
<head>
<title>Seagate Report Viewer Plug-In</title>
</head>
<body bgcolor=C6C6C6>
<P align="center">
<script language="javascript">
document.writeln('<EMBED name="startup"');
document.writeln('type=application/x-ssreportviewer-plugin;version=8.0.0.2');
document.writeln('Pluginspage="/viewer/ActiveXViewer/get-npviewer.htm"');
document.writeln('Width=100% ');
document.writeln('Height=100%% ');
document.writeln('Param_URL="' + location.protocol + '//' + location.host + '/scrsamples/Web Component Server/rptserver.asp" ');
document.writeln('Param_EnableExportButton="true" ');
document.writeln('Param_EnableHelpButton="false" ');
document.writeln('Param_DisplayGroupTree="true" ');
document.writeln('Param_DisplayToolbar="true" ');
document.writeln('Param_EnableGroupTree="true" ');
document.writeln('Param_EnablePrintButton="true" ');
document.writeln('Param_EnableRefreshButton="true" ');
document.writeln('Param_EnableZoomControl="true" ');
document.writeln('>');
document.writeln('</EMBED>');
</script>
</p>
</body>
</html>
<%
	Case "java_plugin"
%>
<html>
<head>
<title>Seagate Java Viewer using Java Plug-in</title>
</head>
<body bgcolor=C6C6C6>

<P align="center">
<object
	classid="clsid:8AD9C840-044E-11D1-B3E9-00805F499D93"
	width=100%
	height=100%
	codebase="/viewer/JavaPlugin/Win32/jre1_2_2-win.exe#Version=1,2,2,0">
<param name=type value="application/x-java-applet;version=1.2.2">
<param name=code value="com.seagatesoftware.img.ReportViewer.ReportViewer">
<param name=codebase value="/viewer/JavaViewer/">
<param name=archive value="ReportViewer.jar">
<param name=Language value="es">
<param name=ReportName value="<%= reportPath %>">
<param name=ReportParameter value="">
<param name=HasGroupTree value="true">
<param name=ShowGroupTree value="true">
<param name=HasRefreshButton value="true">
<param name=HasPrintButton value="true">
<param name=HasExportButton value="true">
<param name=HasTextSearchControls value="true">
<param name=CanDrillDown value="true">
<param name=HasZoomControl value="true">
<param name=PromptOnRefresh value="true">
<comment>
<embed
	width=100%
	height=90%
	type="application/x-java-applet;version=1.2.2"
	pluginspage="/viewer/JavaPlugin/Win32/jre1_2_2-win.exe"
	java_code="com.seagatesoftware.img.ReportViewer.ReportViewer"
	java_codebase="/viewer/JavaViewer/"
	java_archive="ReportViewer.jar"
Language="es"
ReportName="<%= reportPath %>"
ReportParameter=""
HasGroupTree="true"
ShowGroupTree="true"
HasRefreshButton="true"
HasPrintButton="true"
HasExportButton="true"
HasTextSearchControls="true"
CanDrillDown="true"
HasZoomControl="true"
PromptOnRefresh="true"
></embed>
</comment>
</object>
</p>
</body>
</html>
<%
	Case "html_frame"
		response.redirect "htmstart.asp"

	Case "html_page"

	response.redirect "rptserver.asp"



	end select

%>
