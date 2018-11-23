<%
'This file contains the HTML code to instantiate the Smart Viewer Java.      
'
'You will notice that the Report Name parameter references the rptserver.asp file.
'This is because the report pages are actually created by rptserver.asp.
'Rptserver.asp accesses session("oApp"), session("oRpt") and session("oPageEngine")
'to create the report pages that will be rendered by the ActiveX Smart Viewer.
'
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