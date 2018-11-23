<%
' 05/02/98
' Added the following features:
' Page Expiry Time
' -  The page will expire when downloaded by browser so that user is insured that all data
' will be current.
Response.Expires = 0
%>
<html>
<title></title>
<%
qs = request.querystring
if qs <> "" then
	qs = "&" & qs
end if 
%>
<frameset rows ="59,*">
<frame marginheight=0 marginwidth=0 noresize scrolling=no name="CrystalViewerToolbar" src="toolbar.asp?<% response.write request.querystring %>">
<frame name="CrystalViewerPreview" src="rptserver.asp?cmd=get_pg&viewer=html_frame&vfmt=html_frame&page=<%response.write session("CurrentPageNumber") %><% response.write qs %>">
</frameset>
</html>