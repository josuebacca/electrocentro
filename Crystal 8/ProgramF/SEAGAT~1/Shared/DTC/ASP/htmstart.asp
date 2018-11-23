<%
' 05/02/98
' Added the following features:
' Tab Query String Parameter
' - This is the selected tab's tabArray index value.  
' Page Expiry Time
' -  The page will expire when downloaded by browser so that user is insured that all data
' will be current.
Response.Expires = 0
On Error Resume Next
qs = request.querystring
if qs <> "" then
	qs = "&" & qs
else
	' Need to make this call for backward compatibility.  Users may be referencing htmstart.asp in their web pages.
	Call InitializeFrameArray
end if 

dim tmpArray
dim index
dim brch
dim val
if request.querystring("TAB") <> "" then
	tmpArray = session("tabArray")
	index = Cint(request.querystring("TAB"))
	
	if tmpArray(index + 1) <> "" then
		 brch = tmpArray(index + 1)
		 qs = "&" & "BRCH=" & brch
	end if
	
	session("CurrentPageNumber") = tmpArray(index + 2)
	session("lastknownpage") = tmpArray(index + 3)
	session("LastPageNumber") = tmpArray(index + 4)
	' clear out all of the other arrays
	if index = 0 then 
		Call InitializeFrameArray
	else
		redim preserve tmpArray(index - 1)
		session("tabArray") = tmpArray
	end if
		

else
	session("CurrentPageNumber") = "1"
	session("lastknownpage") = "0"
	session("LastPageNumber") = ""
end if

'  This function initializes the tabArray. 
SUB InitializeFrameArray()
'initialize the html_frame array
	set session("tabArray") = Nothing
	session("lastBrch") = ""
	dim tmpArray
	tmpArray = Array(4)
	redim tmpArray(4)
	'Initialize the sequence number
	tmpArray(0) = "EMPTY"
	session("tabArray") = tmpArray
END SUB


%>

<html>
<title></title>
<frameset cols="15%,*">
<frame name="CrystalViewerTree" src="rptserver.asp?cmd=get_ttl&viewer=html%5Fframe&vfmt=html_frame<% response.write qs %>">
<frame name="CrystalViewerPageFrame" src="rptserver.asp?cmd=toolbar_page&viewer=html%5Fframe&vfmt=html_frame&page=<% response.write session("CurrentPageNumber") %><% response.write qs %>">
</frameset>
</html>
