<%
' 05/02/98
' Added the following features:
' Tab Query String Parameter
'	- This is the selected tab's tabArray index value.
' Page Expiry Time
'	-  The page will expire when downloaded by browser so that user is insured that all data
' will be current.
' DrillDown Tabs
'	- Added in the session("tabArray") object to keep track of the drill down tabs.
' Search
'	- Added javascript window.alert function call to indicate when text is not found in rpt view.
' Goto Page Text Box
'	- Added textbox and filenew.gif so user can enter and request desired page number.
'	NOTE: Netscape 2.0 browsers do not call the on submit event handler when the image is selected.
'   Thus, the user will not be warned when incorrect data is entered into the goto page box.
'   This problem does not happen when the user selects return.
Response.Expires = 0
' Viewer Tab images
drilld = "<img border=0 src='/viewer/images/toolbar/pdrilld.gif' alt = 'Grupo principal'>"
drillu = "<img border=0 src='/viewer/images/toolbar/cdrillu.gif' alt = 'Grupo actual'>"
previewu = "<img border=0 src='/viewer/images/toolbar/pviewu.gif' alt = 'Vista'>"
previewd =  "<img border=0 src='/viewer/images/toolbar/pviewd.gif' alt = 'Vista'>"
' Set the correct numbers on the paging buttons
brch = request.querystring("BRCH")
if brch <> "" then
	brch = "&" & "brch=" & brch
	basepage = "<a href=" & chr(34) & "javascript:parent.parent.location='rptserver.asp?init=html_frame&page=1'" & chr(34) & ">"

end if

getPageCommand =  "rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe" & brch


searchFound = request.querystring("SEARCHFOUND")
if searchFound <> "" then
	if Cint(searchFound) = 0 then
		messageText = "onLoad = " & chr(34) & "window.alert('No se encuentra el texto en el informe.');" & chr(34)
	end if
end if

CurrentPageNumber = CStr(session("CurrentPageNumber"))
lastknownpage = CStr(session("lastknownpage"))
LastPageNumber = CStr(session("LastPageNumber"))

if CurrentPageNumber = "" then
	CurrentPageNumber = "1"
end if

if lastknownpage = "" then
	lastknownpage = "0"
end if



if LastPageNumber <> ""  and (CurrentPageNumber = LastPageNumber) then
	lastknownpage = CurrentPageNumber
	' remember the last known page
	session("lastknownpage") = CurrentPageNumber
	nextlink = ""
	lastlink = ""
	if CInt(CurrentPageNumber) > 1 then
		previouspage = CInt(CurrentPageNumber) - 1
		previouslink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=" & previouspage & brch & "'" & chr(34) & ">"
		firstlink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=1" & brch & "'" & chr(34) & ">"
	else
		previouslink = ""
		firstlink = ""
	end if
else
	if (CInt(lastknownpage) < CInt(CurrentPageNumber)) and LastPageNumber = "" then
		' remember the last known page
		session("lastknownpage") = CurrentPageNumber
		lastknownpage = CurrentPageNumber & "+"
	else
		if lastknownpage <> LastPageNumber then
			lastknownpage = lastknownpage & "+"
		end if
	end if
	if CInt(CurrentPageNumber) > 1 then
		previouspage = CInt(CurrentPageNumber) -1
		previouslink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=" & previouspage & brch & "'" & chr(34) & ">"
		firstlink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=1" & brch & "'" & chr(34) & ">"
	else
		previouslink = ""
		firstlink = ""
		previouspage = 1
	end if
	nextpage = CInt(CurrentPageNumber) + 1
	nextlink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=" & nextpage & brch & "'" & chr(34) & ">"
	lastlink = "<a href=" & chr(34) & "javascript:parent.location='rptserver.asp?cmd=toolbar%5Fpage&viewer=html%5Fframe&vfmt=html%5Fframe&page=32756" & brch & "'" & chr(34) & ">"
end if

%>


<html>
<script language="javascript">

function ValidateNumber(val, msg)
{

  if (val == "")
  {
    alert("Introduzca un valor por el campo " + msg);
    return (false);
  }

  var checkOK = "0123456789";
  var checkStr = val;
  var allValid = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Introduzca solamente caracteres dígitos en el campo " + msg);
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= "1"))
  {
    alert(" Introduzca un valor mayor que \"0\" en el campo " + msg);
    return (false);
  }
  return (true);
}


var currentValue = "<% response.write CurrentPageNumber %>";

function checkValue(){

	var pageNumber = document.forms[0].elements[0].value;
	if(!ValidateNumber(pageNumber, "Goto Page Number")){
		document.forms[0].elements[0].value = currentValue;
		parent.status = "Introduzca un valor numérico positivo. Sin espacios.";
		return false;
		}
	else
		// a new page will be downloaded with the next page number
		return true;

}

</script>
<body background="/viewer/images/toolbar/toolbg.gif" topmargin=0 leftmargin=0 <% response.write messageText%>>
<form method="POST" name=getPg target = "CrystalViewerPageFrame" action = <% response.write getPageCommand %> onSubmit="return checkValue();">
<table border=0 width=100% cellpadding=0 cellspacing=0 height=100%><tr nowrap>
<td nowrap align=right width=10%><% response.write firstlink %><img border=0 src="/viewer/images/toolbar/first.gif" alt="Primera página"><% response.write previouslink %><img border=0 src="/viewer/images/toolbar/prev.gif" alt="Página previa"></td>
<td nowrap valign=center align=center width=10%> <b> <%response.write CurrentPageNumber %> </b> de <%response.write lastknownpage %></td>
<td nowrap align=left width=10%><%response.write nextlink %><img border=0 src="/viewer/images/toolbar/next.gif" alt="Próxima página"></a><%response.write lastlink %><img border=0 src="/viewer/images/toolbar/last.gif" alt="Ultima página"></a></td>
<td align=left width=5%><input type=text value = "<%response.write CurrentPageNumber %>" size=4 maxlength=5 name=PAGE alt = "Ir a página" ></td>
<td align=left width=5%> <input type=image   src="/viewer/images/toolbar/filenew.gif" alt="Ir a página"></td>
</form>
<form method="POST" name=pf target="CrystalViewerPageFrame" action="rptserver.asp?cmd=srch&viewer=html%5Fframe&vfmt=html_frame&page=<%response.write CurrentPageNumber %>&dir=FOR&case=0<% response.write brch%>">
<td nowrap align=center width=15%><a href="javascript:parent.parent.location='rptserver.asp?cmd=rfsh&viewer=html%5Fframe&vfmt=html%5Fframe&page=<%response.write CurrentPageNumber %>'"><img border=0 src="/viewer/images/toolbar/refresh.gif" alt="Actualizar"></a></td>
<td align=right width=15%><input type=text size=10 maxlength=255 name=text></td>
<td align=left width=5%><input type=image src="/viewer/images/toolbar/search.gif" alt="Buscar texto"></td>
<td nowrap valign=bottom align=right width=20%>
<%
dim counter
dim tmpArray
dim upperBound
tmpArray = session("tabArray")
counter = Int(UBound(tmpArray) / 5)
if tmpArray(0) <> "EMPTY" then
		response.write drillu
		if counter > 0 then
			response.write "<a href=" & chr(34) & "javascript:parent.parent.location = 'htmstart.asp?tab=" & (counter * 5)  & "'" & chr(34) & ">"
			response.write drilld & "</a>"
		End If
		response.write "<a href=" & chr(34) & "javascript:parent.parent.location = 'htmstart.asp?tab=" & 0  & "'" & chr(34) & ">"
		response.write previewd & "</a>"
else
	response.write previewu
end if%>
</td>
</tr></table>
</form>
</body>
</html>
