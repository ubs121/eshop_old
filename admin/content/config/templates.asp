<%
if len(request.form("Submit")) > 0 then
	updateConfig("default_stylesheet")
	
	session("stylesheet") = request.form("default_stylesheet")
end if
%>
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
    <td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Templates</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
  <td>
  <%
  set rsTemplate = server.createobject("ADODB.recordset")
  rsTemplate.cursortype = 3
  
  strSQL = "SELECT config_value FROM config WHERE config_name = 'default_stylesheet';"
  rsTemplate.open strSQL, adoCon
  
  if not rsTemplate.eof then
  	stylesheet = rsTemplate("config_value")
  end if
  
  rsTemplate.close
  set rsTemplate = nothing
  
  set fs=Server.CreateObject("Scripting.FileSystemObject")
  set fo=fs.GetFolder(server.mappath(strVirtualPath & "includes/styles/"))

  %>
  <br />
  <form name="frmTemplates" action="" method="post">
  <table width="500" align="center" cellspacing="0" cellpadding="0" border="0">
  	<%
	intCounter = 1
	for each x in fo.files
		if right(x.name, 3) = "css" then
			if intCounter = 1 then
				response.write("<tr>" & chr(10))
			end if
			folder = left(x.name, len(x.name) - 4)
			response.write("      <td width=""125"" align=""center"">")
			response.write("<img src=""" & strVirtualPath & "/includes/styles/" & folder & "/preview.gif"" width=""115"" height=""65"" alt=""" & x.name & """ style=""border: solid 1px #000000; padding: 1px;"" /><br />")
			response.write("<input name=""default_stylesheet"" type=""radio"" value=""" & folder & """ ")
			if folder = stylesheet then
				response.write("checked=""checked"" ")
			end if
			response.write("/><br />")
			response.write(folder)
			response.write("</td>" & chr(10))
			if intCounter = 4 then
				response.write("</tr>" & chr(10))
				intCounter = 1
			else
				intCounter = intCounter + 1
			end if
		end if
	next
	x = 0
	if intCounter > 1 then
		if intCounter = 4 then
			response.write("      <td>&nbsp;</td>" & chr(10))
		else
			response.write("      <td colspan=""" & (5 - intCounter) & """>&nbsp;</td>" & chr(10))
		end if
		response.write("</tr>" & chr(10))
	end if
	%>
	<tr>
	   <td colspan="4" align="center"><br />
	     <%=BuildSubmitter("submit","Update", request.querystring("p"))%> &nbsp; 
         <input type="button" name="Button" value="Back" onclick="document.location='?p=1';">
	  </td>
	</tr>
  </table>
  </form>
  <br />
  </td>
</tr>
</table>
</center>
