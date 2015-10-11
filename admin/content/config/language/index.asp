<%
set fs=Server.CreateObject("Scripting.FileSystemObject")
set fo=fs.GetFolder(server.mappath(strVirtualPath & "languages/"))

set rsLanguage = server.createobject("ADODB.recordset")
rsLanguage.cursortype = 3

strSQL = "SELECT language_id, language_name, language_folder, language_show, language_default FROM lang"
rsLanguage.open strSQL, adoCon
%>
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Language Setup</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>

<table width="500" border="0" align="center" cellpadding="1" cellspacing="0">
<tr><td>&nbsp;</td></tr>
  <%
for each x in fo.subfolders
	folder_name = x.Name
	rsLanguage.filter = "language_folder = '" & folder_name & "'"
	if rsLanguage.eof then
		installed = 0
		install_link = "?p=" & request.querystring("p") & "&amp;act=install&amp;folder=" & folder_name
		bg_color = "#CC3300"
		language_default = 0
	else
		if cint(rsLanguage("language_show")) = 0 then
			installed = 0
			bg_color = "#CC3300"
			install_link = "?p=" & request.querystring("p") & "&amp;act=install&amp;id=" & rsLanguage("language_id")
			language_default = 0
		else
			installed = 1
			bg_color = "#339900"
			language_default = cint(rsLanguage("language_default"))
		end if
	end if
%>
  <tr bgcolor="<%=bg_color%>"> 
    <td>&nbsp;
	  <b><%=folder_name%></b>
	  <% if language_default = -1 then %>
	  (default)
	  <% end if %>
	</td>
	<% if installed = 0 then %>
    <td width="150" align="right"><a href="<%=install_link%>"><img src="images/button_install.gif" width="71" height="21" alt="install" /></a></td>
	<% else %>
	<td width="150" align="right">
	<a href="?p=<%=request.querystring("p")%>&amp;act=edit&amp;id=<%=rsLanguage("language_id")%>"><img src="images/button_edit.gif" width="71" height="21" alt="edit" /></a>
	<% if cint(rsLanguage("language_default")) = 0 then %>
	<a href="?p=<%=request.querystring("p")%>&amp;act=uninstall&amp;id=<%=rsLanguage("language_id")%>"><img src="images/button_uninstall.gif" width="71" height="21" alt="uninstall" /></a>
	<% end if %>
	</td>
	<% end if %>
  </tr>
<% next %>
<tr><td>&nbsp;</td></tr>
</table>
</td></tr>
</table>
</center>

<%
set fo = nothing
set fs = nothing

rsLanguage.close
set rsLanguage = nothing
%>