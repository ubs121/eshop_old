<center>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Edit Admin</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>
<%
pp = request.querystring("pp")
q  = request.querystring("q")
if len(pp) > 0 AND IsNumeric(pp) then
	pp = cint(pp)
else
	pp = 1
end if

'Create query
	strSQL = "SELECT admin_id, admin_login FROM admin_users"
	if len(q) > 0 then
		strSQL = strSQL & " WHERE admin_login LIKE '%" & q & "%'"
	end if
	strSQL = strSQL & " ORDER BY admin_login ASC;"
	
'Open list
	set rsAdmins = server.createobject("ADODB.recordset")
	rsAdmins.cursortype = 3
	rsAdmins.open strSQL, adoCon
	
	rsAdmins.pagesize = 20
	pages  = rsAdmins.pagecount
	admins = rsAdmins.recordcount
	
	if not rsAdmins.eof then rsAdmins.absolutepage = pp
%>
<table width="500" align="center" cellpadding="2" cellspacing="2">
  <tr align="center"> 
    <td colspan="3">
	<form name="form1" method="get" action="">
        <input name="q" type="text" id="q" value="<%=request.querystring("q")%>">
        &nbsp; 
        <input type="submit" name="Submit" value="Search">
        <input name="p" type="hidden" id="p" value="<%=request.querystring("p")%>">
      </form></td>
  </tr>
<%
bgcolor = "#EEEEEE"
for x = 1 to 20
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
	if rsAdmins.eof then exit for
	admin_id    = rsAdmins("admin_id")
	admin_login = rsAdmins("admin_login")
%>
  <tr bgcolor="<%=bgcolor%>"> 
    <td width="400">&nbsp;<%=admin_login%></td>
    <td align="center">
	  <% if intModuleRights = 2 then %>
	  <a href="?p=<%=request.querystring("p")%>&amp;act=edit&amp;admin_id=<%=admin_id%>">edit</a>
	  <% else %>
	  edit
	  <% end if %>
	</td>
    <td align="center">
	  <% if intModuleRights = 2 then %>
	  <a href="?p=<%=request.querystring("p")%>&amp;act=delete&amp;admin_id=<%=admin_id%>">delete</a>
	  <% else %>
	  delete
	  <% end if %>
	</td>
  </tr>
<%
	rsAdmins.movenext
next
x = x - 1
display_record = (pp - 1) * 20
%>
</table>
<table width="500" cellspacing="2" cellpadding="2">
  <tr>
    <td width="250"><b>&nbsp;There are <%=admins%> administrators (displaying <%=display_record + 1%> - <%=display_record + x%>)</b></td>
	<td width="250" align="right">
	Pages: 
	<%
	for x = 1 to pages
		if x = pp then
			response.write("<b>" & x & "</b>&nbsp;")
		else
			response.write("<a href=""?p=" & request.querystring("p") & "&amp;pp=" & x & "&amp;q=" & q & """>" & x & "</a>&nbsp;")
		end if
	next
	%>
	</td>
  </tr>
</table>
</td></tr>
</table>
</center>
<%
rsAdmins.close
set rsAdmins = nothing
%>