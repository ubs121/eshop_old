<form name="form1" method="post" action="">
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

  <table width="500" align="center" cellpadding="2" cellspacing="2">
<%
admin_id = request.querystring("admin_id")
if len(admin_id) > 0 AND isNumeric(admin_id) then
	admin_id = cint(admin_id)
else
	response.redirect("?p=" & request.querystring("p"))
end if

if len(request.form("submit")) > 0 then
	UpdateAdmin()
end if

set rsModules = server.createobject("ADODB.recordset")
rsModules.cursortype = 3

strSQL = "SELECT module_id, module_name FROM admin_modules"
rsModules.open strSQL, adoCon

set rsRights = server.createobject("ADODB.recordset")
rsRights.cursortype = 3

strSQL = "SELECT module_right_id, module_id, module_right FROM admin_rights WHERE admin_id = " & admin_id
rsRights.open strSQL, adoCon

intCounter = 0
bgcolor = "#EEEEEE"
do while not rsModules.eof
	rsRights.filter = "module_id = " & rsModules("module_id")
	if not rsRights.eof then
		selected = cint(rsRights("module_right"))
		module_right_id = cint(rsRights("module_right_id"))
	else
		selected = 0
		module_right_id = 0
	end if
	intCounter = intCounter + 1
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
%>
    <tr align="center" style="background: <%=bgcolor%>"> 
      <td align="left">&nbsp;<%=rsModules("module_name")%> <input name="module_right_id_<%=intCounter%>" type="hidden" id="module_id_right_<%=intCounter%>" value="<%=module_right_id%>"></td>
      <td align="left"> <input name="module_<%=intCounter%>" type="radio" value="0"<% if selected = 0 then %> checked="checked"<% end if %>>
        None 
        <input type="radio" name="module_<%=intCounter%>" value="1"<% if selected = 1 then %> checked="checked"<% end if %>>
        Read 
        <input type="radio" name="module_<%=intCounter%>" value="2"<% if selected = 2 then %> checked="checked"<% end if %>>
        Write</td>
    </tr>
    <%
	rsModules.movenext
loop

rsModules.close
set rsModules = nothing

rsRights.close
set rsRights = nothing
%>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Change user settings",request.querystring("p"))%>
        &nbsp;
        <input type="button" name="Cancel" value="Cancel" onclick="document.location='?p=<%=request.querystring("p")%>';">
        <input name="intTotalModules" type="hidden" id="intTotalModules" value="<%=intCounter%>"></td>
    </tr>
  </table>
</td></tr>
</table>
</center>
</form>
