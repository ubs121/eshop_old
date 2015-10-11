<%
if len(request.form) > 0 then
	strError = AddLogin()
end if
%>
<form action="" method="post" name="frmAddAdmin" id="frmAddAdmin">
<center>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Add Admin</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>

  <table width="500" border="0" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="250">Login:</td>
      <td width="250"><input name="admin_login" type="text" id="admin_login" value="<%=admin_login%>" size="40"></td>
    </tr>
    <tr> 
      <td colspan="2" height="10"></td>
    </tr>
    <tr> 
      <td>Password:</td>
      <td><input name="password1" type="password" id="password1" size="40"></td>
    </tr>
    <tr> 
      <td>Confirm password:</td>
      <td><input name="password2" type="password" id="password2" size="40"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2" height="10"></td>
    </tr>
<%
set rsModules = server.createobject("ADODB.recordset")
rsModules.cursortype = 3

strSQL = "SELECT module_id, module_name FROM admin_modules"
rsModules.open strSQL, adoCon

intCounter = 0
bgcolor = "#EEEEEE"
do while not rsModules.eof
	intCounter = intCounter + 1
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
	selected = cint(request.form("module_" & intCounter))
%>
    <tr align="center" style="background: <%=bgcolor%>"> 
      <td align="left">&nbsp;<%=rsModules("module_name")%>
        <input name="module_id_<%=intCounter%>" type="hidden" id="module_id_<%=intCounter%>" value="<%=rsModules("module_id")%>"></td>
      <td align="left">
	    <input name="module_<%=intCounter%>" type="radio" value="0"<% if selected = 0 then %> checked="checked"<% end if %>>
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
%>
    <tr align="center"> 
      <td colspan="2"><input name="intTotalModules" type="hidden" id="intTotalModules" value="<%=intCounter%>"> <%=BuildSubmitter("submit","Add administrator", request.querystring("p"))%>&nbsp; 
        <input name="Reset" type="reset" id="Reset" value="Reset"><br><br>
<% if len(strError) > 0 then %><font color="#CC3333" face=arial size=1><%=strError%></font><% end if %>

    </td>
    </tr>
  </table>
</td></tr>
</table>
</center>
</form>