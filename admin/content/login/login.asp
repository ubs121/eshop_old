<%
if request.form("Submit") = "   Login   " then
	if DoLogin(request.form("LoginId"),request.form("Password")) = 1 then
		response.redirect("index.asp")
	end if
end if
%>
<form name="frmLogin" action="" method="post">
<center><br><br>
<table border="0" width="400" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Admin Login</font></b></td>
<td width="30" bgcolor="#6699CC" bordercolor="#6699CC" align="right">
<img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="400" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr>
<td width="150" valign="Top" Align="Right" bgcolor="#FFFFFF"><font size="2" face="Verdana"><br>Login ID:<br>Password:<br><br></font></td>
<td width="250" valign="Top" bgcolor="#FFFFFF" bgcolor="#FFFFFF"><br>
<input name="LoginId" type="text" id="LoginId" value="<%=request.form("LoginId")%>" size="25"><br>
<input name="Password" type="password" id="Password" size="25"></td>
</tr><tr><td colspan="2" Align="Right"><input type="submit" name="Submit" value="   Login   "></td></tr>
</table></td></tr>
</table>
</center>
</form>