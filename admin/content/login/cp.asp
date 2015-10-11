<%
if len(request.form("Submit")) > 0 then
	strMessage = AdjustLogin()
end if

set rsLogin = server.createobject("ADODB.recordset")
rsLogin.cursortype = 3

strSQL = "SELECT admin_login FROM admin_users WHERE admin_id = " & session("admin_user_id")
rsLogin.open strSQL, adoCon

admin_login = rsLogin("admin_login")

rsLogin.close
set rsLogin = nothing
%>
<form action="" method="post" name="frmCp" id="frmCp">
<center>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Change Password</font></b></td>
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
      <td width="250"><input name="Login_fake" type="text" id="Login" value="<%=admin_login%>" size="40" disabled="disabled">
        <input name="Login" type="hidden" id="Login" value="<%=admin_login%>"></td>
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
      <td colspan="2"><input type="submit" name="Submit" value="Update">
        &nbsp;
        <input name="Reset" type="reset" id="Reset" value="Reset"><br><br>
<% if len(strMessage) > 0 then %><font color="#CC3333" face=arial size=1><%=strMessage%></font><% end if %>
    </td>
    </tr>
  </table>
</td></tr>
</table>
</center>
</form>
