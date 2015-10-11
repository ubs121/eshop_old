<%
intError = 1
if len(request.form("email")) > 0 then
	email = request.form("email")
	
	set rsEmail = server.createobject("ADODB.recordset")
	rsEmail.cursortype = 3
	
	strSQL = "SELECT user_email FROM users WHERE user_email = '" & email & "'"
	rsEmail.open strSQL, adoCon
	
	if rsEmail.eof then
		intEmail = 1
	else
		intEmail = 0
	end if
	
	rsEmail.close
	set rsEmail = nothing
	
	if intEmail = 0 then
		password = hexvalue(10)
		strSalt = getSalt(len(email))
		
		strSecret = hashEncode(password & strSalt)
		
		strSQL = "UPDATE users SET user_salt = '" & strSalt & "', user_password = '" & strSecret & "' WHERE user_email = '" & email & "'"
		adoCon.execute(strSQL)
		
		'Send email with password
		MailSubject = Replace(strLostPassSubject,"[shopname]",strShopName)
		MailBody    = Replace(strLostPassBody,"[shopname]",strShopName)
		MailBody    = Replace(MailBody,"[password]", password)
		MailTo      = email
		MailFrom    = strMailNoReply
		
		call SendMail()
		
		intError = 0
	end if
end if
%>
<% if intError = 0 then %>
<p align="center">
  <%=strNewPasswordSent%>
</p>
<% else %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
	<form action="<%=strCurrPage%>" method="post" name="frmLostPass">
        <p>
	  <%=strEnterEmailBelow%>
	</p>
	    <input name="email" type="text" value="<%=request.form("email")%>" />
        &nbsp;
        <input name="Submit" type="submit" id="Submit" value="<%=strSendNewPass%>" />
	</form>
	</td>
  </tr>
</table>
<% if intEmail = 1 then %>
<br />
<table width="100%" cellspacing="0" cellpadding="0" class="error">
  <tr>
    <td class="content"><%=strInvalidEmail%>!</td>
  </tr>
</table>
<% end if %>
<% end if %>