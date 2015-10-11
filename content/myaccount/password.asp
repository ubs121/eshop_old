<%
intError  = 0
intSucces = 0

if request.form("Submit") = "ok" then
	password  = killChars(request.form("password"))
	password1 = killChars(request.form("new_password1"))
	password2 = killChars(request.form("new_password2"))
	
	if len(password1) = 0 then
		strError = strNewPasswordNotNull
		intError = 1
	end if
	
	if password1 <> password2 then
		strError = strPasswordsNotTheSame
		intError = 1
	end if
	
	if len(password) = 0 then
		strError = strPasswordNotNull
		intError = 1
	end if
	
	if intError = 0 then
		set rsPassword = server.createobject("ADODB.recordset")
		strSQL = "SELECT user_salt, user_password FROM users WHERE user_id = " & session("customer_id")
		
		rsPassword.open strSQL, adoCon, 2, 2
		
		strSecret = hashEncode(password & rsPassword("user_salt"))
		
		if strSecret = rsPassword("user_password") then
			rsPassword("user_password") = hashEncode(password1 & rsPassword("user_salt"))
			rsPassword.update()
			
			intSucces = 1
		else
			intError = 1
			strError = strPasswordNotCorrect
		end if
		
		rsPassword.close
		set rsPassword = nothing
	end if
end if
%>
<% if intError = 1 then %>
<table width="100%" cellspacing="0" cellpadding="0" class="error">
  <tr>
    <td class="content"><%=strError%>!</td>
  </tr>
</table>
<br />
<% end if %>
<% if intSucces = 1 then %>
<p><b><%=strPasswordUpdateSucces%></b></p>
<br />
<% end if %>
<form action="" method="post" name="frmPassword" id="frmPassword">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strCurrentPassword%>:</td>
            <td class="content"><input name="password" type="password" id="password" size="25"<% if intTelephone = 1 then %> class="required"<% end if %>> 
            </td>
          </tr>
          <tr> 
            <td width="160" class="content">&nbsp;</td>
            <td class="content">&nbsp;</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strNewPassword%>:</td>
            <td class="content"><input name="new_password1" type="password" id="new_password1" size="25"<% if intTelephone = 1 then %> class="required"<% end if %>></td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strConfirmNewPassword%>:</td>
            <td class="content"><input name="new_password2" type="password" id="new_password2" size="25"<% if intTelephone = 1 then %> class="required"<% end if %>></td>
          </tr>
        </table></td>
    </tr>
  </table>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            <td width="50%"><a href="?mod=myaccount"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strBack%>" width="122" height="22" border="0" /></a></td>
          <td width="50%" align="right"><input name="Submit" type="hidden" id="Submit" value="ok">
              <a href="javascript:document.frmPassword.submit();"> <img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
