<%
if request.form("Submit") = "ok" then
	gender = killChars(request.form("gender"))
	firstname = killChars(request.form("firstname"))
	lastname = killChars(request.form("lastname"))
	dateofbirth = killChars(request.form("dateofbirth"))
	email = killChars(request.form("email"))
	telephone = killChars(request.form("telephone"))
	fax = killChars(request.form("fax"))
	intError = 0
	
	'Check required fields
	if len(firstname) = 0 then
		intFirstname = 1
		intError = 1
	end if
	
	if len(lastname) = 0 then
		intLastname = 1
		intError = 1
	end if
	
	if len(email) = 0 then
		intEmail = 1
		intError = 1
	end if

	if len(telephone) = 0 then
		intTelephone = 1
		intError = 1
	end if

	if intError = 0 then
		strSQL = "UPDATE users SET user_firstname = '" & firstname & "', user_lastname = '" & lastname & "', "
		strSQL = strSQL & "user_date_of_birth = '" & dateofbirth & "', user_email = '" & email & "', "
		strSQL = strSQL & "user_telephone = '" & telephone & "', user_fax = '" & fax & "', "
		strSQL = strSQL & "user_gender = " & gender & " WHERE user_id = " & session("customer_id")
		adoCon.execute(strSQL)
	end if
end if

set rsDetails = server.createobject("ADODB.recordset")
rsDetails.cursortype = 3

strSQL ="SELECT user_lastname, user_firstname, user_date_of_birth, user_email, user_telephone, user_fax, user_gender FROM users WHERE user_id = " & session("customer_id") 
rsDetails.open strSQL, adoCon

firstname   = rsDetails("user_firstname")
lastname    = rsDetails("user_lastname")
dateofbirth = rsDetails("user_date_of_birth")
email       = rsDetails("user_email")
gender      = rsDetails("user_gender")

telephone   = rsDetails("user_telephone")
fax         = rsDetails("user_fax")

rsDetails.close
set rsDetails = nothing
%>
<form action="" method="post" name="frmRegister">
  <p><b><%=strPersonalInfo%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strGender%>:</td>
            <td class="content"><input type="radio" name="gender" value="1"<% if gender="1" then %> checked="checked"<% end if %>> 
              <%=strMale%>&nbsp; <input name="gender" type="radio" value="2" <% if gender="2" then %> checked="checked"<% end if %>> 
              <%=strFemale%></td>
          </tr>
          <tr> 
            <td width="160" height="24" class="content"><%=strFirstName%>:</td>
            <td class="content"><input name="firstname" type="text" id="firstname" value="<%=firstname%>" size="25"<% if intFirstname = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strLastName%>:</td>
            <td class="content"><input name="lastname" type="text" id="lastname" value="<%=lastname%>" size="25"<% if intLastname = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strDateOfBirth%>:</td>
            <td class="content"><input name="dateofbirth" type="text" id="dateofbirth" value="<%=dateofbirth%>" size="25"></td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strEmail%>:</td>
            <td class="content"><input name="email_dummy" type="text" id="email_dummy" value="<%=email%>" size="25"<% if intEmail = 1 then %> class="required"<% end if %> disabled="disabled" /><input type="hidden" name="email" id="email" value="<%=email%>" />
              *</td>
          </tr>
        </table></td>
    </tr>
  </table>
  <p><b><%=strYourContactInformation%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strTelephone%>:</td>
            <td class="content"><input name="telephone" type="text" id="country" value="<%=telephone%>" size="25"<% if intTelephone = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strFax%>:</td>
            <td class="content"><input name="fax" type="text" value="<%=fax%>" size="25"></td>
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
            <td width="50%" align="right">
              <input name="Submit" type="hidden" id="Submit" value="ok">
              <a href="javascript:document.frmRegister.submit();"> <img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
