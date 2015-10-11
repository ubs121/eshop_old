<%
if request.form("Submit") = "ok" then 
	gender = killChars(request.form("gender"))
	firstname = killChars(request.form("firstname"))
	lastname = killChars(request.form("lastname"))
	dateofbirth = killChars(request.form("dateofbirth"))
	email = killChars(request.form("email"))
	companyname = killChars(request.form("companyname"))
	vat         = killChars(request.form("vat"))
	street = killChars(request.form("street"))
	postcode = killChars(request.form("postcode"))
	city = killChars(request.form("city"))
	province = killChars(request.form("province"))
	country = killChars(request.form("country"))
	telephone = killChars(request.form("telephone"))
	mobile    = killChars(request.form("mobilePhone"))
	fax = killChars(request.form("fax"))
	password1 = killChars(request.form("password1"))
	password2 = killChars(request.form("password2"))
	newsletter = request.form("newsletter")
	intError = 0
	SecurityCode = killChars(request.form("SecurityCode"))
	
	if len(gender) > 0 then
		gender = cint(gender)
	end if
	
	'Check required fields
	if cstr(SecurityCode) <> cstr(session("SecurityCode")) then
		intError = 1
		intSecurityCode = 1
	end if
	
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
	
	if len(street) = 0 then
		intStreet = 1
		intError = 1
	end if
	
	if len(postcode) = 0 then
		intPostcode = 1
		intError = 1
	end if
	
	if len(city) = 0 then
		intCity = 1
		intError = 1
	end if
	
	if len(province) = 0 then
		intProvince = 1
		intError = 1
	end if
	
	if len(country) = 0 then
		intCountry = 1
		intError = 1
	end if
	
	if len(telephone) = 0 AND len(mobile) = 0 then
		intTelephone = 1
		intError = 1
	end if
	
	if password1 <> password2 OR len(password1) = 0 then
		intPassError = 1
		intError = 1
	end if
	
	if intError = 0 then
		set rsUsers = server.createobject("ADODB.recordset")
		rsUsers.cursortype = 3
		strSQL = "SELECT user_id FROM users WHERE user_email = '" & email & "'"
		rsUsers.open strSQL, adoCon
		
		intEmailExists = 0
		if not rsUsers.eof then
			intEmailExists = 1
			intError = 1
		end if
		
		rsUsers.close
		set rsUsers = nothing
		
		set rsUsers = server.Createobject("ADODB.recordset")
		rsUsers.open "users", adoCon, 2, 2
		
		if intEmailExists = 0 then
			if newsletter = "1" then
				set rsNewsletter = server.createobject("ADODB.recordset")
				strSQL = "SELECT user_email, user_lang_id FROM newsletter WHERE user_email = '" & email & "'"
				rsNewsletter.open strSQL, adoCon, 2, 2
				
				if not rsNewsletter.eof then
					rsNewsletter("user_lang_id") = session("language_id")
					rsNewsletter.update()
				else
					rsNewsletter.addnew()
						rsNewsletter("user_email") = email
						rsNewsletter("user_lang_id") = session("language_id")
					rsNewsletter.update()
				end if
				
				rsNewsletter.close
				set rsNewsletter = nothing
			end if
			strSalt = getSalt(len(email))
			strSecret = hashEncode(password1 & strSalt)
			
			rsUsers.addNew()
				rsUsers("user_firstname") = firstname
				rsUsers("user_lastname") = lastname
				rsUsers("user_date_of_birth") = dateofbirth
				rsUsers("user_gender") = gender
				rsUsers("user_salt") = strSalt
				rsUsers("user_password") = strSecret
				rsUsers("user_email") = email
				rsUsers("user_telephone") = telephone
				rsUsers("user_fax") = fax
				rsUsers("user_mobile") = mobile
			rsUsers.update()
			
			user_id = rsUsers("user_id")
			
			strSQL = "INSERT INTO user_address (user_street, user_postcode, user_city, user_province, user_country, user_company_name, user_id, user_firstname, user_lastname, user_default_address, user_vat) VALUES('"
			strSQL = strSQL & street & "','" & postcode & "','" & city & "','" & province & "','" & country & "','" & companyname & "'," & user_id & ",'" & firstname & "','" & lastname & "',-1, '" & vat & "');"
			
			adoCon.execute(strSQL)
		end if
		
		rsUsers.close
		set rsUsers = nothing
		
		if intError = 0 then
			session("regSucces") = true
			response.redirect("?mod=myaccount&sub=login&red=" & request.querystring("red"))
		end if
	end if
end if

'Create a security code
	Randomize()
	session("SecurityCode") = round((99999 * Rnd()))
	for x = 1 to (5 - len(session("SecurityCode")))
		session("SecurityCode") = "0" & session("SecurityCode")
	next
%>
<p align="right">* <%=strRequiredFields%></p>
<% if intEmailExists = 1 then %>
<table width="100%" cellspacing="0" cellpadding="0" class="error">
  <tr>
    <td class="content"><%=strEmailExists%>!</td>
  </tr>
</table>
<% end if %>
<% if intSecurityCode = 1 then %>
<table width="100%" cellspacing="0" cellpadding="0" class="error">
  <tr>
    <td class="content"><%=strInvalidSecurityCode%>!</td>
  </tr>
</table>
<% end if %>
<form action="<%=strCurrPage%>" method="post" name="frmRegister"><p><b><%=strPersonalInfo%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strGender%>:</td>
            <td class="content"><input type="radio" name="gender" value="1"<% if gender="1" then %> checked="checked"<% end if %>>
              <%=strMale%>&nbsp;
              <input name="gender" type="radio" value="2" <% if gender="2" then %> checked="checked"<% end if %>>
              <%=strFemale%></td>
          </tr>
          <tr> 
            <td width="160" height="24" class="content"><%=strFirstName%>:</td>
            <td class="content"><input name="firstname" type="text" id="firstname" value="<%=firstname%>" size="25"<% if intFirstname = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strLastName%>:</td>
            <td class="content"><input name="lastname" type="text" id="lastname" value="<%=lastname%>" size="25"<% if intLastname = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strDateOfBirth%>:</td>
            <td class="content"><input name="dateofbirth" type="text" id="dateofbirth" value="<%=dateofbirth%>" size="25"></td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strEmail%>:</td>
            <td class="content"><input name="email" type="text" id="email" value="<%=email%>" size="25"<% if intEmail = 1 then %> class="required"<% end if %>>*</td>
          </tr>
        </table></td>
  </tr>
</table>
<p><b><%=strCompanyInfo%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strCompanyName%>:</td>
            <td class="content"><input name="companyname" type="text" value="<%=companyname%>" size="25"></td>
          </tr>
          <tr>
            <td class="content"><%=strVAT%></td>
            <td class="content"><input name="vat" type="text" id="vat" value="<%=vat%>" size="25"></td>
          </tr>
        </table></td>
  </tr>
</table>
<p><b><%=strYourAddress%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strStreet%>:</td>
            <td class="content"><input name="street" type="text" value="<%=street%>" size="25"<% if intStreet = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strPostCode%>:</td>
            <td class="content"><input name="postcode" type="text" value="<%=postcode%>" size="25"<% if intPostcode = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strCity%>:</td>
            <td class="content"><input name="city" type="text" value="<%=city%>" size="25"<% if intCity = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strProvince%>:</td>
            <td class="content"><input name="province" type="text" value="<%=province%>" size="25"<% if intProvince = 1 then %> class="required"<% end if %>>*</td>
          </tr>
          <tr>
            <td width="160" class="content"><%=strCountry%>:</td>
            <td class="content"><input name="country" type="text" value="<%=country%>" size="25"<% if intCountry = 1 then %> class="required"<% end if %>>*</td>
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
              * </td>
          </tr>
          <tr>
            <td height="24" class="content"><%=strMobilePhone%>:</td>
            <td class="content"><input name="mobilePhone" type="text" id="telephone" value="<%=mobile%>" size="25"<% if intTelephone = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strFax%>:</td>
            <td class="content"><input name="fax" type="text" value="<%=fax%>" size="25"></td>
          </tr>
        </table></td>
  </tr>
</table>
<p><b><%=strYourPassword%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strPassword%>:</td>
            <td class="content"><input name="password1" type="password" id="country3" size="25"<% if intPassError = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strConfirmPassword%>:</td>
            <td class="content"><input name="password2" type="password" id="country4" size="25"<% if intPassError = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strSecurityCode%>:</td>
            <td class="content"><img src="security_image.asp?I=1&<%=hexValue(3)%>" height="20" width="15" /><img src="security_image.asp?I=2&<%=hexValue(3)%>" height="20" width="15" /><img src="security_image.asp?I=3&<%=hexValue(3)%>" height="20" width="15" /><img src="security_image.asp?I=4&<%=hexValue(3)%>" height="20" width="15" /><img src="security_image.asp?I=5&<%=hexValue(3)%>" height="20" width="15" /></td>
          </tr>
          <tr>
            <td width="160" class="content"><%=strConfirmSecurityCode%></td>
            <td class="content"><input name="SecurityCode" type="text" id="SecurityCode" size="25"<% if intSecurityCode = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
        </table></td>
  </tr>
</table>
<% if strShowNewsletter = 1 then %>
<p><b><%=strNewsLetter%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strSubscribeForNewsletter%>:</td>
            <td class="content"><input name="newsletter" type="checkbox" id="newsletter" value="1"> 
            </td>
          </tr>
		</table>
	  </td>
	</tr>
  </table>
<% end if %>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
      <td width="100%" class="table_content"><a href="javascript:document.frmRegister.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a> 
        <input name="Submit" type="hidden" id="Submit" value="ok">
      </td>
  </tr>
</table>
</form>