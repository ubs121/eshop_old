<%
intStreet    = 0
intFirstname = 0
intLastname  = 0
intPostcode  = 0
intCity      = 0
intProvince  = 0
intCountry   = 0
intError     = 0

private function CheckValues(stringval)
	CheckValues = 0
	if len(stringval) = 0 then
		intError = 1
		CheckValues = 1
	end if
end function

if request.form("submitter") = "ok" then
	firstname = request.form("firstname")
	lastname  = request.form("lastname")
	street    = request.form("street")
	postcode  = request.form("postcode")
	city      = request.form("city")
	province  = request.form("province")
	country   = Rtrim(Ltrim(Replace(request.form("country"),",","")))
	primary   = request.form("primary")
	company   = request.form("company")
	vat       = request.form("vat")
	
	if len(primary) > 0 then
		primary = cint(primary)
	else
		primary = 0
	end if
	
	intFirstname = CheckValues(firstname)
	intLastname  = CheckValues(lastname)
	intStreet    = CheckValues(street)
	intPostcode  = CheckValues(postcode)
	intCity      = CheckValues(city)
	intProvince  = CheckValues(province)
	intCountry   = CheckValues(country)
	
	if intError = 0 then
		if primary = -1 then
			strSQL = "UPDATE user_address SET user_default_address = 0 WHERE user_default_address = -1 AND user_id = " & session("customer_id")
			adoCon.execute(strSQL)
		end if	
		strSQL = "INSERT INTO user_address (user_id, user_firstname, user_lastname, user_street, user_postcode, user_city, user_province, user_country, user_company_name, user_default_address, user_vat)"
		strSQL = strSQL & " VALUES(" & session("customer_id") & ",'" & firstname & "','" & lastname & "','" & street & "','"
		strSQL = strSQL & postcode & "','" & city & "','" & province & "','" & country & "','" & company & "'," & primary & ",'" & vat & "');"
		adocon.execute(strSQL)
		
		response.redirect("?mod=myaccount&sub=addressbook")
	end if
end if
%>
<p align="right">* <%=strRequiredFields%></p>
<form name="frmAddAddress" method="post" action="">
  <p><b><%=strNewAddress%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strFirstName%>:</td>
            <td class="content"> <input name="firstname" type="text" id="firstname2" value="<%=firstname%>" size="25"<% if intFirstname = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strLastName%>:</td>
            <td class="content"> <input name="lastname" type="text" id="lastname2" value="<%=lastname%>" size="25"<% if intLastname = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strStreet%>:</td>
            <td class="content"><input name="street" type="text" value="<%=street%>" size="25"<% if intStreet = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strPostCode%>:</td>
            <td class="content"><input name="postcode" type="text" value="<%=postcode%>" size="25"<% if intPostcode = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strCity%>:</td>
            <td class="content"><input name="city" type="text" value="<%=city%>" size="25"<% if intCity = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strProvince%>:</td>
            <td class="content"><input name="province" type="text" value="<%=province%>" size="25"<% if intProvince = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strCountry%>:</td>
            <td class="content"><input name="country" type="text" value="<%=country%>" size="25"<% if intCountry = 1 then %> class="required"<% end if %>>
              *</td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strCompanyName%>:</td>
            <td class="content"><input name="country" type="text" value="<%=company%>" size="25"> 
            </td>
          </tr>
          <tr>
            <td class="content"><%=strVat%>:</td>
            <td class="content"><input name="vat" type="text" id="vat" value="<%=vat%>" size="25"></td>
          </tr>
          <tr> 
            <td width="160" class="content"><%=strPrimaryAddress%>:</td>
            <td class="content"><input name="primary" type="checkbox" id="primary" value="-1"> 
              <input name="submitter" type="hidden" id="submitter" value="ok"></td>
          </tr>
        </table></td>
    </tr>
  </table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      <td><a href="?mod=myaccount&amp;sub=addressbook"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a> 
      </td>
      <td align="right"><a href="javascript:document.frmAddAddress.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>
</form>