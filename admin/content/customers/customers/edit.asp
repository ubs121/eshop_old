<%
cid = request.querystring("cid")

if len(cid) = 0 or not isnumeric(cid) then response.redirect("?p=" & request.querystring("p"))

'update the whole thingy
if len(request.form()) then
	firstname  = request.form("firstname")
	lastname   = request.form("lastname")
	gender     = request.form("gender")
	email      = request.form("email")
	password   = request.form("password")
	orig_email = request.form("orig_email")
	dob        = request.form("dob")
	telephone  = request.form("telephone")
	fax        = request.form("fax")	
	
	set rsUpdate = server.createobject("ADODB.recordset")
	strSQL = "SELECT * FROM users WHERE user_id = " & cid
	
	rsUpdate.open strSQL, adoCon, 2, 2
	
	if not rsUpdate.eof then
		rsUpdate("user_firstname")     = firstname
		rsUpdate("user_lastname")      = lastname
		rsUpdate("user_gender")        = gender
		rsUpdate("user_date_of_birth") = dob
		rsUpdate("user_telephone")     = telephone
		rsUpdate("user_fax")           = fax
		
		strMsg = "Customer updated with succes"
		'if email has changed or there is a new password update that as well
		if len(password) > 0 then
			error_id = 0
			if len(email) = 0 then email = orig_email
			
			'Check if email exists
			set rsEmail = server.createobject("ADODB.recordset")
			rsEmail.cursortype = 3
			
			strSQL = "SELECT user_id FROM users WHERE user_email = '" & email & "' AND user_id <> " & cid
			rsEmail.open strSQL, adoCon
			
			if not rsEmail.eof then
				error_id = 1
				strMsg = "This email is allready in use by another customer"
			end if
			
			rsEmail.close
			set rsEmail = nothing
			
			'Update salt & password
			if error_id = 0 then
				strSalt = getSalt(len(email))
				strSecret = hashEncode(password & strSalt)
				
				rsUpdate("user_salt")     = strSalt
				rsUpdate("user_password") = strSecret
				rsUpdate("user_email")    = email
			end if
		end if
		
		rsUpdate.update()
	end if
	
	rsUpdate.close
	set rsUpdate = nothing
end if

set rsCustomer = server.createobject("ADODB.recordset")
rsCustomer.cursortype = 3

strSQL = "SELECT user_firstname, user_lastname, user_date_of_birth, user_email, user_telephone, user_fax, user_gender FROM users WHERE user_id = " & cid
rsCustomer.open strSQL, adoCon

if not rsCustomer.eof then
	firstname = rsCustomer("user_firstname")
	lastname  = rsCustomer("user_lastname")
	dob       = rsCustomer("user_date_of_birth")
	email     = rsCustomer("user_email")
	telephone = rsCustomer("user_telephone")
	fax       = rsCustomer("user_fax")
	gender    = cint(rsCustomer("user_gender"))
end if

rsCustomer.close
set rsCustomer = nothing

set rsEntries = server.createobject("ADODB.recordset")
rsEntries.cursortype = 3

strSQL = "SELECT COUNT(user_address_id) AS TotalEntries FROM user_address GROUP BY user_id HAVING user_id = " & cid
rsEntries.open strSQL, adoCon

if not rsEntries.eof then
	totalEntries = clng(rsEntries("TotalEntries"))
else
	totalEntries = 0
end if

rsEntries.close
set rsEntries = nothing

if totalEntries = 1 then
	strTotalEntries = "(1 entry)"
else
	strTotalEntries = "(" & totalEntries & " entries)"
end if
%>
<form name="form1" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
  <% if len(strMsg) > 0 then %>
    <tr>
	  <td colspan="2"><b><font class="tiny"><%=strMsg%></font></b></td>
	</tr>
  <% end if %>
    <tr> 
      <td colspan="2">&nbsp;<b>Edit customer</b></td>
    </tr>
    <tr> 
      <td width="100">Firstname:</td>
      <td><input name="firstname" type="text" id="firstname" value="<%=firstname%>"></td>
    </tr>
    <tr> 
      <td>Lastname:</td>
      <td><input name="lastname" type="text" id="lastname" value="<%=lastname%>"></td>
    </tr>
    <tr> 
      <td>Gender:</td>
      <td><input type="radio" name="gender" value="1"<% if gender = 1 then %> checked="checked"<% end if %> />Male 
        <input type="radio" name="gender" value="2"<% if gender = 2 then %> checked="checked"<% end if %> >Female</td>
    </tr>
    <tr> 
      <td colspan="2" height="10"></td>
    </tr>
    <tr> 
      <td>Email:</td>
      <td><input name="email" type="text" id="email" value="<%=email%>">
        <input name="orig_email" type="hidden" id="orig_email" value="<%=email%>"><font class="tiny">(emailaddress will only be updated if you give a new password)</font></td>
    </tr>
    <tr> 
      <td>Password:</td>
      <td><input name="password" type="password" id="password">
        <font class="tiny">(leave blank if you do not want to change this)</font></td>
    </tr>
    <tr> 
      <td colspan="2" height="10"></td>
    </tr>
    <tr> 
      <td>Date of birth:</td>
      <td><input name="dob" type="text" id="dob" value="<%=dob%>"></td>
    </tr>
    <tr> 
      <td>Telephone:</td>
      <td><input name="telephone" type="text" id="telephone" value="<%=telephone%>"></td>
    </tr>
    <tr> 
      <td>Fax:</td>
      <td><input name="fax" type="text" id="fax" value="<%=fax%>"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input name="btnAddressbook" type="button" id="btnAddressbook" value="Addressbook <%=strTotalEntries%>" onclick="document.location='?p=<%=request.querystring("p")%>&action=addressbook&cid=<%=cid%>';" /></td>
    </tr>
    <tr align="center">
      <td colspan="2"><%=BuildSubmitter("submit","Update customer", request.querystring("p"))%> 
        <input type="button" name="Cancel" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
</form>
