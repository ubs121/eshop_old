<%
cid = request.querystring("cid")

if len(cid) = 0 or not isnumeric(cid) then response.redirect("?p=" & request.querystring("p"))

if len(request.form()) > 0 then
	entry_id = request.form("entry_id")
	action   = request.form("submit")
	
	'if action = update
	if action = "Update entry" then
		firstname = request.form("firstname")
		lastname  = request.form("lastname")
		street    = request.form("street")
		postcode  = request.form("postcode")
		city      = request.form("city")
		province  = request.form("province")
		country   = request.form("country")
		company   = request.form("company")
		primary   = cint(request.form("primary"))
		o_primary = cint(request.form("o_primary"))
		
		if primary = -1 AND o_primary = 0 then
			strSQL = "UPDATE user_address SET user_default_address = 0 WHERE user_id = " & cid
			adoCon.execute(strSQL)
		end if
		
		if o_primary = 1 AND primary = 0 then
			primary = -1
		end if
		
		set rsUpdate = server.createobject("ADODB.recordset")
		strSQL = "SELECT * FROM user_address WHERE user_address_id = " & entry_id
		
		rsUpdate.open strSQL, adoCon, 2, 2
		
		if not rsUpdate.eof then
			rsUpdate("user_firstname")       = firstname
			rsUpdate("user_lastname")        = lastname
			rsUpdate("user_street")          = street
			rsUpdate("user_postcode")        = postcode
			rsUpdate("user_city")            = city
			rsUpdate("user_province")        = province
			rsUpdate("user_country")         = country
			rsUpdate("user_default_address") = primary
			rsUpdate("user_company_name")    = company
			
			rsUpdate.update()
			
			strMsg = "Addressbook updated with succes"
		end if
		
		rsUpdate.close
		set rsUpdate = nothing
	'else
	elseif action = "Delete entry" then
		strSQL = "DELETE * FROM user_address WHERE user_id = " & cid & " AND user_address_id = " & entry_id & " AND user_default_address = 0"
		adoCon.execute(strSQL)
		
		strMsg = "Address deleted with succes"	
	end if
end if

set rsEntries = server.createobject("ADODB.recordset")
rsEntries.cursortype = 3

strSQL = "SELECT * FROM user_address WHERE user_id = " & cid
rsEntries.open strSQL, adoCon

totalEntries = rsEntries.recordcount
%>
<table width="500" align="center" cellpadding="2" cellspacing="2">
  <% if len(strMsg) > 0 then %>
  <tr> 
    <td colspan="2"><b><font class="tiny"><%=strMsg%></font></b></td>
  </tr>
  <% end if %>
  <tr> 
    <td colspan="2">&nbsp;<b>Update addressbook</b></td>
  </tr>
  <%
  	intCounter = 0
	
  	do while not rsEntries.eof
		intCounter = intCounter + 1
		entry_id   = rsEntries("user_address_id")
		firstname  = rsEntries("user_firstname")
		lastname   = rsEntries("user_lastname")
		street     = rsEntries("user_street")
		postcode   = rsEntries("user_postcode")
		city       = rsEntries("user_city")
		province   = rsEntries("user_province")
		country    = rsEntries("user_country")
		company    = rsEntries("user_company_name")
		primary    = cint(rsEntries("user_default_address"))
		
		if primary = -1 then
			o_primary = 1
		else
			o_primary = 0
		end if
  %>
  <tr> 
    <td width="20">&nbsp;</td>
    <td>
	<form name="frmAddressbook<%=intCounter%>" action="" method="post">
	    <table width="100%" style="border: solid 1px #000000;">
          <tr> 
            <td width="100">First name:</td>
            <td><input name="firstname" type="text" id="firstname" value="<%=firstname%>">
              <input name="entry_id" type="hidden" id="entry_id" value="<%=entry_id%>"></td>
          </tr>
          <tr> 
            <td>Last name:</td>
            <td><input name="lastname" type="text" id="lastname" value="<%=lastname%>"></td>
          </tr>
          <tr> 
            <td>Street:</td>
            <td><input name="street" type="text" id="street" value="<%=street%>"></td>
          </tr>
          <tr> 
            <td>Postcode:</td>
            <td><input name="postcode" type="text" id="postcode" value="<%=postcode%>" size="8"></td>
          </tr>
          <tr> 
            <td>City:</td>
            <td><input name="city" type="text" id="city" value="<%=city%>"></td>
          </tr>
          <tr> 
            <td>Province:</td>
            <td><input name="province" type="text" id="province" value="<%=province%>"></td>
          </tr>
          <tr> 
            <td>Country:</td>
            <td><input name="country" type="text" id="country" value="<%=country%>"></td>
          </tr>
          <tr>
            <td>Company:</td>
            <td><input name="company" type="text" id="company" value="<%=company%>"></td>
          </tr>
          <tr> 
            <td>Primary address:</td>
            <td><input name="primary" type="checkbox" id="primary" value="-1"<% if primary = -1 then %> checked="checked" disabled="disabled"<% end if %> />
              <input name="o_primary" type="hidden" id="o_primary" value="<%=o_primary%>"> </td>
          </tr>
          <tr align="right"> 
            <td colspan="2"><%=BuildSubmitter("submit","Update entry", request.querystring("p"))%><% if totalEntries > 1 then %>&nbsp;<%=BuildSubmitter("submit","Delete entry", request.querystring("p"))%><% end if %></td>
          </tr>
        </table>
	  </form>
	</td>
  </tr>
  <tr> 
    <td colspan="2" height="10"></td>
  </tr>
  <%
  		rsEntries.movenext
	loop
  %>
  <tr align="center"> 
    <td colspan="2"> 
      <input type="button" name="btnBack" value="Back to customer" onclick="document.location='?p=<%=request.querystring("p")%>&cid=<%=cid%>&action=edit';" />
    </td>
  </tr>
</table>
<%
rsEntries.close
set rsEntries = nothing
%>