<%
if len(request.form("address_id")) > 0 then
	session("checkout_addressbook_id") = request.form("address_id")
	response.redirect("?mod=checkout")
end if
set  rsAddressbook = server.createobject("ADODB.recordset")
rsAddressbook.cursortype = 3

strSQL = "SELECT user_address_id, user_firstname, user_lastname, user_street, user_postcode, user_city, user_province, user_country, user_default_address FROM user_address WHERE user_id = " & session("customer_id")
rsAddressbook.open strSQL, adoCon
%>
<p><b><%=strAddressbookEntries%></b></p>
<form name="frmChangeAddress" action="<%=strCurrFile%>?mod=checkout&amp;action=change" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!-- default address -->
        <tr> 
          <td width="20" class="content">&nbsp;</td>
          <td class="content"> &nbsp;<b><%=rsAddressbook("user_lastname") & " " & rsAddressbook("user_firstname")%></b>&nbsp;<i>(<%=strPrimaryAddress%>)</i><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_street")%><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_postcode") & " " & rsAddressbook("user_city")%><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_province") & ", " & rsAddressbook("user_country")%> 
          </td>
          <td width="200" class="content"><input type="radio" name="address_id" value="<%=rsAddressbook("user_address_id")%>"<% if cint(session("checkout_addressbook_id")) = cint(rsAddressbook("user_address_id")) then %> checked="checked"<% end if %>></td>
        </tr>
        <!-- other addresses -->
        <%
	  rsAddressbook.filter = "user_default_address = 0"
	  do while not rsAddressbook.eof
	  %>
        <tr> 
          <td width="20" class="content">&nbsp;</td>
          <td class="content"> &nbsp;<b><%=rsAddressbook("user_lastname") & " " & rsAddressbook("user_firstname")%></b><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_street")%><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_postcode") & " " & rsAddressbook("user_city")%><br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_province") & ", " & rsAddressbook("user_country")%> 
          </td>
            <td width="200" class="content"><input type="radio" name="address_id" value="<%=rsAddressbook("user_address_id")%>"<% if cint(session("checkout_addressbook_id")) = cint(rsAddressbook("user_address_id")) then %> checked="checked"<% end if %>></td>
        </tr>
        <%
	  	rsAddressbook.movenext
	  loop
	  %>
      </table></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
      <td width="50%" class="content"><a href="?mod=checkout"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
	  <td width="50%" align="right" class="content"><a href="javascript:document.frmChangeAddress.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>
</form>
<%
rsAddressbook.close
set rsAddressbook = nothing
%>