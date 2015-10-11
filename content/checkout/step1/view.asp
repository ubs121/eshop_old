<%
addressbook_id = session("checkout_addressbook_id")

set rsAddressbook = server.createobject("ADODB.recordset")
rsAddressbook.cursortype = 3

if len(addressbook_id) > 0 and Isnumeric(addressbook_id) then
	strSQL = "SELECT user_address_id, user_street, user_postcode, user_city, user_province, user_country, user_firstname, user_lastname FROM user_address WHERE user_address_id = " & addressbook_id & " AND user_id = " & session("customer_id")
else
	strSQL = "SELECT user_address_id, user_street, user_postcode, user_city, user_province, user_country, user_firstname, user_lastname FROM user_address WHERE user_default_address = -1 AND user_id = " & session("customer_id")
end if

rsAddressbook.open strSQL, adoCon

if not rsAddressbook.eof then
	session("checkout_addressbook_id") = rsAddressbook("user_address_id")
	user_firstname = rsAddressbook("user_firstname")
	user_lastname  = rsAddressbook("user_lastname")
	user_street    = rsAddressbook("user_street")
	user_postcode  = rsAddressbook("user_postcode")
	user_city      = rsAddressbook("user_city")
	user_province  = rsAddressbook("user_province")
	user_country   = rsAddressbook("user_country")
end if

rsAddressbook.close
set rsAddressbook = nothing
%>
<p><b><%=strShippingAddress%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="458" align="left" valign="top" class="content">
		    <%=strShippingAddressExplanation%><br /><br />
            <a href="?mod=checkout&amp;action=change"><img src="languages/<%=session("language")%>/images/button_change_address.gif" alt="<%=strChangeAddress%>" width="122" height="22" border="0" /></a> 
          </td>
          <td width="120" align="center" class="content"><b><%=strShippingAddress%></b><br /><img src="images/pinned.gif" alt="<%=strPrimaryAddress%>" width="40" height="20" align="middle" /></td>
          <td width="220" class="content">
		  <%=user_firstname%>&nbsp;<%=user_lastname%><br />
		  <%=user_street%><br />
		  <%=user_postcode%>&nbsp;<%=user_city%><br />
		  <%=user_province%>,&nbsp;<%=user_country%>
		  </td>
        </tr>
      </table></td>
    </tr>
  </table>
<form name="frmStep1" method="post" action="<%=strCurrFile%>?mod=checkout&amp;p=2">
<br />
<p><b><%=strPaymentMethod%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="458" align="left" valign="top" class="content">
		  <%
		  set rsPayment = server.createobject("ADODB.recordset")
		  rsPayment.cursortype = 3
		  
		  strSQL = "SELECT payment_id, payment_name FROM payment WHERE payment_lang_id = " & session("language_id")
		  rsPayment.open strSQL, adoCon
		  
		  do while not rsPayment.eof
		  	response.write("<input type=""radio"" name=""paymentMethod"" value=""" & rsPayment("payment_id") & """")
			if cint(session("checkout_payment_method")) = cint(rsPayment("payment_id")) then
				response.write(" checked=""checked""")
			end if
			response.write(" />&nbsp;" & rsPayment("payment_name") & "<br />" & chr(10))
		  	rsPayment.movenext
		  loop
		  %>
		  </td>
		</tr>
	  </table></td>
	</tr>
  </table>
<br />
<p><b><%=strOrderComments%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="center" valign="middle" class="content"><textarea name="order_comment" cols="100" id="order_comment"><%=session("checkout_order_comment")%></textarea></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="50%" class="content">&nbsp;</td>
    <td width="50%" align="right" class="content"><a href="javascript:document.frmStep1.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>