
<%
oid = CInt(killChars(request.QueryString("oid")))

if len(oid) > 0 AND IsNumeric(oid) then
	oid = cint(oid)
end if

set rsOrder = server.createobject("ADODB.recordset")
rsOrder.cursortype = 3

strSQL = "SELECT address_ID, total_price, date_ordered, date_confirmed, confirmed, status, salt, comment, paid, payment FROM orders WHERE order_ID = " & oid & " AND user_ID = " & session("customer_ID")
rsOrder.open strSQL, adoCon

if rsOrder.eof then
	rsOrder.close
	set rsOrder = nothing
	
	response.redirect("?mod=myaccount&sub=orders_history")
end if

if cint(rsOrder("confirmed")) = -1 then
	date_confirmed = rsOrder("date_confirmed")
else
	date_confirmed = "<font color=""#FF0000""><b>" & strNotConfirmed & "</b></font>"
end if

set rsUser = server.createobject("ADODB.recordset")
rsUser.cursortype = 3

strSQL = "SELECT user_firstname, user_lastname, user_email, user_telephone, user_fax FROM users WHERE user_ID = " & session("customer_ID")
rsUser.open strSQL, adoCon

u_firstname = rsUser("user_firstname")
u_lastname  = rsUser("user_lastname")
u_telephone = rsUser("user_telephone")
u_fax       = rsUser("user_fax")
u_email     = rsUser("user_email")

rsUser.close
set rsUser = nothing

set rsAddress = server.createobject("ADODB.recordset")
rsAddress.cursortype = 3

strSQL = "SELECT user_address_ID, user_street, user_city, user_province, user_postcode, user_country, user_company_name, user_default_address, user_firstname, user_lastname FROM user_address WHERE user_ID = " & session("customer_ID")
rsAddress.open strSQL, adoCon

rsAddress.filter = "user_default_address = -1"

u_street = rsAddress("user_street")
u_postal = rsAddress("user_postcode")
u_province = rsAddress("user_province")
u_city     = rsAddress("user_city")
u_country  = rsAddress("user_country")
u_company  = rsAddress("user_company_name")

set rsStatus = server.createobject("ADODB.recordset")
rsStatus.cursortype = 3

strSQL = "SELECT order_status FROM order_status WHERE order_status_ID = " & rsOrder("status") & " AND lang_ID = " & session("language_ID")
rsStatus.open strSQL, adoCon

if not rsStatus.eof then
	o_status = rsStatus("order_status")
else
	o_status = strUnknown
end if

rsStatus.close
set rsStatus = nothing

set rsOrderinfo = server.createobject("ADODB.recordset")
rsOrderinfo.cursortype = 3

strSQL = "SELECT product_ID, products_ordered, product_type, product_available, product_name, product_price FROM order_info WHERE order_ID = " & oid
rsOrderinfo.open strSQL, adoCon

rsOrderinfo.filter = "product_type = 'delivery'"

if not rsOrderinfo.eof then
	set rsDelivery = server.createobject("ADODB.recordset")
	rsDelivery.cursortype = 3
	
	strSQL = "SELECT delivery_name, a, b FROM delivery WHERE delivery_ID = " & rsOrderinfo("product_ID") & " AND lang_ID = " & session("language_ID")
	rsDelivery.open strSQL, adoCon
	
	if not rsDelivery.eof then
		if rsDelivery("a") = "1" then
			delivery_price = roundNumber(Replace(rsDelivery("b"), ".", strServerComma))
		else
			arrPrices = split(rsDelivery("b"), ";")
			arrConditions = split(rsDelivery("a"), ";")
				
			x = 0
			for x = 0 to ubound(arrConditions)
					if instr(arrConditions(x), ">") > 0 then
					condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 1), ".", strServerComma))
					if csng(session("totalWeight")) > condition then
						delivery_price = csng(replace(arrPrices(x), ".", strServerComma))
					end if
				else
					if csng(session("totalWeight")) < csng(right(arrConditions(x), len(arrConditions(x)) - 1)) then
						delivery_price = csng(replace(arrPrices(x), ".", strServerComma))
					end if
				end if					
			next
		end if
	
		delivery_name  = rsDelivery("delivery_name")
	else
		delivery_name  = strUnknown
		delivery_price = 0
	end if
	
	rsDelivery.close
	set rsDelivery = nothing
else
	delivery_name  = strUnknown
	delivery_price = 0
end if

if isnumeric(rsOrder("payment")) then
	payment_id = cint(rsOrder("payment"))
else
	payment_id = 0
end if

set rsPayment = server.createobject("ADODB.recordset")
rsPayment.cursortype = 3

strSQL = "SELECT payment_name FROM payment WHERE payment_ID = " & payment_id & " AND payment_lang_ID = " & session("language_ID")
rsPayment.open strSQL, adoCon

if not rsPayment.eof then
	payment_name = rsPayment("payment_name")
else
	payment_name = strUnknown
end if

rsPayment.close
set rsPayment = nothing

if cint(rsOrder("paid")) = -1 then
	paid = "<font color=""#009900""><b>Yes</b></font>"
else
	paid = "<font color=""#FF0000""><b>No</b></font>"
end if
%>
<p><b><%=strOrderDetails%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2" class="content">
		  <%
		  response.write u_firstname & " " & u_lastname & "<br />" & chr(10)
		  if len(u_company) > 0 then response.write u_company & "<br />" & chr(10)
		  response.write u_street & "<br />" & chr(10)
		  response.write u_postal & " " & u_city & "<br />" & chr(10)
		  if len(u_province) > 0 AND len(u_country) > 0 then
		  	response.write u_province & ", " & u_country & "<br />" & chr(10)
		  elseif len(u_province) > 0 then
		  	response.write u_province & "<br />" & chr(10)
		  elseif len(u_country) > 0 then
		  	response.write u_country
		  end if
		  response.write "<br />"
		  
		  response.write("<b>" & strTel & ":</b> " & u_telephone & "<br />" & chr(10))
		  response.write("<b>" & strFax & ":</b> " & u_fax & "<br />" & chr(10))
		  response.write("<b>" & strEmail & ":</b> " & u_email & "<br />" & chr(10))
		  %>
		  <br />
		  </td>
	    </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strOrderNumber%>:</b></td>
          <td>&nbsp;<%=oid%></td>
        </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strDateOrdered%>:</b></td>
          <td>&nbsp;<%=rsOrder("date_ordered")%></td>
        </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strDateConfirmed%>:</b></td>
          <td>&nbsp;<%=date_confirmed%></td>
        </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strStatus%>:</b></td>
          <td>&nbsp;<%=o_status%></td>
        </tr>
		<% if len(rsOrder("comment")) > 0 then %>
		<tr>
		  <td valign="top" width="120">&nbsp;<b><%=strComments%>:</b></td>
		  <td>&nbsp;<%=rsOrder("comment")%></td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
</table>
<p><b><%=strDeliveryDetails%></b></p>
<%
rsAddress.filter = "user_address_ID = " & rsOrder("address_ID")
if not rsAddress.eof then
	u_firstname = rsAddress("user_firstname")
	u_lastname  = rsAddress("user_lastname")
	u_company   = rsAddress("user_company_name")
	u_street    = rsAddress("user_street")
	u_postal    = rsAddress("user_postcode")
	u_city      = rsAddress("user_city")
	u_province  = rsAddress("user_province")
	u_country   = rsAddress("user_country")
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2" class="content">
		  <%
		  response.write u_firstname & " " & u_lastname & "<br />" & chr(10)
		  if len(u_company) > 0 then response.write u_company & "<br />" & chr(10)
		  response.write u_street & "<br />" & chr(10)
		  response.write u_postal & " " & u_city & "<br />" & chr(10)
		  if len(u_province) > 0 AND len(u_country) > 0 then
		  	response.write u_province & ", " & u_country & "<br />" & chr(10)
		  elseif len(u_province) > 0 then
		  	response.write u_province & "<br />" & chr(10)
		  elseif len(u_country) > 0 then
		  	response.write u_country
		  end if
		  response.write "<br />"
		  
		  response.write("<b>" & strTel & ":</b> " & u_telephone & "<br />" & chr(10))
		  response.write("<b>" & strFax & ":</b> " & u_fax & "<br />" & chr(10))
		  %>
		  <br />
		  </td>
	    </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strDeliveryMethod%>:</b></td>
          <td>&nbsp;<%=delivery_name%></td>
        </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strPaymentMethod%>:</b></td>
          <td>&nbsp;<%=payment_name%></td>
        </tr>
        <tr>
          <td width="120">&nbsp;<b><%=strPaid%>:</b></td>
          <td bgcolor="#FFFFFF">&nbsp;<%=paid%></td>
        </tr>
	  </table>
	</td>
  </tr>
</table>
<p><b><%=strOrderDetails%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
		  <td class="content" width="20">&nbsp;</td>
		  <td class="content">&nbsp;<b><%=strProductName%></b></td>
		  <td width="100" align="right" class="content"><b><%=strQuantity%></b>&nbsp;</td>
		  <td width="100" align="right" class="content"><b><%=strProductPrice%></b>&nbsp;</td>
		  <td width="100" align="right" class="content"><b><%=strTotal%></b>&nbsp;</td>
		</tr>
		<%
		rsOrderinfo.filter = "product_type = 'product'"
		total_price = 0
		do while not rsOrderinfo.eof
			quant = rsOrderInfo("products_ordered")
			avail = cint(rsOrderinfo("product_available"))
			product_name  = rsOrderinfo("product_name")
			product_price = Replace(rsOrderinfo("product_price"), ".", strServerComma)
			sub_price     = quant * product_price
			
			total_price = total_price + sub_price
			
			if avail = -1 then
				avail = "<img src=""images/product_available.gif"" alt=""Available"" title=""Available"" width=""12"" height=""12"" align=""absmiddle"" />"
			else
				avail = "<img src=""images/product_not_available.gif"" alt=""Not available"" title=""Not available"" width=""12"" height=""12"" align=""absmiddle"" />"
			end if
			
			response.write "<tr>" & chr(10) & _
				"  <td class=""content"" align=""center"">" & avail & "</td>" & chr(10) & _
				"  <td class=""content"">" & product_name & "</td>" & chr(10) & _
				"  <td class=""content"" align=""right"">" & quant & "</td>" & chr(10) & _
				"  <td class=""content"" align=""right""><b>" & strCurrency & "</b>" & RoundNumber(product_price) & "</td>" & chr(10) & _
				"  <td class=""content"" align=""right""><b>" & RoundNumber(sub_price) & "</b></td>" & chr(10) & _
				"</tr>" & chr(10)
			rsOrderinfo.movenext
		loop
		if delivery_price > 0 then
			total_price = total_price + delivery_price
			response.write "<tr>" & chr(10) & _
				"  <td class=""content"">&nbsp;</td>" & chr(10) & _
				"  <td class=""content"">" & delivery_name & "</td>" & chr(10) & _
				"  <td class=""content"" align=""right"">1</td>" & chr(10) & _
				"  <td class=""content"" align=""right""><b>" & strCurrency & "</b>" & delivery_price & "</td>" & chr(10) & _
				"  <td class=""content"" align=""right""><b>" & delivery_price & "</b></td>" & chr(10) & _
				"</tr>" & chr(10)
		end if
		%>
		<tr>
		  <td colspan="4" align="right" class="content"><b><%=strTotal%>:</b></td>
		  <td class="content" align="right"><b><%=strCurrency & RoundNumber(rsOrder("total_price"))%></b></td>
		</tr>
	  </table>
	</td>
  </tr>
</table>
<p>
  <img src="images/product_available.gif" width="12" height="12" align="absmiddle" />&nbsp;<%=strProductAvailableLable%><br />
  <img src="images/product_not_available.gif" width="12" height="12" align="absmiddle" />&nbsp;<%=strProductNotAvailableLable%>
</p>
<%
rsOrder.close
set rsOrder = nothing

rsOrderinfo.close
set rsOrderinfo = nothing
%>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><a href="?mod=myaccount&amp;sub=orders_history"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strBack%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>