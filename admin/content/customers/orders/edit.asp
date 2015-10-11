<%
oid = request.querystring("oid")
go  = request.querystring("go")
msg = ""

if len(oid) > 0 and isnumeric(oid) then
	oid = clng(oid)
else
	response.redirect("?p=" & request.querystring("p"))
end if

select case go
	case "status":
		strSQL = "UPDATE orders SET status = " & request.querystring("id") & " WHERE order_id = " & oid
		adoCon.execute(strSQL)
		msg = "the orderstatus has been updated"
	case "payment":
		payment = request.querystring("value")
		if payment = "true" then
			payment = -1
		else
			payment = 0
		end if
		strSQL = "UPDATE orders SET Paid = " & payment & " WHERE order_id = " & oid
		adoCon.execute(strSQL)
end select

if len(request.form()) > 0 then
	total_products = cint(request.form("total_products"))
	x = 0
	
	for x = 1 to total_products
		order_info_id = clng(request.form("order_info_id_" & x))
		available     = request.form("chkProductAv_" & x)
		prevAvail     = request.form("prevAvail_" & x)
		nStock        = request.form("nStock_" & x)
		PID           = request.form("PID_" & x)
		
		if prevAvail <> available then
			if stock_autoUpdate = 1 then
				strSQL = "UPDATE products SET product_stock = " & nStock & " WHERE product_ID = " & PID
				adoCon.execute(strSQL)
			end if
			if len(available) > 0 then
				strSQL = "UPDATE order_info SET product_available = -1 WHERE order_info_id = " & order_info_id
			else
				strSQL = "UPDATE order_info SET product_available = 0 WHERE order_info_id = " & order_info_id
			end if
		end if
		
		adoCon.execute(strSQL)
	next
	msg = "The productstatus has been updated"
end if

set rsOrder = server.createobject("ADODB.recordset")
rsOrder.cursortype = 3

strSQL = "SELECT address_id, total_price, date_ordered, confirmed, date_confirmed, status, comment, user_firstname, user_lastname, users.user_id, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id WHERE order_id = " & oid
rsOrder.open strSQL, adoCon

if not rsOrder.eof then
	address_id = rsOrder("address_id")
	total_price = rsOrder("total_price")
	date_ordered = rsOrder("date_ordered")
	date_confirmed = rsOrder("date_confirmed")
	confirmed      = rsOrder("confirmed")
	order_status   = rsOrder("status")
	comment        = rsOrder("comment")
	username       = rsOrder("user_lastname") & " " & rsOrder("user_firstname")
	user_id        = rsOrder("user_id")
	paid           = rsOrder("Paid")
	
	if isnumeric(paid) then
		paid = cint(paid)
	else
		paid = 0
	end if
	
	if cint(confirmed) = -1 then
		confirmed = "<font color=""#006600"">yes</font>"
	else
		confirmed = "<font color=""#FF0000"">No</font>"
	end if
end if

rsOrder.close
set rsOrder = nothing

'Get the deliveryaddress
strSQL = "SELECT user_street, user_postcode, user_city, user_province, user_country, user_company_name, user_firstname, user_lastname FROM user_address WHERE "
if len(address_id) > 0 and isnumeric(address_id) then
	strSQL = strSQL & " user_address_id = " & address_id & " AND user_id = " & user_id
else
	strSQL = strSQL & " user_default_address = -1 AND user_id = " & user_id
end if

set rsAddress = server.createobject("ADODB.recordset")
rsAddress.cursortype = 3
rsAddress.open strSQL, adoCon

if not rsAddress.eof then
	delivery_name = rsAddress("user_lastname") & " " & rsAddress("user_firstname")
	delivery_address = rsAddress("user_street") & "<br />" & vbCrlf
	delivery_address = delivery_address & rsAddress("user_postcode") & " " & rsAddress("user_city") & "<br />" & vbCrlf
	delivery_address = delivery_address & rsAddress("user_province") & ", " & rsAddress("user_country") & "<br />" & vbCrlf
	delivery_company = rsAddress("user_company_name")
end if

rsAddress.close
set rsAddress = nothing
%>
<script>
<!--
function changeStatus(){
	status = document.getElementById("status").value;
	document.location = "?p=<%=request.querystring("p")%>&action=edit&oid=<%=oid%>&go=status&id=" + status;
}

function changePayment(){
	payment = document.getElementById("paid").checked;
	document.location = "?p=<%=request.querystring("p")%>&action=edit&oid=<%=oid%>&go=payment&value=" + payment;
}
-->
</script>
<form name="form1" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <% if len(msg) > 0 then %>
    <tr> 
      <td colspan="2"><font color="#FF0000"><b><%=msg%></b></font></td>
    </tr>
    <% end if %>
    <tr> 
      <td colspan="2">&nbsp;<b>Details for order #<%=oid%></b></td>
    </tr>
    <tr> 
      <td width="100">&nbsp;Customer:</td>
      <td>&nbsp;<%=username%></td>
    </tr>
    <tr> 
      <td>&nbsp;Total price:</td>
      <td>&nbsp;<%=shop_currency & total_price%></td>
    </tr>
    <tr> 
      <td>&nbsp;Date ordered:</td>
      <td>&nbsp;<%=date_ordered%></td>
    </tr>
    <tr> 
      <td>&nbsp;Confirmed:</td>
      <td>&nbsp;<b><%=confirmed%></b></td>
    </tr>
    <tr> 
      <td>&nbsp;Status:</td>
      <td> <%
	  set rsStatus = server.createobject("ADODB.recordset")
	  rsStatus.cursortype = 3
	  
	  strSQL = "SELECT order_status_id, order_status FROM order_status WHERE lang_id = " & default_lang_id & " ORDER BY order_status_id ASC;"
	  rsStatus.open strSQL, adoCon
	  %> <select name="status" id="status" onchange="javascript:changeStatus();">
          <%
		do while not rsStatus.eof
			response.write("<option value=""" & rsStatus("order_status_id") & """")
			if clng(rsStatus("order_status_id")) = clng(order_status) then
				response.write(" selected=""selected""")
			end if
			response.write(">" & rsStatus("order_status") & "</option>" & vbCrlf)
			rsStatus.movenext
		loop
		%>
        </select> <%
	  rsStatus.close
	  set rsStatus = nothing
	  %> </td>
    </tr>
    <tr> 
      <td>&nbsp;Paid:</td>
      <td><input name="paid" type="checkbox" id="paid" value="yes" onchange="javascript:changePayment();"<% if paid = -1 then %> checked="checked"<% end if %>></td>
    </tr>
    <tr>
      <td>&nbsp;Comments:</td>
      <td><%=comment%></td>
    </tr>
    <%
'Get the ordered products
set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

strSQL = "SELECT product_id FROM order_info WHERE order_id = " & oid & " AND product_type = 'delivery'"
rsProducts.open strSQL, adoCon

'get the deliverymethod
if not rsProducts.eof then
	delivery_id = rsProducts("product_id")
else
	delivery_id = 1
end if

rsProducts.close

strSQL = "SELECT product_cat_id, order_info_id, products_ordered,product_available, products.product_id, order_info.product_name, order_info.product_price, product_stock FROM order_info INNER JOIN products ON order_info.product_id = products.product_id WHERE order_id = " & oid & " AND product_type = 'product'"
rsProducts.open strSQL, adoCon

set rsDelivery = server.createobject("ADODB.recordset")
rsDelivery.cursortype = 3

strSQL = "SELECT delivery_name FROM delivery WHERE delivery_id = " & delivery_id & " AND lang_id = " & default_lang_id
rsDelivery.open strSQL, adoCon

if not rsDelivery.eof then
	delivery_method = rsDelivery("delivery_name")
end if

rsDelivery.close
set rsDelivery = nothing

set rsCats = server.createobject("ADODB.recordset")
rsCats.cursortype = 3

strSQL = "SELECT menu_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id
rsCats.open strSQL, adoCon
%>
    <tr> 
      <td colspan="2">&nbsp;<b>Delivery</b></td>
    </tr>
    <tr> 
      <td>&nbsp;Method:</td>
      <td><%=delivery_method%></td>
    </tr>
    <tr> 
      <td>&nbsp;Name</td>
      <td><%=delivery_name%></td>
    </tr>
    <% if len(delivery_company) > 0 then %>
    <tr> 
      <td>&nbsp;Company</td>
      <td><%=delivery_company%></td>
    </tr>
    <% end if %>
    <tr> 
      <td align="left" valign="top">&nbsp;Address:</td>
      <td><%=delivery_address%></td>
    </tr>
    <tr> 
      <td colspan="2">&nbsp;<b>Ordered products</b></td>
    </tr>
    <tr> 
      <td colspan="2"><table width="100%" cellspacing="4" cellpadding="4" style="border: solid 1px #000000;">
          <tr> 
		    <td width="100"><strong>Category</strong></td>
            <td><strong>Productname</strong></td>
			<td width="30" align="center"><strong>QTY</strong></td>
			<td width="30" align="center"><strong>Stock</strong></td>
            <td width="70" align="center"><strong>Price</strong></td>
            <td width="70" align="center"><strong>Available</strong></td>
          </tr>
          <%
			x = 0
			bgcolor = "#EEEEEE"
			do while not rsProducts.eof
				rsCats.filter = "menu_id = " & rsProducts("product_cat_id")
				if not rsCats.eof then
					cat_name = rsCats("menu_name")
				else
					cat_name = "Unknown"
				end if
				x = x + 1
				if bgcolor = "#EEEEEE" then
					bgcolor = "#FFFFFF"
				else
					bgcolor = "#EEEEEE"
				end if
				available = cint(rsProducts("product_available"))
				stock = rsProducts("product_stock")
				
				if len(stock) > 0 and isnumeric(stock) then
					stock = cint(stock)
				else
					stock = 0
				end if
				
				if available = -1 then
					nStock = stock + cint(rsProducts("products_ordered"))
				else
					nStock = stock - cint(rsProducts("products_ordered"))
				end if
				
				if nStock < 0 then nStock = 0
				
				if stock > cint(rsProducts("products_ordered")) then
					stock = "<font color=""#FFcc00"">" & stock & "</font>"
				else
					stock = "<font color=""#FF0000""><b>" & stock & "</b></font>"
				end if
			%>
          <tr bgcolor="<%=bgcolor%>"> 
		    <td>&nbsp;<%=cat_name%></td>
            <td>&nbsp;<%=rsProducts("product_name")%></td>
			<td align="center"><%=rsProducts("products_ordered")%></td>
			<td align="center"><%=stock%></td>
            <td align="center"><%=shop_currency & rsProducts("product_price")%></td>
            <td align="center"> <input name="chkProductAv_<%=x%>" type="checkbox" id="chkProductAv_<%=x%>" value="ok"<% if available = -1 then %> checked="checked"<% end if %>> 
              <input type="hidden" name="order_info_id_<%=x%>" value="<%=rsProducts("order_info_id")%>" />
			  <input type="hidden" name="prevAvail_<%=x%>" value="<% if available = -1 then%>ok<% else %>nok<% end if %>" />
			  <input type="hidden" name="nStock_<%=x%>" value="<%=nStock%>" />
			  <input type="hidden" name="PID_<%=x%>" value="<%=rsProducts("product_ID")%>" />
            </td>
          </tr>
          <%
		  		rsProducts.movenext
			loop
		  %>
        </table></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Update productstatus", request.querystring("p"))%> <input type="hidden" name="total_products" value="<%=x%>" /> 
        <input type="button" name="Cancel" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>&order_type=<%=order_status%>';"></td>
    </tr>
    <%
rsProducts.close
set rsProducts = nothing

rsCats.close
set rsCats = nothing
%>
  </table>
</form>