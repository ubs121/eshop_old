<%
if Isnumeric(request.form("delivery_method")) and len(request.form("delivery_method")) > 0 then
	session("checkout_order_comment") = request.form("order_comment")
	session("checkout_delivery_method") = request.form("delivery_method")
else
	if len(session("checkout_delivery_method")) = 0 then
		response.redirect("?mod=checkout&p=2")
	end if
end if

%>
<p><b><%=strConfirmOrder%></b></p>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="productListing">
  <tr> 
    <td class="productListing-heading">&nbsp;<%=strProductName%></td>
    <td width="100" align="center" class="productListing-heading"><%=strProductsOrdered%></td>
    <td width="100" align="center" class="productListing-heading"><%=strProductPrice%></td>
    <td width="100" align="center" class="productListing-heading"><%=strTotalPrice%></td>
  </tr>
  <%
even = "even"
total_products = request.cookies("total_products" & session.SessionID)

set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

strSQL = "SELECT product_name, product_id, newPrice, discount, discount_type FROM products"
rsProducts.open strSQL, adoCon

pLooper = 0
for pLooper = 1 to total_products
	if even = "even" then
		even = "odd"
	else
		even = "even"
	end if
	rsProducts.filter = "product_id = " & request.cookies("product" & pLooper & session.SessionID)("product_id")
	product_ordered = cint(request.cookies("product" & pLooper & session.SessionID)("product_ordered"))
	product_price = csng(Replace(rsProducts("newPrice"), ".", strServerComma))

	sub_price = product_price * product_ordered
	total_price = total_price + sub_price
%>
  <tr class="productListing-<%=even%>"> 
    <td class="productListing-data">&nbsp;<%=rsProducts("product_name")%></td>
    <td align="center" class="productListing-data"><%=product_ordered%></td>
    <td align="center" class="productListing-data"><%=strCurrency%><%=roundNumber(product_price)%></td>
    <td align="center" class="productListing-data"><%=strCurrency%><%=roundNumber(sub_price)%></td>
  </tr>
  <% 
next

rsProducts.close
set rsProducts = nothing

delivery_method = cint(session("checkout_delivery_method"))

set rsDelivery = server.createobject("ADODB.recordset")
rsDelivery.cursortype = 3

strSQL = "SELECT a, b, delivery_name FROM delivery WHERE lang_id = " & session("language_id") & " AND delivery_id = " & delivery_method
rsDelivery.open strSQL, adoCon

if not rsDelivery.eof then
	if rsDelivery("a") = "1" then
		delivery_price = csng(Replace(rsDelivery("b"), ".", strServerComma))
	else
		arrPrices = split(rsDelivery("b"), ";")
		arrConditions = split(rsDelivery("a"), ";")
		
		x = 0
		for x = 0 to ubound(arrConditions)
			if instr(arrConditions(x), ">=") > 0 then
				condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 2), ".", strServerComma))
				if csng(session("totalWeight")) >= condition then
					delivery_price = arrPrices(x)
				end if
			elseif instr(arrConditions(x), ">") > 0 then
				condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 1), ".", strServerComma))
				if csng(session("totalWeight")) > condition then
					delivery_price = arrPrices(x)
				end if
			else
				if csng(session("totalWeight")) < csng(right(arrConditions(x), len(arrConditions(x)) - 1)) then
					delivery_price = arrPrices(x)
				end if
			end if					
		next
	end if
	delivery_name = rsDelivery("delivery_name")
end if

rsDelivery.close
set rsDelivery = nothing

if delivery_price > 0 then
	if even = "even" then
		even = "odd"
	else
		even = "even"
	end if
	total_price = roundNumber(total_price + delivery_price)
%>
  <tr class="productListing-<%=even%>"> 
    <td class="productListing-data">&nbsp;<%=delivery_name%></td>
    <td align="center" class="productListing-data">1</td>
    <td align="center" class="productListing-data"><%=strCurrency%><%=delivery_price%></td>
    <td align="center" class="productListing-data"><%=strCurrency%><%=delivery_price%></td>
  </tr>
<% end if %>
  <tr> 
    <td colspan="2" class="productListing-heading">&nbsp;</td>
    <td align="right" class="productListing-heading"><%=strTotalPrice%>:&nbsp;</td>
    <td align="center" class="productListing-heading"><%=strCurrency%><%=roundNumber(total_price)%></td>
  </tr>
  <% if session("totalWeight") <> "0" then %>
  <tr>
    <td colspan="2" class="productListing-heading">&nbsp;</td>
    <td align="right" class="productListing-heading">&nbsp;</td>
    <td align="center" class="productListing-heading"><%=session("totalWeight") & strWeightsign%></td>
  </tr>
  <% end if %>
</table>
<br />
<p><b><%=strMenuConditions%></b></p>
<form name="frmStep3" method="post" action="<%=strCurrFile%>?mod=checkout&amp;p=4">
  <input type="hidden" name="total_price" value="<%=total_price%>" />
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="31" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td align="center" valign="middle" class="content"><iframe src="languages/<%=session("language")%>/conditions.asp" width="600" height="100"></iframe><br>
              <input type="submit" name="agree" value="<%=strIAgree%>">
              &nbsp;
              <input type="button" name="dontagree" value="<%=strIDontAgree%>" onclick="javascript:document.location='?mod=checkout&p=agreement';"> </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>