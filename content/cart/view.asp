<%
total_products = request.cookies("total_products" & session.SessionID)
if len(total_products) = 0 then
	total_products = 0
else
	total_products = cint(total_products)
end if

set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

strSQL = "SELECT product_id, product_name, newPrice, product_cat_id, weight FROM products"
rsProducts.open strSQL, adoCon
%>
<% if total_products > 0 then %>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="productListing">
  <tr> 
    <td width="50%" class="productListing-heading">&nbsp;<%=strProductName%></td>
    <td width="50" class="productListing-heading">&nbsp;<%=strNumberOf%></td>
    <td width="150" class="productListing-heading">&nbsp;<%=strPrice%></td>
    <td class="productListing-heading" width="80">&nbsp;<%=strRemove_add%></td>
    <td width="150" class="productListing-heading">&nbsp;<%=strTotalPrice%></td>
  </tr>
  <%
even = "even"
total_price = 0
total_weight = 0

for pCounter = 1 to total_products
	
	rsProducts.filter = "product_id = " & request.cookies("product" & pCounter & session.SessionID)("product_id")
	product_ordered = cint(request.cookies("product" & pCounter & session.SessionID)("product_ordered"))
	product_price = csng(Replace(rsProducts("newPrice"), ".", strServerComma))
	
	sub_price = product_price * product_ordered
	total_price = total_price + sub_price
	total_weight = total_weight + (product_ordered * csng(replace(rsProducts("weight"), ".", strServerComma)))
	if even = "even" then
		even = "odd"
	else
		even = "even"
	end if
	cat_id = rsProducts("product_cat_Id")
%>
  <tr class="productListing-<%=even%>"> 
    <td class="productListing-data">&nbsp;<a href="?mod=product&amp;product_id=<%=request.cookies("product" & pCounter & session.SessionID)("product_id")%>&amp;cat_id=<%=getLink(cat_id)%>"><%=rsProducts("product_name")%></a></td>
    <td class="productListing-data">&nbsp;<%=product_ordered%></td>
    <td class="productListing-data">&nbsp;<%=strCurrency%>&nbsp;<%=roundNumber(product_price)%></td>
    <td class="productListing-data">
	<div class="center">
	  <input name="products_to_alter_<%=x%>" type="text" id="products_to_alter_<%=pCounter%>" value="0" size="2" class="cart_input" />
      <a href="javascript:deleteFromCart('<%=pCounter%>','<%=pCounter%>');"><img src="images/icon_minus.gif" width="12" height="12" border="0" align="middle"></a><a href="javascript:addToCart('<%=request.cookies("product" & pCounter & session.SessionID)("product_id")%>','<%=pCounter%>');"><img src="images/icon_plus.gif" width="12" height="12" border="0" align="middle"></a> 
	</div>
    </td>
    <td class="productListing-data">&nbsp;<%=strCurrency%>&nbsp;<%=roundNumber(sub_price)%></td>
  </tr>
  <% next %>
  <tr> 
    <td colspan="4" class="productListing-heading">&nbsp;</td>
    <td class="productListing-heading">&nbsp;<%=strCurrency%>&nbsp;<%=roundNumber(total_price)%></td>
  </tr>
  <% if total_weight > 0 then %>
  <tr>
    <td colspan="4" class="productListing-heading">&nbsp;</td>
	<td class="productListing-heading">&nbsp;<%=total_weight%>&nbsp;<%=strWeightsign%></td>
  </tr>
  <% end if %>
</table>
<% else %>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="productListing">
  <tr> 
    <td align="center" class="productListing-data"><%=strCartIsEmpty%></td>
  </tr>
</table>
<% end if %>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	  <div align="right">
	  <a href="?mod=home"><img src="languages/<%=session("language")%>/images/button_continue_shopping.gif" alt="<%=strContinueShopping%>" width="122" height="22" border="0" /></a>
	  <a href="?mod=cart&amp;action=delete&amp;id=all"><img src="languages/<%=session("language")%>/images/button_empty_cart.gif" alt="<%=strEmptyCart%>" width="122" height="22" border="0"></a>
      <a href="?mod=checkout"><img src="languages/<%=session("language")%>/images/button_checkout.gif" alt="<%=strCheckout%>" width="122" height="22" border="0" /></a></div></td>
  </tr>
</table>
<%
session("totalWeight") = total_weight
rsProducts.close
set rsProducts = nothing
%>
