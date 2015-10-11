<br />
<%
set rsNewProds = server.createobject("ADODB.recordset")
rsNewProds.cursortype = 3

strSQL = "SELECT TOP 8 product_id, product_cat_id, product_name, newPrice, product_image FROM products WHERE (year(product_date_added) = year(now()) AND month(product_date_added) = month(now())) ORDER BY product_date_added DESC;"
rsNewProds.open strSQL, adoCon
%>
<div class="box_large">
  <h2><%=Replace(strNewProductsFor, "[monthname]",strMonthName(month(now()))) %></h2>
  <table width="600" cellspacing="2" cellpadding="2" align="center">
  <% if rsNewProds.eof then %>
    <tr>
	  <td><p align="center"><b><%=Replace(strNoNewProductsFor, "[monthname]",strMonthName(month(now()))) %></b></p></td>
	</tr>
  <% else %>
  <%
  intTeller = 1
  for new_prod_loop = 1 to 8
  	if rsNewProds.eof then exit for
  	if intTeller = 1 then
		response.write("<tr>" & chr(13))	
	end if
	product_img   = rsNewProds("product_image")
	product_name  = rsNewProds("product_name")
	product_price = replace(rsNewProds("newPrice"), ".", strServerComma)
	product_id    = rsNewProds("product_id")
	product_cat_id = rsNewProds("product_cat_id")
	
	if instr(product_img, ";") > 0 then
		product_img = left(product_img, instr(product_img, ";") - 1)
	end if
	response.write("<td class=""new_products"">" & chr(13))
	
	response.write("<b>" & strCurrency & RoundNumber(product_price) & "</b><br />" & chr(13))
	response.write("<a href=""?mod=product&amp;cat_id=" & getLink(product_cat_id) & "&amp;product_id=" & product_id & """>")
	if len(product_img) > 0 then
		response.write("<img src=""images/products/" & product_img & """ alt=""" & product_name & """ width=""100"" /><br />" & chr(13))
	end if
	response.write(product_name & "</a><br />" & chr(13))
	response.write("</td>" & chr(13))
	
	if intTeller = 4 then
		intTeller = 1
		response.write("</tr>" & chr(13))
	else
		intTeller = intTeller + 1
	end if
  	rsNewProds.movenext
  next
  if intTeller > 1 then
	for x = 1 to (5 - intTeller)
		response.write("<td width=""150"">&nbsp;</td>")
	next
	response.write("</tr>")
  end if
  %> 
  <% end if %>
  </table>
</div>
<%
rsNewProds.close
set rsNewProds = nothing
%>