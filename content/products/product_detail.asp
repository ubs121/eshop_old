<%
id = killChars(request.querystring("product_id"))

if len(id) > 0 AND IsNumeric(id) then
	id = cint(id)
else
	response.redirect("?mod=cat")
end if

set rsProduct = server.Createobject("ADODB.recordset")
rsProduct.cursortype = 3

strSQL = "SELECT product_name, product_manufacturer_id, product_price, product_image, product_link, weight, newPrice FROM products WHERE product_id = " & id
rsProduct.open strSQL, adoCon

product_name    = rsProduct("product_name")
price           = csng(replace(rsProduct("product_price"), ".", strServerComma))
newPrice        = csng(replace(rsProduct("newPrice"), ".", strServerComma))
product_url     = rsProduct("product_link")
product_image   = rsProduct("product_image")
manufacturer_id = rsProduct("product_manufacturer_id")
weight          = rsProduct("weight")
weight          = Replace(rsProduct("weight"), ".", strServerComma)

if isnumeric(weight) and len(weight) > 0 then
	weight = csng(weight)
else
	weight = 0
end if

rsProduct.Close
set rsProduct = nothing

'Calculate the discount for this product!
if price = newPrice then
	'no discount
	strProductPrice = "<p class=""normalprice"">" & strCurrency & roundNumber(price) & "</p>" & chr(10)
else
	'discount
	strProductPrice = "<p class=""oldprice"">" & strCurrency & roundNumber(price) & "</p>" & chr(10) & _
		"<p class=""newprice"">" & strCurrency & roundNumber(newPrice) & "</p>" & chr(10)
end if

'Get the images of the product
if len(product_image) > 0 then
	if instr(product_image, ";") > 0 then
		'There is more then 1 image
		arrImages = split(product_image, ";")
		
		imageLooper = 0
		for imageLooper = 0 to ubound(arrImages)
			strImgLink = strImgLink & "<a href=""javascript:PopImage('product','" & id & "','" & imageLooper & "')"" />" & chr(10) & _
				"  <img src=""images/products/" & arrImages(imageLooper) & """ width=""80"" alt=""" & product_name & """>" & chr(10) & _
				"</a>"
		next
	else
		'There is only 1 image
		strImgLink = "<a href=""javascript:PopImage('product','" & id & "','0');"">" & _
			"  <img src=""images/products/" & product_image & """ & width=""80"" alt=""" & product_name & """ />" & _
			"</a>"
	end if
	strImgLink = "<div align=""center"">" & strImgLink & "</div>"
end if

set rsDescription = server.createobject("ADODB.recordset")
rsDescription.cursortype = 3

strSQL = "SELECT product_description FROM product_description WHERE product_id = " & id & " AND product_lang_id = " & session("language_id")
rsDescription.open strSQL, adoCon

if not rsDescription.eof then
	product_description = rsDescription("product_description")
end if

rsDescription.close
set rsDescription = nothing
%>
<div class="productDetail">
<h1><%=product_name%></h1>
<p>
  <% if weight > 0 then %>
  <b>Weight:</b>&nbsp;<%=weight & " " & strWeightSign %><br />
  <% end if %>
  <% if showStock = 1 then %>
  <b>Stock:</b>&nbsp;<br />
  <% end if %>
</p>
<%
if len(product_description) > 0 then
	response.write(product_description)
else
	response.write("<br /><br /><p align=""center""><b>" & strItemNotAvailable & "</b></p>")
end if
%>
<br />&nbsp;
<table width="400" border="0" cellspacing="0" cellpadding="0" class="productListing">
        <tr> 
          <td colspan="3" class="productListing-heading" width="400">&nbsp;<%=strSpecifications%></td>
        </tr>
        <%
set rsSpecs = server.createobject("ADODB.recordset")
rsSpecs.cursortype = 3

strSQL = "SELECT product_info_name, product_info_description FROM product_info WHERE product_id = " & id & " AND product_lang_id = " & session("language_id")
rsSpecs.open strSQL, adoCon

even = "odd"
do while not rsSpecs.eof
%>
        <tr class="productListing-<%=even%>"> 
          <td width="160" class="productListing-data">&nbsp;<b><%=rsSpecs("product_info_name")%></b></td>
          <td width="10" align="center" class="productListing-data">:</td>
          <td class="productListing-data">&nbsp;<%=rsSpecs("product_info_description")%></td>
        </tr>
        <%
	if even = "even" then
		even = "odd"
	else
		even = "even"
	end if
	rsSpecs.movenext
loop

rsSpecs.close
set rsSpecs = nothing
%>
      </table>

<p> 
  <%
if len(product_url) > 0 then
	strUrlToProductInfo = Replace(strUrlToProductInfo,"[producturl]","[" & product_url & "]")
	response.write(strUrlToProductInfo)
end if
%>
<br />
</p>
<table width="400" border="0" cellspacing="0" cellpadding="0" class="productListing">
  <tr> 
    <td width="50%" height="20"><a href="javascript:history.back();"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strInCart%>" width="122" height="22" border="0" align="absmiddle" /></a></td>
	<td width="50%" height="20" align="right"> 
      <input name="products_to_alter" type="text" id="products_to_alter_1" value="1" size="2" class="cart_input" />
      <a href="javascript:addToCart('<%=product_id%>','1');"><img src="languages/<%=session("language")%>/images/button_in_cart.gif" alt="<%=strInCart%>" width="122" height="22" border="0" align="absmiddle" /></a></td>
  </tr>
</table>
</div>
<div class="product_price">
	<%=strProductPrice%>
</div>
<% if len(strImgLink) > 0 then %>
<div id="prodImages">
	<%=strImgLink%>
</div>
<% end if %>
<div class="clear"></div>