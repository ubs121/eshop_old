<%
function getStatus(div)
	temp_status = request.cookies(div & "_status")
	if temp_status = "collapsed" then
		getStatus = "boxcontent_collapsed"
	else
		getStatus = "boxcontent_open"
	end if
end function
%>
<% if strShowMcart = 1 then %>
<div class="box">
  <h2>
    <a href="javascript:switchBox('mcart');">
      <img class="boximg_collapse" src="images/cart_collapse.gif" alt="collapse/expand" title="collapse/expand" />
	</a>
	<%=strShoppingCart%>
  </h2>
  <div class="<%=getStatus("mcart")%>" id="mcart_content">
  <%
	total_products = request.cookies("total_products" & session.SessionID)
	if len(total_products) = 0 then
		total_products = 0
	else
		total_products = cint(total_products)
	end if
	
	total_price = 0
	if total_products = 0 then
		'there are no products
		response.write strCartIsEmpty
	else
		set rsProducts = server.createobject("ADODB.recordset")
		rsProducts.cursortype = 3
		
		strSQL = "SELECT product_ID, product_name, product_cat_ID, newPrice FROM products"
		rsProducts.open strSQL, adoCon
		
		x = 0
		for x = 1 to total_products
			product_ordered = cint(request.cookies("product" & x & session.SessionID)("product_ordered"))
			product_ID      = cint(request.cookies("product" & x & session.SessionID)("product_id"))
			
			rsProducts.filter = "product_id = " & product_ID
			if not rsProducts.eof then
				product_name = "<a href=""?mod=product&amp;product_ID=" & product_ID & "&amp;cat_ID=" & getLink(rsProducts("product_cat_ID")) & """>" & rsProducts("product_name") & "</a>"
				total_price = total_price + product_ordered * csng(Replace(rsProducts("newPrice"), ".", strServerComma))
			else
				product_name = "Unknown"
			end if
			
			response.write product_ordered & " x " & product_name & "<br />" & chr(10)
		next
		response.write "<p id=""mCartTotals"">" & strCurrency & roundNumber(total_price) & "</p>" & chr(10)
		
		rsProducts.close
		set rsProducts = nothing
	end if
  %>
  </div>
</div>
<br />
<% end if %>
<% if strShowLastproducts = 1 then %>
<div class="box">
  <%
	'Load 5 last products and randomly pick 1
	set rsProducts = server.createobject("ADODB.recordset")
	rsProducts.cursortype = 3
	
	strSQL = "SELECT TOP 5 product_ID, product_cat_ID, newPrice, product_name, product_image FROM products ORDER BY product_ID DESC;"
	rsProducts.open strSQL, adoCon
	
	if not rsProducts.eof then
		newProducts = 1
		strProducts = ""
		
		do while not rsProducts.eof
			if len(strProducts) = 0 then
				strProducts = rsProducts("product_ID")
			else
				strProducts = strProducts & ";" & rsProducts("product_ID")
			end if
			rsProducts.movenext
		loop
		rsProducts.movefirst
		
		arrProducts = split(strProducts, ";")
		if ubound(arrProducts) > 0 then
			'more then 1 product --> randomize
			randomize()
			intNumber = Int(ubound(arrProducts) * rnd())
			
			rsProducts.filter = "product_ID = " & arrProducts(intNumber)
		end if
		strImage = rsProducts("product_image")
		product_ID = rsProducts("product_ID")
		product_cat_id = rsProducts("product_cat_id")
		product_price = Replace(rsProducts("newPrice"), ".", strServerComma)
		product_name = rsProducts("product_name")
			
		if len(strImage) > 0 then
			if instr(strImage, ";") > 0 then
				strImage = left(strImage, instr(strImage, ";") - 1)	
			end if
			strImage = "products/" & strImage
		else
			strImage = "no_picture.jpg"
		end if
	end if
	
	rsProducts.close
	set rsProducts = nothing
  %>
  <h2>
    <a href="javascript:switchBox('lastProducts');">
      <img class="boximg_collapse" src="images/cart_collapse.gif" alt="collapse/expand" title="collapse/expand" />
	</a>
	<%=strLastProducts%>
  </h2>
  <div id="lastProducts_content" class="<%=getStatus("lastProducts")%>" align="center">
    <%
	if newProducts = 1 then
		response.write "<a href=""?mod=product&amp;product_ID=" & product_ID & "&amp;cat_ID=" & getLink(product_cat_id) & """>" & chr(10) & _
			"  <img src=""images/" & strImage & """ width=""80"" alt=""" & product_name & """ title=""" & product_name & """ /><br />" & chr(10) & _
			"  " & product_name & "<br />" & chr(10) & _
			"</a>" & chr(10) & _
			strCurrency & roundNumber(product_price)
	else
		response.write "There are no new products"
	end if
	%>
  </div>
</div>
<br />
<% end if %>
<% if strShowPopular = 1 then %>
<div class="box">
  <h2>
    <a href="javascript:switchBox('popular');">
      <img class="boximg_collapse" src="images/cart_collapse.gif" alt="collapse/expand" title="collapse/expand" />
	</a>
    <%=strMostPopular%>
  </h2>
  <div id="popular_content" class="<%=getStatus("popular")%>">
    <%
	set rsProducts = server.createobject("ADODB.recordset")
	rsProducts.cursortype = 3
	
	strSQL = "SELECT product_ID, product_name, product_cat_ID FROM products"
	rsProducts.open strSQL, adoCon
	
	set rsPop = server.createobject("ADODB.recordset")
	rsPop.cursortype = 3
	
	strSQL = "SELECT product_ID, SUM(products_ordered) AS totalOrdered FROM order_info INNER JOIN orders ON order_info.order_ID = orders.order_ID WHERE status = 4 AND product_type = 'product' GROUP BY product_ID ORDER BY SUM(products_ordered) DESC;"
	rsPop.open strSQL, adoCon
	
	popCounter = 0
	do while not rsPop.eof AND popCounter < 10
		rsProducts.filter = "product_ID = " & rsPop("product_ID")
		if not rsProducts.eof then
			popCounter = popCounter + 1
			response.write popCounter & ".&nbsp;<a href=""?mod=product&amp;product_ID=" & rsProducts("product_ID") & "&amp;cat_id=" & getLink(rsProducts("product_cat_ID")) & """>" & rsProducts("product_name") & "</a> (" & rsPop("totalOrdered") & ")<br />" & chr(10)
		end if
		rsPop.movenext
	loop
	
	rsPop.close
	set rsPop = nothing
	
	rsProducts.close
	set rsProducts = nothing
	%>
  </div>
</div>
<% end if %>