<%
if len(request.form()) > 0 then
	tStock = request.form("tStock")
	
	x = 0
	for x = 1 to tStock
		strSQL = "UPDATE products SET product_stock = " & request.form("stock_" & x) & " WHERE product_ID = " & request.form("id_" & x)
		adoCon.execute(strSQL)
	next
end if
%>
<div id="search">

</div>
<form name="frmUpdate" action="" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr style="color: #FFFFFF">
    <td width="30" align="center" bgcolor="#666666"><strong>ID</strong></td>
    <td bgcolor="#666666"><span class="style1"><strong>Productname</strong></span></td>
    <td width="50" align="center" bgcolor="#666666"><span class="style1"><strong>Stock</strong></span></td>
  </tr>
  <%
  pp = request.querystring("pp")
  
  if len(pp) > 0 and isnumeric(pp) then
  	pp = cint(pp)
  else
  	pp = 1
  end if
  
  set rsProducts = server.createobject("ADODB.recordset")
  rsProducts.cursortype = 3
  
  strSQL = "SELECT product_name, product_stock, product_ID FROM products ORDER BY product_stock ASC;"
  rsProducts.open strSQL, adoCon
  
  rsProducts.pagesize = 40
  pages = rsProducts.pagecount
  
  if not rsProducts.eof then rsProducts.absolutepage = pp
  
  bgcolor = "#EEEEEE"
  x = 0
  for x = 1 to 40
  	if rsProducts.eof then exit for
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
  %>
  <tr bgcolor="<%=bgcolor%>">
    <td align="center"><%=rsProducts("product_ID")%></td>
    <td>&nbsp;<%=rsProducts("product_name")%>
      <input name="id_<%=x%>" type="hidden" id="id_<%=x%>" value="<%=rsProducts("product_ID")%>"></td>
    <td align="center"><input name="stock_<%=x%>" type="text" id="stock_<%=x%>" value="<%=rsProducts("product_stock")%>" size="4"></td>
  </tr>
  <%
  	rsProducts.movenext
  next
  
  if x = 40 then
  	tStock = 40
  else
  	tStock = x - 1
  end if
  rsProducts.close
  set rsProducts = nothing
  %>
  <tr>
    <td colspan="3" align="right" bgcolor="#666666"><input name="tStock" type="hidden" id="tStock" value="<%=tStock%>">
      <input name="cmdUpdate" type="submit" id="cmdUpdate" value="Update stock"></td>
  </tr>
  <tr>
    <td colspan="3" align="left">
	  Pages:
	  <%
	  searchQuery = "?p=" & request.querystring("p")
	  for page = 1 to pages
	  	if page = pp then
			response.write "&nbsp;<b>" & pp & "</b>"
		else
			response.write "&nbsp;<a href=""" & searchQuery & "&amp;pp=" & page & """>" & page & "</a>"
		end if
	  next
	  %>
	</td>
  </tr>
</table>
</form>