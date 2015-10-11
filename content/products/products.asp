<%
order = killChars(request.querystring("order"))
pp = killChars(request.querystring("pp"))

if len(pp) > 0 AND IsNumeric(pp) then
	pp = cint(pp)
else
	pp = 1
end if

set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

strSQL = "SELECT products.product_id, products.product_name, manufacturer.manufacturer_id, manufacturer.manufacturer_name, products.newPrice, products.product_image "
strSQL = strSQL & "FROM manufacturer INNER JOIN products ON manufacturer.manufacturer_id = products.product_manufacturer_id WHERE product_cat_id = " & arrCats(ubound(arrCats))
select case left(order,1)
	case "1":
		strOrderBy = " ORDER BY product_name"
		if right(order,1) = "d" then
			OppOrder1 = "a"
			Order1 = "-"
			strOrderBy = strOrderBy & " DESC;"
		else
			OppOrder1 = "d"
			Order1 = "+"
			strOrderBy = strOrderBy & " ASC;"
		end if
	case "2":
		strOrderBy = " ORDER BY manufacturer_name"
		if right(order,1) = "d" then
			OppOrder2 = "a"
			Order2 = "-"
			strOrderBy = strOrderBy & " DESC;"
		else
			OppOrder2 = "d"
			Order2 = "+"
			strOrderBy = strOrderBy & " ASC;"
		end if
	case "3":
		strOrderBy = " ORDER BY csng(newPrice)"
		if right(order,1) = "d" then
			OppOrder3 = "a"
			Order3 = "-"
			strOrderBy = strOrderBy & " DESC;"
		else
			OppOrder3 = "d"
			Order3 = "+"
			strOrderBy = strOrderBy & " ASC;"
		end if
	case else
		Order1 = ""
		Order2 = ""
		Order3 = ""
end select

strSQL = strSQL & strOrderBy

rsProducts.open strSQL, adoCon

TotalRecords = rsProducts.recordcount
rsProducts.pagesize = strProductsPerPage
pages = rsProducts.pagecount

FirstRecord = (pp * strProductsPerPage) - (strProductsPerPage - 1)

if not rsProducts.eof then
	rsProducts.absolutepage = pp
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="productListing">
  <tr> 
	<td width="5%" class="productListing-heading">&nbsp;</td>
    <td width="50%" class="productListing-heading">&nbsp;<a href="?mod=cat&parent_id=<%=parent_id%>&amp;cat_id=<%=cat_id%>&amp;order=1<%=OppOrder1%>&amp;pp=<%=pp%>" class="productListing-heading"><%=strProductName%><%=Order1%></a></td>
    <td width="25%" class="productListing-heading">&nbsp;<a href="?mod=cat&parent_id=<%=parent_id%>&amp;cat_id=<%=cat_id%>&amp;order=2<%=OppOrder2%>&amp;pp=<%=pp%>" class="productListing-heading"><%=strManufacturer%><%=Order2%></a></td>
    <td width="25%" class="productListing-heading">&nbsp;<a href="?mod=cat&parent_id=<%=parent_id%>&amp;cat_id=<%=cat_id%>&amp;order=3<%=OppOrder3%>&amp;pp=<%=pp%>" class="productListing-heading"><%=strPrice%><%=Order3%></a></td>
  </tr>
<%
even = "even"
LastRecord = FirstRecord - 1
for all_prod_looper = 1 to strProductsPerPage
	if rsProducts.eof then exit for
	LastRecord = LastRecord + 1
	if even = "even" then
		even = "odd"
	else
		even = "even"
	end if
	product_image = ""
	product_image = rsProducts("product_image")
	if instr(product_image, ";") > 0 then
		product_image = left(product_image, instr(product_image, ";") - 1)
	end if
	
	productPrice = csng(replace(rsProducts("newPrice"), ".", strServerComma))
%>
  <tr class="productListing-<%=even%>">
	<td class="productListing-data">&nbsp;
      <% if len(product_image) > 0 then %>
		<a href="javascript:;" onmouseover="doTooltip(event,'images/products/<%=product_image%>', '')" onmouseout="hideTip()" onclick="PopImage('product','<%=rsProducts("product_id")%>','0');"><img src="images/foto.gif" width="20" alt="<%=rsProducts("product_name")%>"></a>
	  <% end if %>
	</td>
    <td class="productListing-data">&nbsp;<a href="?mod=product&amp;cat_id=<%=cat_id%>&amp;product_id=<%=rsProducts("product_id")%>"><%=rsProducts("product_name")%></a></td>
    <td class="productListing-data">&nbsp;<%=rsProducts("manufacturer_name")%></td>
    <td class="productListing-data">&nbsp;<%=strCurrency%>&nbsp;<%=RoundNumber(productPrice)%></td>
  </tr>
<%
	rsProducts.movenext
next
%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="50%">
	  <%=strProduct%>&nbsp;<b><%=FirstRecord%></b>&nbsp;<%=strTo%>&nbsp;<b><%=LastRecord%></b>&nbsp;(&nbsp;<%=strOf%>&nbsp;<b><%=TotalRecords%></b>&nbsp;<%=strProducts%>&nbsp;)
	</td>
	<td width="50%" align="right">
	  <%=strPage%>:
		<%
		if pages = 1 then
			response.write("<b>1</b>&nbsp;")
		else
			if pp > 1 then
				response.write("<a href=""?mod=cat&amp;cat_id=" & getLink(arrCats(ubound(arrCats))) & "&amp;pp=" & pp - 1 & """>[&lt;&lt;" & strPrevious & "]</a>&nbsp;")
				if pp - 1 > 3 then
					response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=1"">1</a>&nbsp;")
					response.write("...&nbsp;")
					for x = pp - 2 to pp - 1
						response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & x & """>" & x & "</a>&nbsp;")
					next
				else
					for x = 1 to pp - 1
						response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & x & """>" & x & "</a>&nbsp;")
					next
				end if
			end if
			response.write("<b>" & pp & "</b>&nbsp;")
			if pages > pp then
				if pages - pp > 3 then
					for x = pp + 1 to pp + 2
						response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & x & """>" & x & "</a>&nbsp;")
					next
					response.write("...&nbsp;")
					response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & pages & """>" & pages & "</a>&nbsp;")
				else
					for x = pp + 1 to pages
						response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & x & """>" & x & "</a>&nbsp;")
					next
				end if
				response.write("<a href=""?mod=cat&amp;parent_id=" & parent_id & "&amp;cat_id=" & cat_id & "&amp;pp=" & (pp + 1) & """>[" & strNext & "&gt;&gt;]</a>&nbsp;")
			end if
		end if		
		%>
	</td>
  </tr>
</table>
<% else %>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="productListing">
  <tr>
    <td class="productListing-data"><%=strNoProductsInCategory%></td>
  </tr>
</table>
<% end if %>
<%
rsProducts.close
set rsProducts = nothing
%>