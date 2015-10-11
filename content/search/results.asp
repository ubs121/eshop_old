<%
search_type   = request.querystring("type")
query         = request.querystring("q")
pp            = request.querystring("pp")
search_cat_id = cint(request.querystring("scat_id"))
man_id        = request.querystring("man_id")

if len(man_id) > 0 and isnumeric(man_id) then
	man_id = cint(man_id)
else
	man_id = 0
end if

if len(search_cat_id) > 0 and isnumeric(search_cat_id) then
	search_cat_id = cint(search_cat_id)
else
	search_cat_id = 0
end if

if IsNumeric(pp) and len(pp) > 0 then
	pp = cint(pp)
else 
	pp = 1
end if

if len(search_type) = 0 then
	search_type = ""
end if
%>
<p class="pageheader"><% if search_type = "advanced" then %><%=strAdvancedSearch%>&nbsp;<% else %><%=strSimpleSearch%><% end if %></p>
<form action="" method="get" name="frmAdvancedSearch" id="frmAdvancedSearch">
  <table width="600" cellspacing="0" class="search" align="center">
    <tr> 
      <td width="300" class="searchoption"><div align="right"><%=strProductName%>:&nbsp;</div></td>
      <td width="300" class="searchoption">
	  	  <input name="mod" type="hidden" id="mod" value="search">
          <input name="type" type="hidden" id="type" value="<%=search_type%>">
          <input name="q" type="text" id="q" value="<%=query%>">
	  </td>
    </tr>
	<% if search_type = "advanced" then %>
	<!-- #include file="advanced.asp" -->
	<% end if %>
    <tr> 
      <td colspan="2" class="searchoption"><div align="center">
          <input type="submit" name="cmdSearch" value="<%=strSearch%>" /></div></td>
    </tr>
  </table>
  <table width="600px" align="center" style="border: 0px">
    <tr>
	  <td class="searchoption">
	    <div align="right">
		  <% if search_type = "advanced" then %>
		  <a href="?mod=search&amp;type=simple&amp;q=<%=query%>"><%=strSimpleSearch%></a>
		  <% else %>
	      <a href="?mod=search&amp;type=advanced&amp;q=<%=query%>"><%=strAdvancedSearch%></a>
		  <% end if %>
		</div>
	  </td>
	</tr>
  </table>
</form>
<br />
<% if len(query) > 0 or search_cat_id > 0 or man_id > 0 then %>
<%
set rsSubs = server.createobject("ADODB.recordset")
rsSubs.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id FROM menu WHERE menu_lang_ID = " & session("language_id")
rsSubs.open strSQL, adoCon

	function getSearchSubs(parent_id)
		temp = ""
		rsSubs.filter = "menu_parent_id = " & parent_id
		do while not rsSubs.eof
			if len(temp) > 0 then
				temp = temp & ";" & rsSubs("menu_id")
			else
				temp = rsSubs("menu_id")
			end if
			rsSubs.movenext
		loop
		getSearchSubs = temp
	end function
	
	function DisplaySubs(menu_ID)
		allMenus   = getSearchSubs(menu_id)
		tempBefore = ""
		
		if len(allMenus) > 0 then
			arrSubMenus = Split(allMenus, ";")
			
			x = 0
			
			for x = 0 to ubound(arrSubMenus)		
				tempBefore = tempBefore & "," & arrSubMenus(x)
				tempBefore = tempBefore & DisplaySubs(arrSubmenus(x))
			next
		end if
		DisplaySubs = tempBefore
	end function

if search_cat_id > 0 then		
	arrCats = search_cat_id & DisplaySubs(search_cat_id)
	if instr(arrCats, ",") > 0 then
		arrCats = Split(arrCats, ",")
		x = 0
		for x = 0 to ubound(arrCats)
			if x = 0 then
				strSearchCats = " (product_cat_ID = " & arrCats(x)
			else
				strSearchCats = strSearchCats & " OR product_cat_ID = " & arrCats(x)
			end if
		next
		strSearchCats = " AND " & strSearchCats & ")"
	else
		strSearchCats = " AND product_cat_ID = " & search_cat_id
	end if
end if

rsSubs.close
set rsSubs = nothing

if search_type = "advanced" then
	strCatQuery = ""
	strManQuery = ""
	
	if man_id > 0 then
		strManQuery = " AND product_manufacturer_id = " & man_id
	end if
	
	strSQL = "SELECT products.product_id, products.product_name, products.product_image, products.product_cat_id FROM products WHERE products.product_name LIKE '%" & query & "%'" & strSearchCats & strManQuery & " ORDER BY products.product_name ASC;"
	strSearchString = "?mod=search&amp;type=advanced&amp;q=" & query & "&amp;scat_id=" & search_cat_id & "&amp;man_id=" & man_id
else
	strSQL = "SELECT products.product_id, products.product_name, products.product_image, products.product_cat_id FROM products WHERE products.product_name LIKE '%" & query & "%' ORDER BY products.product_name ASC;"
	strSearchString = "?mod=search&amp;type=simple&amp;q=" & query
end if

set rsSearch = server.createobject("ADODB.recordset")
rsSearch.cursortype = 3

rsSearch.open strSQL, adoCon
rsSearch.pagesize = strProductsPerPage

intItemsFound = rsSearch.recordcount
intPages      = rsSearch.pagecount

if not rsSearch.eof then rsSearch.absolutepage = pp
%>
<div class="box_large">
  <h2>&nbsp;<%=strSearchResults%></h2>
  <p>
    &nbsp;<%=strItemsFound%>:&nbsp;<%=intItemsFound%><br />
	&nbsp;<%=strPages%>:
	<%
	if intPages = 0 then
		response.write("&nbsp;0")
	end if
	for x = 1 to intPages
		response.write("&nbsp;<a href=""" & strSearchString & "&amp;pp=" & x & """>" & x & "</a>")
	next
	%>
  </p>
</div>
<br>
<div class="box_large">
  <h2>&nbsp;<%=strItemsFound%></h2>
<%
for search_results_loop = 1 to strProductsPerPage
	if rsSearch.eof then exit for
	product_id = rsSearch("product_id")
	product_name = rsSearch("product_name")
	cat_id = getLink(rsSearch("product_cat_id"))
	product_image = rsSearch("product_image")
	
	if instr(product_image, ";") > 0 then
		'there is more then 1 image
		product_image = left(product_image, instr(product_image, ";") - 1)
	end if
%>
  <p class="searchresults">
  	<% if len(product_image) > 0 then %>
	<a href="javascript:;" onmouseover="doTooltip(event,'images/products/<%=product_image%>', '')" onmouseout="hideTip()" onclick="PopImage('product','<%=product_id%>','0');"><img src="images/foto.gif" width="20" alt="<%=product_name%>" align="absmiddle" /></a>
	<% else %>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<% end if %>
	&nbsp;<a href="?mod=product&amp;cat_id=<%=cat_id%>&amp;product_id=<%=product_id%>"><%=product_name%></a>
  </p>
<%
	rsSearch.movenext
next
%>
</div>
<%
rsSearch.close
set rsSearch = nothing
%>
<% end if %>