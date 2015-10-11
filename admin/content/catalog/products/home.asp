<%
cat_id = request.querystring("cat_id")
pp     = request.querystring("pp")

if len(pp) = 0 OR NOT isnumeric(pp) then
	pp = 1
else
	pp = cint(pp)
end if

if len(cat_id) = 0 OR not IsNumeric(cat_id) then
	cat_id = 0
else
	cat_id = cint(cat_id)
end if

if cat_id = 0 then
	strSQL = "SELECT product_name, product_id FROM products ORDER BY product_name ASC;"
else
	strSQL = "SELECT product_name, product_id FROM products WHERE product_cat_id = " & cat_id & " ORDER BY product_name ASC;"
end if

set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

rsProducts.open strSQL, adoCon

rsProducts.pagesize = 20
pages = rsProducts.pagecount

if not rsProducts.eof then rsProducts.absolutepage = pp
%>
<div align="right">
<form name="frmChangeCat" action="" method="get">
  <input type="hidden" name="p" value="<%=request.querystring("p")%>" />
  Category:
  <select name="cat_id" id="cat_id" onchange="document.frmChangeCat.submit();">
    <option value="0">All categories</option>
          <%
			set rsCat = server.createobject("ADODB.recordset")
			rsCat.cursortype = 3
			
			strSQL = "SELECT menu_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " AND menu_parent_id = 0"
			rsCat.open strSQL, adoCon
			
			set rsSubcat = server.createobject("ADODB.recordset")
			rsSubcat.cursortype = 3
			
			strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " AND menu_parent_id > 0"
			rsSubcat.open strSQL, adoCon
			
			do while not rsCat.eof
				response.write("<option value=""" & rsCat("menu_id") & """")
				if cint(rsCat("menu_id")) = cat_id then
					response.write(" selected=""selected""")
				end if
				response.write(">" & rsCat("menu_name") & "</option>" & chr(13))
				rsSubcat.filter = "menu_parent_id = " & rsCat("menu_id")
				do while not rsSubcat.eof
					response.write("<option value=""" & rsSubcat("menu_id") & """")
					if cint(rsSubcat("menu_id")) = cat_id then
						response.write(" selected=""selected""")
					end if
					response.write(">&nbsp;&nbsp;&raquo;" & rsSubcat("menu_name") & "</option>" & chr(13))
					rsSubcat.movenext
				loop
				rsCat.movenext
			loop
			
			rsSubcat.close
			set rsSubcat = nothing
			rsCat.close
			set rsCat = nothing
			%>
    </select>
</form>
</div>
<table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
<%
bgcolor = "#EEEEEE"
for x = product_looper to 20
	if rsProducts.eof then exit for
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
	response.write("<tr bgcolor=""" & bgcolor & """>" & chr(13))
		response.write("<td>" & rsProducts("product_name") & "</td>" & chr(13))
		if intModuleRights = 2 then
			response.write("<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;pid=" & rsProducts("product_id") & """>Edit</a></td>")
		else
			response.write("<td width=""60"">&nbsp;</td>")
		end if
		if intModuleRights = 2 then
			response.write("<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=delete&amp;pid=" & rsProducts("product_id") & """>Delete</a></td>")
		else
			response.write("<td width=""60"">&nbsp;</td>")
		end if
	response.write("</tr>")
	rsProducts.movenext
next
%>
</table>
<%
rsProducts.close
set rsProducts = nothing
%>
<p>
Pages:
	<%
	for page_looper = 1 to pages
		if page_looper = pp then
			response.write("&nbsp;<b>" & page_looper & "</b>")
		else
			response.write("&nbsp;<a href=""?p=" & request.querystring("p") & "&amp;cat_id=" & cat_id & "&amp;pp=" & page_looper & """>" & page_looper & "</a>")
		end if
	next
	%>
</p>
<p align="center">
  <a href="?p=<%=request.querystring("p")%>&amp;action=add">Add a product </a>
</p>