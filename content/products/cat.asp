<%
set rsCat = server.createobject("ADODB.recordset")
rsCat.cursortype = 3

strSQL = "SELECT menu_ID FROM menu WHERE menu_parent_ID = " & arrCats(ubound(arrCats))
rsCat.open strSQL, adoCon

if not rsCat.eof then
	showCat = true
end if

rsCat.close
set rsCat = nothing

set rsProd = server.createobject("ADODB.recordset")
rsProd.Cursortype = 3

strSQL = "SELECT product_ID FROM products WHERE product_cat_ID = " & arrCats(ubound(arrCats))
rsProd.open strSQL, adoCon

if not rsProd.eof then
	showProducts = true
else
	if showCat = false then
		showProducts = true
	end if
end if

rsProd.close
set rsProd = nothing

'Display subcategories
	set rsCat = server.createobject("ADODB.recordset")
	rsCat.cursortype = 3
	
	strSQL = "SELECT menu_id, menu_name, menu_image, menu_parent_id FROM menu WHERE menu_lang_id = " & session("language_id")
	rsCat.open strSQL, adoCon
	
	rsCat.filter = "menu_id = " & arrCats(ubound(arrCats))
	
	menu_name = rsCat("menu_name")
	
	rsCat.filter = "menu_parent_id = " & arrCats(ubound(arrCats))
%>
<p class="pageheader"><%=menu_name%></p>
<% if not rsCat.eof then %>
<table width="600" cellspacing="0" cellpadding="4" style="border: 0px;" align="center">
<%
set rsSubMenu = server.createobject("ADODB.recordset")
rsSubMenu.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id FROM menu WHERe menu_lang_id = " & session("language_id")
rsSubmenu.open strSQL, adoCon

intTeller = 1
do while not rsCat.eof
	if intTeller = 1 then
		response.write("<tr align=""center"">")
	end if
	response.write("<td width=""25%"" class=""catdisplay"">")
		response.write("<a href=""?mod=cat&amp;cat_id=" & getLink(rsCat("menu_id")) & """>")
		if len(rsCat("menu_image")) > 0 then
			response.write("<img src=""images/category/" & rsCat("menu_image") & """ width=""100"" alt=""" & rsCat("menu_name") & """ />")
		else
		
		end if
		response.write("<br />" & rsCat("menu_name"))
		response.write("</a>")
	response.write("</td>")
	if intTeller = 4 then
		intTeller = 1
		response.write("</tr>")
	else
		intTeller = intTeller + 1
	end if
	rsCat.movenext
loop
if intTeller > 1 then
	for x = intTeller to 4
		response.write("<td>&nbsp;</td>")
	next
	response.write("</tr>")
end if

rsSubmenu.close
set rsSubmenu = nothing
%>
</table>
<% end if %>
<% if showProducts then %>
<!-- #include file="products.asp" -->
<% end if %>
<%
	rsCat.close
	set rsCat = nothing
%>
<!-- #include file="new.asp" -->