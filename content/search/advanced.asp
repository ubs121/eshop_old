<%
set rsCat = server.createobject("ADODB.recordset")
rsCat.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & session("language_id") & " ORDER BY menu_name ASC;"
rsCat.open strSQL, adoCon

set rsSubcat = server.createobject("ADODB.recordset")
rsSubcat.cursortype = 3

rsSubcat.open strSQL, adoCon
%>
<tr>
  <td class="searchoption">
  <%=strOnlySearchCategories%>:<br />
<%
set rsMain = server.createobject("ADODB.recordset")
rsMain.cursortype = 3

set rsSubmenu = server.createobject("ADODB.recordset")
rsSubmenu.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & session("language_id") & " ORDER BY menu_name ASC;"

rsMain.open strSQL, adoCon
rsSubmenu.open strSQL, adoCon

function getSubs(parent_id)
	temp = ""
	rsSubMenu.filter = "menu_parent_id = " & parent_id
	do while not rsSubMenu.eof
		if len(temp) > 0 then
			temp = temp & ";" & rsSubMenu("menu_id")
		else
			temp = rsSubMenu("menu_id")
		end if
		rsSubMenu.movenext
	loop
	getSubs = temp
end function

function getName(menu_id)
	rsSubMenu.filter = "menu_id = " & menu_id
	getName = rsSubMenu("menu_name")
end function

function countSubs(menu_id)
	countSubs = 0
	
	rsSubmenu.filter = "menu_id = " & menu_id
	if not rsSubmenu.eof then
		parent_id = rsSubmenu("menu_parent_id")
	else
		parent_id = 0
	end if
	
	do until cint(parent_id) = 0
		countSubs = countSubs + 1
		
		rsSubMenu.filter = "menu_id = " & parent_id
		if not rsSubmenu.eof then
			parent_id = rsSubMenu("menu_parent_ID")
		else
			parent_id = 0
		end if	
	loop
end function

function writeSubMenusDrop(menu_ID)
	allMenus   = getSubs(menu_id)
	tempBefore = ""
	tempAfter  = ""
	usedID     = 0
	
	if len(allMenus) > 0 then
		arrSubMenus = Split(allMenus, ";")
		
		x = 0
		
		for x = 0 to ubound(arrSubMenus)
			spaces = 3 * countSubs(arrSubMenus(x))
			y = 0
			tempSpaces = ""
			
			for y = 1 to spaces
				tempSpaces = tempSpaces & "&nbsp;"
			next
			tempSpaces = tempSpaces & "|--&nbsp;"
			
			tempBefore = tempBefore & "<option value=""" & arrSubMenus(x) & """"
			if cint(arrSubMenus(x)) = search_cat_id then
				tempBefore = tempBefore & " selected=""selected"""
			end if
			tempBefore = tempBefore & ">" & tempSpaces & getName(arrSubMenus(x)) & "</option>" & chr(10) & writeSubmenusDrop(arrSubmenus(x))
		next
	end if
	writeSubMenusDrop = tempBefore
end function

response.write("<select name=""scat_id"" id=""scat_id"">" & chr(10))
	response.write("<option value=""0""")
	if search_cat_id = 0 then
		response.write(" selected=""selected""")
	end if
	response.write(">" & strAllCategories & "</option>" & chr(10))
	
rsMain.filter = "menu_parent_id = 0"
do while not rsMain.eof
	response.write("<option class=""main"" value=""" & rsMain("menu_id") & """")
	if cint(rsMain("menu_id")) = search_cat_id then
		response.write(" selected=""selected""")
	end if
	response.write(">" & getName(rsMain("menu_ID")) & "</option>" & chr(10))
	
	response.write writeSubmenusDrop(rsMain("menu_ID"))
	
	rsMain.movenext
loop
response.write("</select>" & chr(10))

rsMain.close
set rsMain = nothing

rsSubmenu.close
set rsSubmenu = nothing
%>
  </td>
  <td class="searchoption">
  <%=strOnlySearchManufacturers%>:<br />
  <%
  set rsMan = server.createobject("ADODB.recordset")
  rsMan.cursortype = 3
  
  strSQL = "SELECT manufacturer_id, manufacturer_name FROM manufacturer ORDER BY manufacturer_name ASC;"
  rsMan.open strSQL, adoCon
  %>
  <select name="man_id">
    <option vale="0"><%=strAllManufacturers%></option>
	<%
	do while not rsMan.eof
		response.write("<option value=""" & rsMan("manufacturer_id") & """")
		if man_id = cint(rsMan("manufacturer_id")) then
			response.write(" selected=""selected""")
		end if
		response.write(">" & rsMan("manufacturer_name") & "</option>" & chr(13))
		rsMan.movenext
	loop
	%>
  </select>
  <%
  rsMan.close
  set rsMan = nothing
  %>
  </td>
<%
rsCat.close
set rsCat = nothing

rsSubcat.close
set rsSubcat = nothing
%>
</tr>