<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Categories</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>

<table width="500" align="center" cellpadding="2" cellspacing="2">
<%
set rsSubMenu = server.createobject("ADODB.recordset")
rsSubMenu.cursortype = 3
	
strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " ORDER BY menu_name ASC;"
rsSubMenu.open strSQL, adoCon
	
set rsMenu = server.createobject("ADODB.recordset")
rsMenu.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " ORDER BY menu_name ASC;"
rsMenu.open strSQL, adoCon

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

function writeSubMenus(menu_ID)
	allMenus   = getSubs(menu_id)
	tempBefore = ""
	tempAfter  = ""
	usedID     = 0
	
	if len(allMenus) > 0 then
		arrSubMenus = Split(allMenus, ";")
		
		x = 0
		useBefore = 1
		
		for x = 0 to ubound(arrSubMenus)
			spaces = 3 * countSubs(arrSubMenus(x))
			y = 0
			tempSpaces = ""
			
			for y = 1 to spaces
				tempSpaces = tempSpaces & "&nbsp;"
			next
			tempSpaces = tempSpaces & "|--&nbsp;"
			
			tempBefore = tempBefore & "<tr bgcolor=""#EEEEEE""><td>" & tempSpaces & getName(arrSubMenus(x)) & "</td>" & chr(10) & _
				"<td align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;menu_id=" & arrSubmenus(x) & """>Edit</a></td>" & chr(10) & _
				"<td align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=delete&amp;menu_id=" & arrSubmenus(x) &""">Delete</a></td>" & chr(10) & _
				"</tr>" & chr(10) & _
				writeSubmenus(arrSubmenus(x))
		next
	end if
	writeSubMenus = tempBefore
end function

response.write("<table width=""100%"" cellspacing=""1"" cellpadding=""4"">" & chr(10))
rsMenu.filter = "menu_parent_id = 0"
do while not rsMenu.eof
	response.write "<tr bgcolor=""#CCCCCC"">" & chr(10) & _
		"<td>" & getName(rsMenu("menu_ID")) & "</td>" & chr(10) & _
		"<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;menu_id=" & rsMenu("menu_ID") & """>Edit</a></td>" & chr(10) & _
		"<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=delete&amp;menu_id=" & rsMenu("menu_ID") & """>Delete</a></td>" & chr(10) & _
		"</tr>" & chr(10)
	response.write writeSubMenus(rsMenu("menu_ID"))
	rsMenu.movenext
loop
response.write("</table>" & chr(10))
	
rsSubMenu.close
set rsSubMenu = nothing

rsMenu.Close
set rsMenu = nothing
%>
</table>
<% if intModuleRights = 2 then %>
	
<p align="center"><a href="?p=<%=request.querystring("p")%>&amp;action=add">Add 
  a category</a></p>
<% end if %>
</td></tr>
</table>
</center>
