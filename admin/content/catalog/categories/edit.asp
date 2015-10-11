<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
if len(request.form("total_languages")) > 0 then
	total_languages = cint(request.form("total_languages"))
	menu_id         = cint(request.form("menu_id"))
	menu_parent_id  = request.form("menu_parent_id")
	menu_parent_id = cint(left(menu_parent_id, (instr(menu_parent_id, ",") - 1)))
	
	for x = 1 to total_languages
		language_id = cint(request.form("language_id_" & x))
		menu_new    = cint(request.form("menu_new_" & x))
		menu_name   = makeChars(request.form("menu_name_" & x))
		menu_image  = request.form("menu_img_" & x)
		
		if menu_new = 1 then
			strSQL = "INSERT INTO menu (menu_id, menu_lang_id, menu_name, menu_image, menu_parent_id) VALUES("
			strSQL = strSQL & menu_id & "," & language_id & ",'" & menu_name & "','" & menu_image & "'," & menu_parent_id & ");"
		else
			strSQL = "UPDATE menu SET menu_name = '" & menu_name & "', menu_image='" & menu_image & "', menu_parent_id = " & menu_parent_id & " WHERE menu_id = " & menu_id & " AND menu_lang_id = " & language_id
		end if
		adoCon.execute(strSQL)
	next
	'response.redirect("?p=" & request.querystring("p"))	
end if
menu_id = cint(request.querystring("menu_id"))
intTeller = 0

set rsLanguages = server.createobject("ADODB.recordset")
rsLanguages.cursortype = 3

strSQL = "SELECT language_id, language_name FROM lang WHERE language_show = -1"
rsLanguages.open strSQL, adoCon

set rsMenu = server.createobject("ADODB.recordset")
rsMenu.cursortype = 3

strSQL = "SELECT menu_lang_id, menu_name, menu_image, menu_parent_id FROM menu WHERE menu_id = " & menu_id
rsMenu.open strSQL, adoCon

menu_parent_id = cint(rsMenu("menu_parent_id"))
%>
<form name="form1" method="post" action="">
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Edit Category</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td align="left">Parent:</td>
      <td align="left">
<%
set rsMain = server.createobject("ADODB.recordset")
rsMain.cursortype = 3

set rsSubmenu = server.createobject("ADODB.recordset")
rsSubmenu.cursortype = 3

strSQL = "SELECT menu_id, menu_parent_id, menu_name FROM menu WHERE menu_lang_id = " & default_lang_id & " ORDER BY menu_name ASC;"

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
			if cint(arrSubMenus(x)) = menu_parent_id then
				tempBefore = tempBefore & " selected=""selected"""
			end if
			tempBefore = tempBefore & ">" & tempSpaces & getName(arrSubMenus(x)) & "</option>" & chr(10) & writeSubmenusDrop(arrSubmenus(x))
		next
	end if
	writeSubMenusDrop = tempBefore
end function

response.write("<select name=""menu_parent_id"" id=""menu_parent_id"">" & chr(10))
	response.write("<option value=""0""")
	if menu_parent_id = 0 then
		response.write(" selected=""selected""")
	end if
	response.write(">Main category</option>" & chr(10))
	
rsMain.filter = "menu_parent_id = 0"
do while not rsMain.eof
	response.write("<option value=""" & rsMain("menu_id") & """")
	if cint(rsMain("menu_id")) = menu_parent_id then
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
    </tr>
	<tr>
	  <td colspan="2" align="center">
	    <input name="btnUpload" type="button" id="btnUpload" value="Upload image" onclick="javascript:doUpload('category');">
      </td>
	</tr>
    <% 
do while not rsLanguages.eof
	intTeller = intTeller + 1
	rsMenu.filter = "menu_lang_id = " & rsLanguages("language_id")
	if not rsMenu.eof then
		menu_name = rsMenu("menu_name")
		menu_img  = rsMenu("menu_image")
		menu_new  = 0
		menu_parent_id = rsMenu("menu_parent_id")
	else
		menu_new       = 1
		menu_name      = ""
		menu_image     = ""
	end if
%>
    <tr> 
      <td width="120" align="left">&nbsp;<%=rsLanguages("language_name")%></td>
      <td align="left">
<input name="menu_name_<%=intTeller%>" type="text" id="menu_name_<%=intTeller%>" value="<%=menu_name%>">
        <input name="language_id_<%=intTeller%>" type="hidden" id="language_id_<%=intTeller%>" value="<%=rsLanguages("language_id")%>">
        <input name="menu_new_<%=intTeller%>" type="hidden" id="menu_new_<%=intTeller%>" value="<%=menu_new%>">
		<select name="menu_img_<%=intTeller%>" id="menu_img_<%=intTeller%>">
		  <option value=""></option>
	  	  <%
		  set objFSO = server.createobject("scripting.FileSystemObject")
		  set objFo  = objFSO.getfolder(server.mappath(strVirtualPath & "images/category/"))
	  
		  for each x in objFO.files
		    response.write("<option value=""" & x.name & """")
			if x.Name = menu_img then
				response.write(" selected=""selected""")
			end if
			response.write(">" & x.name & "</option>" & chr(13))
		  next
		  %>
        </select>
		</td>
    </tr>
    <%
	rsLanguages.movenext
loop
%>
	<tr>
	  <td colspan="2" align="center"><input name="menu_parent_id" type="hidden" id="menu_parent_id" value="<%=menu_parent_id%>">
        <input name="menu_id" type="hidden" id="menu_id" value="<%=menu_id%>">
        <input name="total_languages" type="hidden" id="total_languages" value="<%=intTeller%>">
        <%=BuildSubmitter("submit","Update", request.querystring("p"))%>
                <input type="button" name="Cancel" value="Cancel" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
	</tr>
  </table>
</form>
<%
rsLanguages.close
set rsLanguages = nothing

rsMenu.close
set rsMenu = nothing
%>
</td></tr>
</table>
</td></tr>
</table>
</center>
