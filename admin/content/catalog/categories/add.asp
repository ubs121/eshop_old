<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
if len(request.form("parent_id")) > 0 then
	total_languages = cint(request.form("total_languages"))
	parent_id   = request.form("parent_id")
	menu_id     = request.form("menu_id")
	for x = 1 to total_languages
		language_id = request.form("language_id_" & x)
		menu_name   = makeChars(Replace(request.form("menu_name_" & x), "'", "''"))
		menu_image  = request.form("menu_image_" & x)
		
		strSQL = "INSERT INTO menu (menu_name, menu_lang_id, menu_parent_id, menu_id, menu_image) VALUES('"
		strSQL = strSQL & menu_name & "'," & language_id & "," & parent_id & "," & menu_id & ",'" & menu_image & "');"
		adoCon.execute(strSQL)
	next
	response.redirect("?p=" & request.querystring("p"))
end if
intTeller = 0

set rsMenu = server.createobject("ADODB.recordset")
strSQL = "SELECT TOP 1 menu_id FROM menu ORDER BY menu_id DESC;"
rsMenu.open strSQL, adoCon
	
	if not rsMenu.eof then
		menu_id = cint(rsMenu("menu_id")) + 1
	else
		menu_id = 1
	end if

rsMenu.close
set rsMenu = nothing

set rsLanguages = server.createobject("ADODB.recordset")
rsLanguages.cursortype = 3

strSQL = "SELECT language_id, language_name FROM lang WHERE language_show = -1"
rsLanguages.open strSQL, adoCon
%>
<form name="form1" method="post" action="">
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Add Category</font></b></td>
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

response.write("<select name=""parent_id"" id=""parent_id"">" & chr(10))
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
      <td colspan="2" align="left" height="10"></td>
    </tr>
	<tr>
	  <td colspan="2" align="center">
	    <input name="btnUpload" type="button" id="btnUpload" value="Upload image" onclick="javascript:doUpload('category');">
      </td>
	</tr>
<%
do while not rsLanguages.eof
	intTeller = intTeller + 1
%>
    <tr> 
      <td width="120" align="left">&nbsp;<%=rsLanguages("language_name")%></td>
      <td align="left"> <input name="menu_name_<%=intTeller%>" type="text" id="menu_name_<%=intTeller%>"> 
        <input name="language_id_<%=intTeller%>" type="hidden" id="language_id_<%=intTeller%>" value="<%=rsLanguages("language_id")%>">
		<select name="menu_image_<%=intTeller%>" id="menu_image_<%=intTeller%>">
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
      <td colspan="2" align="center"><input name="menu_id" type="hidden" id="menu_id" value="<%=menu_id%>"> 
        <input name="total_languages" type="hidden" id="total_languages" value="<%=intTeller%>">
       	<%=BuildSubmitter("submit","Add category", request.querystring("p"))%> <input type="button" name="Cancel" value="Cancel" onclick="document.location='?p=<%=request.querystring("p")%>';">	
                </td>
  </tr>
  </table>
</td></tr>
</table>
</center>
</form>
