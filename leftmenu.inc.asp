<div class="box">
  <h2><%=strMenuCategories%></h2>
    <%
	set rsMenu = server.createobject("ADODB.recordset")
	rsMenu.cursortype = 3
	
	set rsSubMenu = server.createobject("ADODB.recordset")
	rsSubmenu.cursortype = 3
	
	strSQL = "SELECT menu_id, menu_name, menu_parent_id FROM menu WHERE menu_lang_id = " & session("language_id") & " ORDER BY menu_name ASC;"
	rsMenu.open strSQL, adoCon
	rsSubmenu.open strSQL, adoCon
	
	openMenus = request.querystring("cat_id")
	if len(openMenus) > 0 then
		arrMenus = Split(openMenus, ",")
	end if
	
	openMenus = "," & openMenus & ","
	strBeforeMenu = "<ul class=""menu"">"
	strAfterMenu  = "</ul>"
	
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
	
	function writeSubMenus(menu_id)
		curLink    = getLink(menu_id)
		allMenus   = getSubs(menu_id)
		tempBefore = "<ul class=""menu"">"
		tempAfter  = ""
		usedID     = 0
			
		if len(allMenus) > 0 then
			arrSubMenus = Split(allMenus, ";")
			
			x = 0
			useBefore = 1
			
			for x = 0 to ubound(arrSubMenus)
				if instr(openMenus, "," & arrSubMenus(x) & ",") > 0 then
					tempBefore = tempBefore & "<li class=""m_selected""><a href=""?mod=cat&amp;cat_id=" & curLink & "," & arrSubmenus(x) & """>" & getName(arrSubMenus(x)) & "</a></li>"
					useBefore  = 0
					usedID     = arrSubmenus(x)
				elseif useBefore = 1 then
					tempBefore = tempBefore & "<li><a href=""?mod=cat&amp;cat_id=" & curLink & "," & arrSubmenus(x) & """>" & getName(arrSubMenus(x)) & "</a></li>"
				else
					tempAfter = tempAfter & "<li><a href=""?mod=cat&amp;cat_id=" & curLink & "," & arrSubmenus(x) & """>" & getName(arrSubMenus(x)) & "</a></li>"
				end if
			next
			tempAfter = tempAfter & "</ul>"
			
			if usedID = 0 then
				writeSubMenus = tempBefore & tempAfter
			else
				writeSubMenus = tempBefore & writeSubMenus(usedID) & tempAfter
			end if
		end if
	end function
	
	if isArray(arrMenus) then
		rsMenu.filter = "menu_parent_id = 0"
		response.write("<ul class=""menu"">" & chr(10))
		do while not rsMenu.eof
			if cint(arrMenus(0)) = cint(rsMenu("menu_id")) then
				response.write("<li class=""m_selected""><a href=""?mod=cat&amp;cat_id=" & rsMenu("menu_id") & """>" & rsMenu("menu_name") & "</a></li>" & chr(10))	
				response.write(writeSubMenus(rsMenu("menu_id")))
			else
				response.write("<li><a href=""?mod=cat&amp;cat_id=" & rsMenu("menu_id") & """>" & rsMenu("menu_name") & "</a></li>" & chr(10))	
			end if
			rsMenu.movenext
		loop
		response.write("</ul>" & chr(10))
	else
		response.write("<ul class=""menu"">" & chr(10))
		rsMenu.filter = "menu_parent_id = 0"
		do while not rsMenu.eof
			response.write("<li><a href=""?mod=cat&amp;cat_id=" & rsMenu("menu_id") & """>" & rsMenu("menu_name") & "</a></li>" & chr(10))
			rsMenu.movenext
		loop
		response.write("</ul>" & chr(10))
	end if	
	
	rsMenu.close
	set rsMenu = nothing
	
	rsSubmenu.close
	set rsSubmenu = nothing	
	%>
</div>
<br />
<div class="box">
<h2><%=strMenuInformation%></h2>
<ul class="menu">
<%
	'Check for custom pages
	set rsPages = server.createobject("ADODB.recordset")
	rsPages.cursortype = 3
	
	strSQL = "SELECT page_id, page_name FROM custom_pages WHERE page_lang_id = " & session("language_id") & " ORDER BY page_order ASC;"
	rsPages.open strSQL, adoCon
	
	do while not rsPages.eof
		response.write("<li><a href=""?mod=cpages&amp;page_id=" & rsPages("page_id") & """>" & rsPages("page_name") & "</a></li>" & chr(13))
		rsPages.movenext
	loop
%>
  <li><a href="?mod=contact"><%=strContact%></a></li>
<%
rsPages.close
set rsPages = nothing
%>
</ul>
</div>
<% if strShowManuModule = 1 then %>
<br />
<div class="box">
<h2><%=strMenuManufacturers%></h2>
<form method="get" action="" style="display:inline">
<p align="center">
  <%
  set rsMan = server.createobject("ADODB.recordset")
  rsMan.cursortype = 3
  
  strSQL = "SELECT manufacturer_id, manufacturer_name FROM manufacturer ORDER BY manufacturer_name ASC;"
  rsMan.open strSQL, adoCon
  %>
  <input type="hidden" name="mod" value="search" />
  <input type="hidden" name="type" value="advanced" />
  <select name="man_id" onchange="javascript:SearchMan(this.value);">
    <option value="0"><%=strAllManufacturers%></option>
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
</p>
</form>
</div>
<% end if %>
<% if strShowSearchModule = 1 then %>
<br />
<div class="box">
<h2><%=strSearch%></h2>
<form name="frmSearch" method="get" action="" style="display:inline">
<p align="center">
		  <input type="hidden" name="mod" value="search" />
		  <input type="hidden" name="type" value="simple" />
          <input name="q" type="text" value="" size="20" />
          <br />
	  <input type="submit" name="cmdSearch" value="<%=strSearch%>" class="cmd" />
</p>
</form>
</div>
<% end if %>
<% if strShowLangModule = 1 then %>
<br>
<div class="box">
<h2><%=strLanguage%></h2>
<form action="" method="post" name="frmLanguage" id="frmLanguage" style="display:inline">
	<p align="center">	
	  <%
	  set rsLanguage = server.createobject("ADODB.recordset")
	  rsLanguage.cursortype = 3
	  
	  strSQL = "SELECT language_id, language_name FROM lang WHERE language_show = -1"
	  rsLanguage.open strSQL, adoCon
	  %>
        <select name="language" id="language" onchange="javascript:SetLanguage(this.value);">
		  <%
		  do while not rsLanguage.eof
		    language_id     = rsLanguage("language_id")
			language_name   = rsLanguage("language_name")
		  %>
          <option value="<%=language_id%>"<% if session("language_id") = language_id then %> selected="selected"<% end if %>><%=language_name%></option>
		  <%
		    rsLanguage.movenext
		  loop
		  %>
        </select>&nbsp;<input name="cmdChangeLanguage" type="button" id="cmdChangeLanguage" value="<%=strCmdOK%>" class="cmd" onclick="javascript:SetLanguage(document.frmLanguage.language.value);" />
		<%
		rsLanguage.close
		set rsLanguage = nothing
		%>
</p>
</form>
</div>
<% end if %>
<% if strShowNewsletter = 1 then %>
<br />
<div class="box">
  <h2><%=strNewsLetter%></h2>
  <form action="default.asp?mod=newsletter" method="post" name="frmNewsletter">
  <p align="center">
    <%=strEmail%>:<br />
	<input name="email_adres" value="" />
  </p>
	<p align="left">
	<input type="radio" name="newsletter" value="subscribe" selected="selected" />&nbsp;<%=strSubscribe%><br />
	<input type="radio" name="newsletter" value="unsubscribe" />&nbsp;<%=strUnsubscribe%>
	</p>
	<p align="center">
	  <input type="submit" name="newsletter_submit" value="<%=strContinue%>" class="cmd" />
	</p>
  </form>
</div>
<% end if %>