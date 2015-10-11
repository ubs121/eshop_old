<%
do while instr(page,"/") > 0
	page = right(page, len(page) - instr(page,"/"))
loop

set rsModules = server.createobject("ADODB.recordset")
rsModules.cursortype = 3

strSQL = "SELECT module_name, module_id, module_page FROM admin_modules ORDER BY module_order ASC;"
rsModules.open strSQL, adoCon
%>  
  <div class="menu">
    <h1>Quick menu</h1>
    <h2 onclick="document.location='config.asp';">Configuration</h2>
	<% rsModules.filter = "module_page = 'config.asp'" %>
	<% if page = "config.asp" then %>
    <ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="config.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
	<h2 onclick="document.location='catalog.asp';">Catalog</h2>
	<% rsModules.filter = "module_page = 'catalog.asp'" %>
	<% if page="catalog.asp" then %>
	<ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="catalog.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
	<h2 onclick="document.location='customers.asp';">Customers</h2>
	<% rsModules.filter = "module_page = 'customers.asp'" %>
	<% if page = "customers.asp" then %>
	<ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="customers.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
	<h2 onclick="document.location='localization.asp';">Localization</h2>
	<% rsModules.filter = "module_page = 'localization.asp'" %>
	<% if page="localization.asp" then %>
	<ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="localization.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
	<h2>Statistics</h2>
	<% rsModules.filter = "module_page = 'stats.asp'" %>
	<% if page = "stats.asp" then %>
	<ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="stats.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
	<h2 onclick="document.location='login.asp';">Administration</h2>
	<% rsModules.filter = "module_page = 'login.asp'" %>
	<% if page = "login.asp" then %>
	<ul>
	  <% do while not rsModules.eof %>
	  <% if cint(request.cookies("admin_rights")("module_" & rsModules("module_id"))) > 0 then %>
	  <li><a href="login.asp?p=<%=rsModules("module_id")%>"><%=rsModules("module_name")%></a></li>
	  <% end if %>
	  <%
	    rsModules.movenext
	  loop
	  %>
	</ul>
	<% end if %>
  </div>
<%
rsModules.close
set rsModules = nothing
%>