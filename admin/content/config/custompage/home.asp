<%
if len(request.form()) > 0 then
	total_pages = request.form("total_pages")
	
	x = 0
	for x = 1 to total_pages
		page_id = request.form("page_id_" & x)
		page_order = request.form("slOrder_" & x)
		
		strSQL = "UPDATE custom_pages SET page_order = " & page_order & " WHERE page_id = " & page_id
		adoCon.Execute(strSQL)
	next
end if
bgcolor = "#EEEEEE"

set rsCpages = server.createobject("ADODB.recordset")
rsCpages.cursortype = 3

strSQL = "SELECT page_id, page_name, page_order FROM custom_pages WHERE page_lang_id = " & default_lang_id & " ORDER BY page_order ASC;"
rsCpages.open strSQL, adoCon

totalPages = rsCpages.recordcount
%>
<form name="frmOrderPages" action="" method="post">
<table width="714" align="center" cellpadding="2" cellspacing="2">
  <%
curpage = 0
do while not rsCpages.eof
	curpage = curpage + 1
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
%>
  <tr bgcolor="<%=bgcolor%>"> 
    <td width="20">
			  <%
			  x = 0
			  response.write("<input type=""hidden"" name=""page_id_" & curpage & """ value=""" & rsCpages("page_id") & """ />" & chr(10))
			  response.write("<select name=""slOrder_" & curpage & """>" & chr(10))
			  for x = 1 to totalPages
			  	response.write("  <option value=""" & x & """")
				if x = cint(rsCPages("page_order")) then
					response.write(" selected=""selected""")
				end if
				response.write(">" & x & "</option>" & chr(10))
			  next
			  response.write("</select>" & chr(10))
			  %> </td>
    <td><%=rsCpages("page_name")%></td>
    <td width="50"> <% if intModuleRights = 2 then %> <a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;pid=<%=rsCpages("page_id")%>">edit</a> 
      <% else %>
      edit 
      <% end if %> </td>
    <td width="50"> <% if intModuleRights = 2 then %> <a href="?p=<%=request.querystring("p")%>&amp;action=delete&amp;pid=<%=rsCpages("page_id")%>">delete</a> 
      <% else %>
      delete 
      <% end if %> </td>
  </tr>
  <%
	rsCpages.movenext
loop
%>
  <tr align="left"> 
      <td colspan="4"><img src="images/arrow_down.gif" width="20" height="20" /><%=buildSubmitter("cmdSubmit", "Update order", request.querystring("p"))%></td>
  </tr>
  <tr> 
    <td colspan="4" align="center"> <% if intModuleRights = 2 then %> <a href="?p=<%=request.querystring("p")%>&amp;action=add">Add 
      a custom page</a> <% end if %> </td>
  </tr>
</table>
<input type="hidden" name="total_pages" value="<%=totalPages%>" />
</form>
<%
rsCpages.close
set rsCpages = nothing
%>