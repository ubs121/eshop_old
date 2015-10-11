<%
set rsPage = server.createobject("ADODB.recordset")
rsPage.cursortype = 3

strSQL = "SELECT TOP 1 page_id FROM custom_pages ORDER BY page_id DESC;"
rsPage.open strSQL, adoCon

if not rsPage.eof then
	page_id = rsPage("page_id") + 1
else
	page_id = 1
end if

rsPage.close
set rsPage = nothing

response.redirect("?p=" & request.querystring("p") & "&action=edit&pid=" & page_id)
%>