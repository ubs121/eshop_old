<%
page_id = request.querystring("page_id")
if len(page_id) = 0 or not Isnumeric(page_id) then response.redirect("?mod=home")

set rsPage = server.createobject("ADODB.recordset")
rsPage.cursortype = 3

strSQL = "SELECT page_content FROM custom_pages WHERE page_id = " & page_id & " AND page_lang_ID = " & session("language_ID")
rsPage.open strSQL, adoCon

if not rsPage.eof then
	page_content = rsPage("page_content")
end if

rsPage.close
set rsPage = nothing
%>
<p class="pageheader"><%=page_name%></p>
<%=page_content%>