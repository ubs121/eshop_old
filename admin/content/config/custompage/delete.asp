<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
pid = request.querystring("pid")
if len(pid) = 0 or not isnumeric(pid) then response.redirect("?p=" & request.querystring("p"))

strSQL = "DELETE * FROM custom_pages WHERE page_id = " & pid
adoCon.execute(strSQL)

response.redirect("?p=" & request.querystring("p"))
%>