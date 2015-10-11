<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
id = request.querystring("id")

if len(id) > 0 and isnumeric(id) then
	strSQL = "UPDATE lang SET language_show = 0 WHERE language_id = " & id & " AND language_default = 0"
	adoCon.execute(strSQL)
end if
response.redirect("?p=" & request.querystring("p") & "&act=writelangfile")
%>