<%
nid = request.querystring("nid")
if isnumeric(nid) and len(nid) > 0 and intModuleRights = 2 then
	strSQL = "DELETE * FROM news WHERE news_id = " & nid
	adoCon.execute(strSQL)
end if

response.redirect("?p=" & request.querystring("p"))
%>