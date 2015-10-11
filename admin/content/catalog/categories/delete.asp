<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
menu_id = request.querystring("menu_id")

strSQL = "DELETE * FROM menu WHERE menu_id = " & menu_id
adoCon.Execute(strSQl)

response.redirect("?p=" & request.querystring("p"))
%>