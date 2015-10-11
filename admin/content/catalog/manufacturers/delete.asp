<%
if intModuleRights = 2 then
	man_id = request.querystring("mid")
	
	if len(man_id) > 0 AND isnumeric(man_id) then
		strSQL = "DELETE * FROM manufacturer WHERE manufacturer_id = " & man_id
		adoCon.execute(strSQL)
	end if
end if
response.redirect("?p=" & request.querystring("p"))
%>