<%
if intModuleRights = 2 then
	pid = request.querystring("pid")
	
	if len(pid) = 0 or not isnumeric(pid) then
		response.redirect("?p=" & request.querystring("p"))
	end if
	
	strSQL = "DELETE * FROM products WHERE product_id = " & pid
	adoCon.execute(strSQL)
	
	strSQL = "DELETE * FROM product_info WHERE product_id = " & pid
	adoCon.execute(strSQL)
	
	strSQL = "DELETE * FROM product_description WHERE product_id = " & pid
	adoCon.execute(strSQL)
	
	response.redirect("?p=" & request.querystring("p"))
end if
%>