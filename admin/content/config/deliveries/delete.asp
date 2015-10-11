<%
if intModuleRights = 2 then
	did = request.querystring("did")
	
	if len(did) > 0 and isnumeric(did) then
		strSQL = "DELETE * FROM delivery WHERE delivery_ID = " & did
		adoCon.execute(strSQL)
	end if
end if

response.redirect("?p=" & request.querystring("p"))
%>