<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
admin_id = request.querystring("admin_id")
if len(admin_id) > 0 and IsNumeric(admin_id) then
	if cint(admin_id) <> cint(session("admin_user_id")) then
		strSQL = "DELETE * FROM admin_users WHERE admin_id = " & admin_id
		adoCon.execute(strSQL)
		strSQL = "DELETE * FROM admin_rights WHERE admin_id = " & admin_id
		adoCon.execute(strSQL)
	end if
end if
response.redirect("?p=" & request.querystring("p"))
%>