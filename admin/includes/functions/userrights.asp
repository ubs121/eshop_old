<%
function getRights(page)
	set rsCheck = server.createobject("ADODB.recordset")
	rsCheck.cursortype = 3
	getRights = 0
	
	if page = "login" then
		page = 0
		session("admin_user_id") = 0
	end if
	
	strSQL = "SELECT module_right FROM admin_rights WHERE module_id = " & page & " AND admin_id = " & session("admin_user_id")
	rsCheck.open strSQL, adoCon
	
	if not rsCheck.eof then
		getRights = cint(rsCheck("module_right"))
	else
		getRights = 0
	end if
end function

function BuildSubmitter(btnName, btnValue, module)
	module_right = getRights(request.querystring("p"))
	if module_right = 2 then
		response.write("<input type=""submit"" name=""" & btnName & """ value=""" & btnValue & """>")
	else
		response.write("<input type=""submit"" name=""" & btnName & """ value=""" & btnValue & """ disabled=""disabled"">")
	end if
end function

private function getPage(page_id)
	set rsPage = server.createobject("ADODB.recordset")
	rsPage.cursortype = 3
	
	strSQL = "SELECT module_sub FROM admin_modules INNER JOIN admin_rights ON admin_modules.module_id = admin_rights.module_id WHERE admin_modules.module_id = " & page_id & " AND module_right > 0 AND admin_id = " & session("admin_user_id")
	rsPage.open strSQL, adoCon
	
	if not rsPage.eof then
		getPage = rsPage("module_sub")
	else
		getPage = "accessDenied"
	end if
	
	rsPage.close
	set rsPage = nothing
end function
%>