<%
id = request.querystring("id")

if intModuleRights < 2 then response.redirect("?p=" & request.querystring("p"))

if isnumeric(id) and len(id) > 0 then
	id = cint(id)
	set rsOrderstatus = server.createobject("ADODB.recordset")
	rsOrderstatus.cursortype = 3
	
	strSQL = "SELECT TOP 1 order_status_id FROM order_status ORDER BY order_status_id DESC;"
	rsOrderstatus.open strSQL, adoCon
	
	if not rsOrderstatus.eof then
		orderstatus_high = cint(rsOrderstatus("order_status_id"))
	else
		orderstatus_high = 0
	end if
	
	rsOrderstatus.close
	set rsOrderstatus = nothing
	
	if orderstatus_high = id then
		strSQL = "DELETE * FROM order_status WHERE order_status_id = " & id
		adoCon.execute(strSQL)
		response.redirect("?p=" & request.querystring("p"))
	else
		strSQL = "DELETE * FROM order_status WHERE order_status_id = " & id
		adoCon.execute(strSQL)
		for x = id to orderstatus_high
			strSQL = "UPDATE order_status SET order_status_id = " & x - 1 & " WHERE order_status_id = " & x
			adoCon.execute(strSQL)
		next
		response.redirect("?p=" & request.querystring("p"))
	end if
end if
%>