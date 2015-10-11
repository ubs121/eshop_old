<%
if intModuleRights = 2 then
	set rsDelivery = server.createobject("ADODB.recordset")
	rsDelivery.cursortype = 3
	
	strSQL = "SELECT TOP 1 delivery_ID FROM delivery ORDER BY delivery_ID DESC;"
	rsDelivery.open strSQL, adoCon
	
	if not rsDelivery.eof then
		DID = rsDelivery("delivery_ID") + 1
	else
		DID = 1
	end if
	
	rsDelivery.close
	set rsDelivery = nothing
	
	response.redirect("?p=" & request.querystring("p") & "&action=edit&DID=" & DID & "&opt=" & request.form("slType") & "&cats=1")
else
	response.redirect("?p=" & request.querystring("p"))
end if
%>