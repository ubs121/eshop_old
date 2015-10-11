<%
address_id = request.querystring("id")

if len(address_id) > 0 AND IsNumeric(address_id) then
	address_id = cint(address_id)
else
	response.redirect("?mod=myaccount&sub=addressbook")
end if

strSQL = "DELETE * FROM user_address WHERE user_id = " & session("customer_id") & " AND user_address_id = " & address_id & " AND user_default_address = 0"
adoCon.execute(strSQL)

response.redirect("?mod=myaccount&sub=addressbook")
%>