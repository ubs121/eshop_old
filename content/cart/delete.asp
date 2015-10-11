<%
cookie_id = killChars(request.querystring("id"))
total_products = killChars(request.cookies("total_products" & session.SessionID))
quant = killChars(request.querystring("quant"))

if len(quant) = 0 OR NOT Isnumeric(quant) then
	quant = 1
else
	quant = cint(quant)
end if

if len(cookie_id) > 0 and IsNumeric(cookie_id) then
	cookie_id = cint(cookie_id)
elseif cookie_id <> "all" then
	response.redirect("?mod=cart&action=view")
end if

if cookie_id = "all" then
	total_products = 0	
else
	products_ordered = cint(request.cookies("product" & cookie_id & session.SessionID)("product_ordered"))
	
	if (products_ordered - quant) <= 0 then
		new_product_id = cint(request.cookies("product" & total_products & session.sessionID)("product_id"))
		new_product_ordered = cint(request.cookies("product" & total_products & session.sessionID)("product_ordered"))
		
		response.cookies("product" & cookie_id & session.SessionID)("product_id") = new_product_id
		response.cookies("product" & cookie_id & session.SessionID)("product_ordered") = new_product_ordered
		
		total_products = total_products - 1
	else
		response.cookies("product" & cookie_id & session.sessionID)("product_ordered") = cint(request.cookies("product" & cookie_id & session.SessionID)("product_ordered")) - quant
	end if
end if

response.cookies("total_products" & session.SessionID) = total_products
response.redirect("?mod=cart&action=view")
%>