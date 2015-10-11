<%
product_id = killChars(request.querystring("id"))
total_products = killChars(request.cookies("total_products" & session.SessionID))
quant = killChars(request.querystring("quant"))

if len(quant) = 0 OR not Isnumeric(quant) then
	quant = 1
else
	quant = cint(quant)
end if

if len(total_products) = 0 then
	total_products = 0
else
	total_products = cint(total_products)
end if

if IsNumeric(product_id) AND len(product_id) > 0 then
	product_id = cint(product_id)
else
	response.redirect("?mod=cart&action=view")
end if

product_isnew = 1

if total_products > 0 then
	for x = 1 to total_products
		if product_isnew = 0 then exit for
		cookie_product_id = cint(request.cookies("product" & x & session.SessionID)("product_id"))
		if cookie_product_id = product_id then
			product_isnew = 0
		end if
		product_cookie = x
	next
else
	product_cookie = 0
end if

if product_isnew = 1 then
	product_cookie = product_cookie + 1
	total_products = total_products + 1
	response.cookies("product" & product_cookie & session.SessionID)("product_id") = product_id
	response.cookies("product" & product_cookie & session.SessionID)("product_ordered") = quant
	response.cookies("total_products" & session.SessionID) = total_products
else
	response.cookies("product" & product_cookie & session.SessionID)("product_ordered") = cint(request.cookies("product" & product_cookie & session.SessionID)("product_ordered")) + quant
end if

response.redirect("?mod=cart&action=view")
%>
