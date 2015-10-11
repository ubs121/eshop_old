<%
if request.form("agree") <> strIAgree then
	response.redirect("?mod=checkout&p=3")
end if

total_price = csng(request.form("total_price"))
status_id   = 1
salt        = getSalt(len(session.SessionID))

set rsOrder = server.createobject("ADODB.recordset")
rsOrder.open "orders", adoCon, 2, 2

rsOrder.addnew()
	rsOrder("user_id") = session("customer_id")
	rsOrder("total_price") = total_price
	rsOrder("date_ordered") = now()
	rsOrder("status") = status_id
	rsOrder("salt")   = salt
	rsOrder("address_id") = session("checkout_addressbook_id")
	rsOrder("comment") = session("checkout_order_comment")
	rsOrder("payment") = session("checkout_payment_method")
rsOrder.update()

order_id = rsOrder("order_id")

rsOrder.close
set rsOrder = nothing

set rsProducts = server.createobject("ADODB.recordset")
rsProducts.cursortype = 3

strSQL = "SELECT product_ID, product_name, newPrice FROM products"
rsProducts.open strSQL, adoCon

set rsOrderinfo = server.createobject("ADODB.recordset")
rsOrderinfo.open "order_info", adoCon, 2, 2

total_products = cint(request.cookies("total_products" & session.SessionID))

for x = 1 to total_products
	product_id = cint(request.cookies("product" & x & session.SessionID)("product_id"))
	product_ordered = cint(request.cookies("product" & x & session.SessionID)("product_ordered"))
	
	rsProducts.filter = "product_ID = " & product_id
	product_name = rsProducts("product_name")
	product_price = Replace(rsProducts("newPrice"), ".", strServerComma)
	
	rsOrderinfo.addnew()
		rsOrderinfo("order_id") = order_id
		rsOrderinfo("product_id") = product_id
		rsOrderinfo("products_ordered") = product_ordered
		rsOrderinfo("product_type") = "product"
		rsOrderinfo("product_name") = product_name
		rsOrderinfo("product_price") = Replace(product_price, ",", ".")
	rsOrderinfo.update()
next

rsProducts.close
set rsProducts = nothing

set rsDelivery = server.createobject("ADODB.recordset")
rsDelivery.cursortype = 3

strSQL = "SELECT a, b, delivery_name FROM delivery WHERE lang_id = " & session("language_id") & " AND delivery_id = " & session("checkout_delivery_method")
rsDelivery.open strSQL, adoCon

if not rsDelivery.eof then
	if rsDelivery("a") = "1" then
		delivery_price = csng(Replace(rsDelivery("b"), ".", strServerComma))
	else
		arrPrices = split(rsDelivery("b"), ";")
		arrConditions = split(rsDelivery("a"), ";")
		
		x = 0
		for x = 0 to ubound(arrConditions)
			if instr(arrConditions(x), ">=") > 0 then
				condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 2), ".", strServerComma))
				if csng(session("totalWeight")) >= condition then
					delivery_price = arrPrices(x)
				end if
			elseif instr(arrConditions(x), ">") > 0 then
				condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 1), ".", strServerComma))
				if csng(session("totalWeight")) > condition then
					delivery_price = arrPrices(x)
				end if
			else
				if csng(session("totalWeight")) < csng(right(arrConditions(x), len(arrConditions(x)) - 1)) then
					delivery_price = arrPrices(x)
				end if
			end if					
		next
	end if
	delivery_name = rsDelivery("delivery_name")
end if

rsDelivery.close
set rsDelivery = nothing

rsOrderinfo.addnew()
	rsOrderinfo("order_id") = order_id
	rsOrderinfo("product_id") = session("checkout_delivery_method")
	rsOrderinfo("products_ordered") = 1
	rsOrderinfo("product_type") = "delivery"
rsOrderinfo.update()

rsOrderinfo.close
set rsOrderinfo = nothing

response.cookies("total_products" & session.SessionID) = 0
session("checkout_order_comment") = ""
session("checkout_delivery_method") = 0

'Send email with confirmation code
	set rsEmail = server.createobject("ADODB.recordset")
	rsEmail.cursortype = 3
	
	strSQL = "SELECT user_email FROM users WHERE user_id = " & session("customer_id")
	rsEmail.open strSQL, adoCon
	
	MailTo = rsEmail("user_email")
	
	rsEmail.close
	set rsEmail = nothing
	
	MailFrom  = strMailOrders
	MailSubject = Replace(strOrderSubject,"[shopname]", strShopName)

	mailHTML = getFileContent(strVirtualPath & "languages/" & session("language") & "/templates/mail_orderconfirmation.html")
	MailBody = transformOrdermail(order_id, mailHTML)
	SendMail()
%>
<p><b><%=strCheckoutFinished%></b></p>
<p>
  <%=Replace(strThanksForOrdering,"[shopname]", strShopname)%>
</p>