<%
session("checkout_addressbook_id")  = null
session("checkout_order_comment")   = null
session("checkout_delivery_method") = 0

response.redirect("?mod=home")
%>