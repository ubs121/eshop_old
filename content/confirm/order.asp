<%
order_id = request.querystring("id")
code     = request.querystring("order_code")
intError = 0

set rsOrder = server.createobject("ADODB.recordset")
rsOrder.cursortype = 3

strSQL = "SELECT confirmed, salt, status, payment FROM orders WHERE order_id = " & order_id
rsOrder.open strSQL, adoCon

if not rsOrder.eof then
	if rsOrder("salt") = code then
		if cint(rsOrder("confirmed")) = -1 then
			intError = 3
		else
			intError = 0
			order_status = rsOrder("status")
			payment  = rsOrder("payment")
		end if
	else
		intError = 2
	end if
else
	intError = 1
end if

rsOrder.close
set rsOrder = nothing

if intError = 0 then
	'update the orderstatus
	set rsStatus = server.createobject("ADODB.recordset")
	rsStatus.cursortype = 3
	
	strSQL = "SELECT next_status FROM order_status WHERE order_status_id = " & order_status
	rsStatus.open strSQL, adoCon
	
	if not rsStatus.eof then
		next_status = cint(rsStatus("next_status"))
	else
		next_status = 0
	end if
	
	rsStatus.close
	set rsStatus = nothing
	
	if next_status = 0 then
		strSQL = "UPDATE orders SET confirmed = -1, date_confirmed = now() WHERE order_id = " & order_id & " AND salt = '" & code & "'"
	else
		strSQL = "UPDATE orders SET confirmed = -1, status = " & next_status & ", date_confirmed = now() WHERE order_id = " & order_id & " AND salt = '" & code & "'"
	end if
	adoCon.execute(strSQL)
	
	if strMailOrderConf = 1 then
		mailFrom  = strMailOrders
		MailTo    = strMailOrders
		mailHTML  = getFileContent(strVirtualPath & "languages/" & session("language") & "/templates/mail_orderconfirmed.html")
		MailBody  = transformOrdermail(order_id, mailHTML)
		SendMail()
	end if
	
	if len(payment) = 0 or not isnumeric(payment) then
		payment = 1
	else
		payment = cint(payment)
	end if
end if
%>
<p>
<%
response.write("<p align=""center"">")
select case intError
	case 0:
		response.write(Replace(strConfirmSucces, "[confirm_type]", strOrder))
	case 1:
		response.write(Replace(strConfirmFailed, "[confirm_type]", strOrder))
	case 2:
		response.write(Replace(strConfirmFailed, "[confirm_type]", strOrder))
	case 3:
		response.write(Replace(strConfirmFailed, "[confirm_type]", strOrder))
end select
response.write("</p>")

select case payment
	case 1:
		%><!-- #include file="orders/cash.asp" --><%
	case 2:
		%><!-- #include file="orders/paypal.asp" --><%
end select
%>
</p>