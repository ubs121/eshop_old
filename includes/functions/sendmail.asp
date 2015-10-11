<%
function transformOrdermail(order_ID, mailHTML)
	'Initialise values
	address_id      = 0
	total_price     = 0
	date_ordered    = ""
	order_code      = 0
	comment         = ""
	payment         = 0
	delivery_method = ""
	
	'Get mailcontent
	temp = mailHTML
	
	'Get general orderinformation
	set rsOrder = server.createobject("ADODB.recordset")
	rsOrder.cursortype = 3
	
	strSQL = "SELECT address_id, total_price, date_ordered, salt, comment, payment FROM orders WHERE order_ID = " & order_ID
	rsOrder.open strSQL, adoCon
	
	if not rsOrder.eof then
		address_id   = rsOrder("address_id")
		total_price  = strCurrency & RoundNumber(rsOrder("total_price"))
		date_ordered = rsOrder("date_ordered")
		comment      = rsOrder("comment")
		payment      = rsOrder("payment")
		order_code   = rsOrder("salt")
	end if
	
	rsOrder.close
	set rsOrder = nothing
	
	'Get customer info
	set rsCustomer = server.createobject("ADODB.recordset")
	rsCustomer.cursortype = 3
	
	strSQL = "SELECT user_firstname, user_lastname, user_street, user_postcode, user_city, user_province, user_country FROM user_address WHERE user_address_ID = " & address_ID
	rsCustomer.open strSQL, adoCon
	
	if not rsCustomer.eof then
		customer_name = rsCustomer("user_lastname") & " " & rsCustomer("user_firstname")
		customer_address = rsCustomer("user_street") & "<br />" & chr(10) & _
			rsCustomer("user_postcode") & " " & rsCustomer("user_city") & "<br />" & chr(10) & _
			rsCustomer("user_province") & ", " & rsCustomer("user_country")
	end if
	
	rsCustomer.close
	set rsCustomer = nothing
	
	'Get payment information
	set rsPayment = server.createobject("ADODB.recordset")
	rsPayment.cursortype = 3
	
	strSQL = "SELECT payment_name FROM payment WHERE payment_ID = " & payment & " AND payment_lang_id = " & session("language_ID")
	rsPayment.open strSQL, adoCon
	
	if not rsPayment.eof then
		temp = Replace(temp, "[payment-method]", rsPayment("payment_name"))
	end if
	
	rsPayment.close
	set rsPayment = nothing
	
	'Get all the products that have been ordered
	set rsOrderInfo = server.createobject("ADODB.recordset")
	rsOrderInfo.cursortype = 3
	
	strSQL = "SELECT product_ID, product_type, products_ordered, product_name, product_price FROM order_info WHERE order_ID = " & order_ID
	rsOrderInfo.open strSQL, adoCon
	
	rsOrderInfo.filter = "product_type = 'delivery'"
	if not rsOrderInfo.eof then
		set rsDelivery = server.createobject("ADODB.recordset")
		rsDelivery.cursortype = 3
		
		strSQL = "SELECT delivery_name, a, b FROM delivery WHERE delivery_ID = " & rsOrderInfo("product_ID") & " AND lang_ID = " & session("language_ID")
		rsDelivery.open strSQL, adoCon
		
		if not rsDelivery.eof then
			if rsDelivery("a") = "1" then
				delivery_price = roundNumber(Replace(rsDelivery("b"), ".", strServerComma))
			else
				arrPrices = split(rsDelivery("b"), ";")
				arrConditions = split(rsDelivery("a"), ";")
					
				x = 0
				for x = 0 to ubound(arrConditions)
						if instr(arrConditions(x), ">") > 0 then
						condition = csng(Replace(right(arrConditions(x), len(arrConditions(x)) - 1), ".", strServerComma))
						if csng(session("totalWeight")) > condition then
							delivery_price = csng(replace(arrPrices(x), ".", strServerComma))
						end if
					else
						if csng(session("totalWeight")) < csng(right(arrConditions(x), len(arrConditions(x)) - 1)) then
							delivery_price = csng(replace(arrPrices(x), ".", strServerComma))
						end if
					end if					
				next
			end if
		
			delivery_name  = rsDelivery("delivery_name")
		end if

		rsDelivery.close
		set rsDelivery = nothing
	end if
	
	'Transform products in an order
	rsOrderinfo.filter = "product_type = 'product'"
	if  instr(mailHTML, "[products-ordered]") > 0 then
		productsOrdered = ""
		
		productsOrdered = "<table width=""100%"" cellspacing=""0"" cellpadding=""4"" class=""productsOrdered-table"">" & chr(10)
		do while not rsOrderinfo.eof
			productsOrdered = productsOrdered & "<tr>" & chr(10) & _
				"  <td class=""products_productname"">" & rsOrderinfo("product_name") & "</td>" & chr(10) & _
				"  <td class=""products_productordered"" align=""right"">" & rsOrderinfo("products_ordered") & "</td>" & chr(10) & _
				"  <td width=""40"" align=""center"">x</td>" & chr(10) & _
				"  <td class=""products_productprice"">" & strCurrency & roundNumber(rsOrderinfo("product_price")) & "</td>" & chr(10) & _
				"</tr>" & chr(10)
			rsOrderinfo.movenext
		loop
		productsOrdered = productsOrdered & "<tr>" & chr(10) & _
			"  <td class=""products_productname"">" & delivery_name & "</td>" & chr(10) & _
			"  <td class=""products_productordered"" align=""right"">1</td>" & chr(10) & _
			"  <td width=""40"" align=""center"">x</td>" & chr(10) & _
			"  <td class=""products_productprice"">" & strCurrency & roundNumber(delivery_price) & "</td>" & chr(10) & _
			"</tr>" & chr(10)
			
		productsOrdered = productsOrdered & "</table>" & chr(10)
	end if
	
	rsOrderinfo.close
	set rsOrderinfo = nothing
	temp = replace(temp, "[products-ordered]", productsOrdered)
	
	'Transform customername
	temp = replace(temp, "[customer-name]", customer_name)
	
	'Transform delivery-method
	temp = replace(temp, "[delivery-method]", delivery_name)
	
	'Transform delivery-address
	temp = replace(temp, "[delivery-address]", customer_address)
	
	'Transform price
	temp = replace(temp, "[total-price]", total_price)
	
	'Transform orderdate
	temp = replace(temp, "[date-ordered]", date_ordered)
	
	'Transform comments
	temp = replace(temp, "[comments]", comment)
	
	'Transform shopname
	temp = replace(temp, "[shopname]", strShopName)
	
	'Transform confirmation-link
	confirmlink = strShopLink & "?mod=confirm&amp;type=order&amp;id=" & order_id & "&amp;order_code=" & order_code
	temp = Replace(temp, "[confirmation-link]", confirmlink)
	
	'Transform order-id
	temp = replace(temp, "[order-id]", order_ID)
	
	'Transform IP
	temp = replace(temp, "[user-ip]", request.servervariables("REMOTE_ADDR"))

	transformOrdermail = temp
end function

Private sub SendMail()
	select case strMailMethod
		case "cdonts":
			SendCdonts()
		case "cdo":
			SendCdosys()
		case "dundas":
			SendDundas()
		case "jmail":
			SendJmail()
		case "persits":
			SendPersits()
		case "aspmail":
			SendAspmail()
	end select
end sub

private sub sendAspmail()
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	
	Mailer.FromAddress = MailFrom
	if len(strMailOut) > 0 then
		Mailer.RemoteHost = strMailOut
	else
		Mailer.RemoteHost = "localhost"
	end if
	
	Mailer.AddRecipient "", MailTo
	Mailer.Subject = MailSubject
	Mailer.BodyText = MailBody
	Mailer.ContentType = "text/html"
	
	Mailer.sendmail
end sub

private sub SendPersits()
  Dim Message 'As New cdonts.NewMail
 
  'Create CDO message object
  Set Message = Server.CreateObject("Persits.MailSender")
  With Message
  
  if len(strMailOut) > 0 then
 	Message.Host = strMailOut
  end if

 Message.FromName = MailFrom ' Specify sender's name
 Message.AddAddress MailTo

 Message.Subject = MailSubject
 Message.IsHTML = True
 Message.Body = MailBody & Chr(13) & Chr(10)
 On Error Resume Next
 Message.Send 
  
    'Send the message
    .Send
  End With
End Sub

Private sub SendCdosys()
	' Create the e-mail server object
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	' Outgoing SMTP server
	if len(strMailOut) = 0 then
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	else
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailOut
	end if
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objCDOSYSCon.Fields.Update
	
	' Update the CDOSYS Configuration
	Set objCDOSYSMail.Configuration = objCDOSYSCon
	objCDOSYSMail.From = MailFrom
	objCDOSYSMail.To = MailTo
	objCDOSYSMail.Subject = MailSubject
	objCDOSYSMail.HTMLBody = MailBody
	objCDOSYSMail.Send
	'Close the server mail object
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing 
end sub

private sub SendCdonts()
  Const CdoBodyFormatHTML = 1
  Dim Message 'As New cdonts.NewMail
  
  'Create CDO message object
  Set Message = CreateObject("cdonts.NewMail")
  With Message
    
    'Set email adress, subject And body
    .To = MailTo
    .Subject = MailSubject
    .Body = MailBody
    
    'set mail And body format
    .MailFormat = CdoMailFormatHTML
    .BodyFormat = CdoBodyFormatHTML
    
    'Set sender address If specified.
    .From = MailFrom
    
    'Send the message
    .Send
  End With
End Sub

private sub SendDundas()
	dim objDundasMail
	Set objDundasMail = Server.CreateObject("Dundas.Mailer")
	
	objDundasMail.TOs.Add MailTo
	objDundasMail.Subject = MailSubject
	objDundasMail.FromAddress = MailFrom
	if len(strMailOut) > 0 then
		objDundasMail.SMTPRelayServers.Add strMailOut
	end if
	objDundasMail.HTMLBody = MailBody
	
	objDundasMail.SendMail
	
	set objDundasMail = nothing
end sub

private sub SendJMail()
	set objJmail = Server.CreateOBject( "JMail.Message" )
	
	objJmail.logging = false
	objJmail.silent  = true
	
	objJmail.from = MailFrom
	objJmail.AddRecipient MailTo
	objJmail.subject = MailSubject
	objJmail.HTMLbody = MailBody
	
	if len(strMailOut) > 0 then
		objJmail.send(strMailOut)
	end if
	
	set objJmail = nothing
end sub
%>