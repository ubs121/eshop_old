<%
subscribe_type = request.form("newsletter")
email_adres    = request.form("email_adres")

function checkMailAdres(email)
	if len(email) > 0 then
		if instr(email, "@") > 0 then
			checkMailAdres = 1
		else
			checkMailAdres = 0
		end if
	else
		checkMailAdres = 0
	end if
end function

'Subscribe
if subscribe_type = "subscribe" then
	if checkMailAdres(email_adres) = 1 then
		set rsNewsletter = server.createobject("ADODB.recordset")
		strSQL = "SELECT user_email, user_lang_id FROM newsletter WHERE user_email = '" & email_adres & "'"
		rsNewsletter.open strSQL, adoCon, 2, 2
		
		if not rsNewsletter.eof then
			rsNewsletter("user_lang_id") = session("language_id")
			rsNewsletter.update()
		else
			rsNewsletter.addnew()
				rsNewsletter("user_email") = email_adres
				rsNewsletter("user_lang_id") = session("language_id")
			rsNewsletter.update()
		end if
		
		rsNewsletter.close
		set rsNewsletter = nothing
		
		response.write("<p align=""center""><b>" & strSubscribed & "</b></p>")
	else
		response.write("error --> " & email_adres)
	end if
end if

'Unsubscribe
if subscribe_type = "unsubscribe" then
	if checkMailAdres(email_adres) = 1 then
		strSQL = "DELETE * FROM newsletter WHERE user_email = '" & email_adres & "'"
		adoCon.execute(strSQL)
		
		response.write("<p align=""center""><b>" & strUnsubscribed & "</b></p>")
	end if
end if
%>