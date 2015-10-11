<%
start     = session("start")
langs     = Replace(session("lang_init"), ";", ",")
completed = false

if len(langs) = 0 then response.redirect("?p=" & request.querystring("p"))
if len(start) > 0 and isnumeric(start) then
	start = cint(start)
else
	start = 1
end if

startP = ((start - 1) * 20) + 1

set rsMails = server.createobject("ADODB.recordset")
rsMails.cursortype = 3

strSQL = "SELECT user_email, user_lang_id FROM newsletter WHERE user_lang_id IN(" & langs & ") ORDER BY user_email ASC;"
rsMails.open strSQL, adoCon

rsMails.pagesize = 20

if not rsMails.eof then
	rsMails.absolutepage = start
else
	completed = true
end if

if start = rsMails.pagecount then
	completed = true
end if

totalMails = rsMails.recordcount
if (startP + 19) > totalMails then
	endP = totalMails - (startP - 1)
else
	endP = startP + 19
end if

mailFrom = mail_noreply
%>
<p><b>Sending newsletter (<%=startP & " - " & endP%>)</b></p>
<%
mailCount = 0
for mailCount = 1 to 20
	if rsMails.eof then exit for
	
	MailTo = rsMails("user_email")
	MailSubject = session("subject_" & rsMails("user_lang_id"))
	MailBody    = session("content_" & rsMails("user_lang_id"))
	on error resume next
		call sendMail()
	on error goto 0
	
	response.write("Newsletter sent to " & rsMails("user_email") & "<br />" & chr(10))
	
	rsMails.movenext
next

rsMails.close
set rsMails = nothing
%>
<p align="center">
<% if completed = true then %>
<% session("lang_init") = "" %>
<b>All newsletters have been sent</b><br />
<input type="button" name="cmdComplete" value="Complete" onclick="document.location='?p=<%=request.querystring("p")%>';" />
<% else %>
<% session("start") = start + 1 %>
<input type="button" name="cmdNext" value="Send next 20 newsletters >>" onclick="document.location='?p=<%=request.querystring("p")%>';" />
<% end if %>
</p>