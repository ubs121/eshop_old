<%
msg = ""
if len(request.form()) > 0 then
	fullname = request.form("fullname")
	email    = request.form("email")
	enquiry  = request.form("enquiry")
	errorId  = 0
	
	if len(fullname) = 0 then
		errorId = 1
		msg = strGiveFullname
	end if
	if len(email) = 0 then
		errorId = 1
		msg = strGiveEmail
	end if
	if len(enquiry) = 0 then
		errorId = 1
		msg = strGiveEnquiry
	end if
	
	if errorId = 0 then
		MailTo = strMailInfo
		MailFrom = email
		MailBody = Replace (strHasSent, "[name]", fullname ) & ":<br />"
		MailBody = MailBody & "----------------------------------------------------<br />"
		MailBody = MailBody & enquiry
		MailSubject = strContactSubject
		
		SendMail
		
		msg = "<b>" & strEmailSent & "</b>"
	else
		msg = "<b><font color=""#FF0000"">[!] " & msg & "</font></b>"
	end if
end if
%>
<p class="pageheader"><%=strContact%></p>
<form name="frmSendMail" action="<%=strCurrPage%>" method="post">
  <table cellspacing="2" cellpadding="2" class="contact">
  <% if len(msg) > 1 then %>
    <tr>
	  <td><%=msg%></td>
	</tr>
  <% end if %>
    <tr> 
      <td><%=strFullName%>:</td>
    </tr>
    <tr> 
      <td><input name="fullname" type="text" id="fullname"></td>
    </tr>
    <tr> 
      <td><%=strEmailAddress%>:</td>
    </tr>
    <tr> 
      <td><input name="email" type="text" id="email"></td>
    </tr>
    <tr> 
      <td><%=strEnquiry%>:</td>
    </tr>
    <tr> 
      <td><textarea name="enquiry" cols="80" rows="8" id="enquiry"></textarea></td>
    </tr>
  </table>
  <br />
  <table width="600" border="0" cellspacing="0" cellpadding="0" class="productListing">
    <tr> 
      <td align="right"><a href="javascript:document.frmSendMail.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
    </tr>
  </table>
</form>
