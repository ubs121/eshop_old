<%
if len(request.form("user_id")) > 0 then
	new_subscription = cint(request.form("new_subscription"))
	newsletter  = request.form("newsletter")
	user_id = request.form("email_adres")
	lang_id = request.form("newsletter_lang_id")
	
	set rsEmail = server.createobject("ADODB.recordset")
	rsEmail.cursortype = 3
	
	strSQL = "SELECT user_email FROM users WHERE user_id = " & session("customer_id")
	rsEmail.open strSQL, adoCon
	
	user_email = rsEmail("user_email")
	
	rsEmail.close
	set rsEmail = nothing
	
	if new_subscription = 1 then
		if newsletter = "1" then
			strSQL = "INSERT INTO newsletter (user_email, user_lang_id) VALUES('" & user_email & "'," & lang_id & ");"
			adoCon.execute(strSQL)
		end if
	else
		if newsletter = "1" then
			strSQL = "UPDATE newsletter SET user_lang_id = " & lang_id & " WHERE user_email = '" & user_email & "'"
			adoCon.execute(strSQL)
		else
			strSQL = "DELETE * FROM newsletter WHERE user_email = '" & user_email & "'"
			adoCon.execute(strSQL)
		end if
	end if
end if

set rsNewsletter = server.createobject("ADODB.recordset")
rsNewsletter.cursortype = 3

strSQL = "SELECT user_lang_id FROM newsletter INNER JOIN users ON newsletter.user_email = users.user_email WHERE user_id = " & session("customer_id")
rsNewsletter.open strSQL, adoCon

if rsNewsletter.eof then
	subscripted      = 0
	new_subscription = 1
	subscr_lang_id   = session("language_id")
else
	subscripted      = 1
	new_subscription = 0
	subscr_lang_id   = cint(rsNewsletter("user_lang_id"))
end if

rsNewsletter.close
set rsNewsletter = nothing
%>
<form action="" method="post" name="frmUpdateNewsletter">
  <p><b><%=strNewsLetter%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td width="100%" class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="160" class="content"><%=strSubscribeForNewsletter%>:</td>
            <td class="content">
			  &nbsp;<input name="newsletter" type="checkbox" id="newsletter" value="1"<% if subscripted = 1 then %> checked="checked"<% end if %>>
              <input name="new_subscription" type="hidden" id="new_subscription" value="<%=new_subscription%>">
              <input name="user_id" type="hidden" id="user_id" value="<%=session("customer_id")%>"></td>
          </tr>
          <tr>
            <td class="content"><%=strLanguage%>:</td>
            <td class="content">&nbsp;
			  <select name="newsletter_lang_id" id="newsletter_lang_id">
			  <%
			  set rsLang = server.createobject("ADODB.recordset")
			  rsLang.cursortype = 3
			  
			  strSQL = "SELECT language_name, language_id FROM lang"
			  rsLang.open strSQL, adoCon
			  
			  do while not rsLang.eof
			  	response.write("<option value=""" & rsLang("language_id") & """")
				if cint(rsLang("language_id")) = subscr_lang_id then
					response.write(" selected=""selected""")
				end if
				response.write(">" & rsLang("language_name") & "</option>" & chr(13))
			  	rsLang.movenext
			  loop
			  
			  rsLang.close
			  set rsLang = nothing
			  %>			  
              </select>
			</td>
          </tr>
        </table></td>
    </tr>
  </table>
  <br />
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><a href="?mod=myaccount"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a> 
      </td>
      <td align="right"><a href="javascript:document.frmUpdateNewsletter.submit();"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" border="0" /></a></td>
    </tr>
  </table>
</form>