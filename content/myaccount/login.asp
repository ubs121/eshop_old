<%
if request.form("submit") = "ok" or request.form("submit2") = "ok" then
	intError = 0
	email = killChars(request.form("txtEmail"))
	password = killChars(request.form("txtPassword"))
	
	set rsLogin = server.createobject("ADODB.recordset")
	rsLogin.cursortype = 3
	
	if len(email) > 0 then
		strSQL = "SELECT user_id, user_salt, user_password FROM users WHERE user_email = '" & email & "'"
		rsLogin.open strSQL, adoCon
		
		if not rsLogin.eof then
			strSalt = rsLogin("user_salt")
			strSecret = hashEncode(password & strSalt)
			
			if strSecret = rsLogin("user_password") then
				session("customer_id") = rsLogin("user_id")
				session("sid") = strSecret
			else
				intError = 1
			end if
		else
			intError = 1
		end if
		rsLogin.close
		set rsLogin = nothing
		
		if intError = 0 then
			if request.querystring("red") = "" then
				red = "myaccount"
			else
				red = request.querystring("red")
			end if
			response.redirect("?mod=" & red)
		end if
	end if
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="49%">
	<% if session("regSucces") = true then %>
	<div class="regSucces">
	  <h2><%=strRegSucces%></h2>
	  <p>
	    <%=strRegSuccesPleaselogin%>
		<% session("regsucces") = false %>
	  </p>
	</div>
	<% else %>
	<div class="box_large">
	<h2><%=strNewCustomer%></h2>
	<p>
	  <%=strNewCustomer%>.<br /><br />
	  <%=Replace(strNewCustomerText, "[shopname]", strShopName)%><br /><br />
	    <p align="right"><a href="?mod=myaccount&amp;sub=register&amp;red=<%=request.querystring("red")%>"><img src="languages/<%=session("language")%>/images/button_continue.gif" alt="<%=strContinue%>" width="122" height="22" /></a></p>	
	</p>
	</div> 
	<% end if %>
    </td>
    <td width="40">&nbsp;</td>
    <td width="49%"> 
	  <div class="box_large">
	  <h2><%=strReturningCustomer%></h2>
	  <p>
		    <p><%=strReturningCustomer%>.</p>
			<form name="frmLogin" method="post" action="<%=strCurrPage%>">
              <table width="100%" cellspacing="0" cellpadding="0" border="0">
                <tr> 
                  <td width="100" class="content"><b><%=strEmail%>:</b></td>
                  <td class="content">
<input name="txtEmail" type="text" id="txtEmail" value="<%=email%>" size="20"></td>
                </tr>
                <tr> 
                  <td width="100" class="content"><b><%=strPassword%>:</b></td>
                  <td class="content"><input name="txtPassword" type="password" id="txtPassword" size="20">
                <input name="Submit2" type="submit" id="Submit2" style="visibility:hidden; height: 0px; width: 0px;" value="ok"></td>
                </tr>
              </table>
			  &nbsp;<a href="?mod=myaccount&amp;sub=lostpass"><%=strForgottenPassword%></a>
			  <a href="javascript:;" onclick="javascript:document.frmLogin.submit();">
                <input name="Submit" type="hidden" id="Submit" value="ok">
          </a> 
        </form>
            
        <p align="right"><a href="javascript:document.frmLogin.submit();"><img src="languages/<%=session("language")%>/images/button_login.gif" alt="<%=strLogin%>" width="122" height="22" border="0" /></a></p>
	  </p>
	  </div>
    </td>
  </tr>
</table>
