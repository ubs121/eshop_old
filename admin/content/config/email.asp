<%
if len(request.form("Submit")) > 0 then
	updateConfig("mail_out")
	updateConfig("mail_method")
	updateConfig("mail_info")
	updateConfig("mail_orders")
	updateConfig("mail_noreply")
	updateConfig("mail_sendOrderconfirmed")
end if

set rsConfig = server.createobject("ADODB.recordset")
rsConfig.cursortype = 3

strSQL = "SELECT config_name, config_value FROM config"
rsConfig.open strSQL, adoCon
%>
<form action="" method="post" name="frmEmail" id="frmEmail">
<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Mail Setup</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>

  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="299">Mail method:<% mail_method = getConfig("mail_method") %></td>
      <td width="299"><select name="mail_method">
          <option value="cdo"<% if mail_method = "cdo" then %> selected="selected"<% end if %>>CDO</option>
          <option value="cdonts"<% if mail_method = "cdonts" then %> selected="selected"<% end if %>>cdonts</option>
          <option value="dundas"<% if mail_method = "dundas" then %> selected="selected"<% end if %>>dundas</option>
          <option value="jmail"<% if mail_method = "jmail" then %> selected="selected"<% end if %>>jmail</option>
		  <option value="persits"<% if mail_method = "persits" then %> selected="selected"<% end if %>>Persits</option>
		  <option value="aspmail"<% if mail_method = "aspmail" then %> selected="selected"<% end if %>>aspMail</option>
        </select></td>
    </tr>
    <tr> 
      <td>SMTP-server:</td>
      <td><input name="mail_out" type="text" id="mail_out" value="<%=getConfig("mail_out")%>" size="40"></td>
    </tr>
    <tr> 
      <td colspan="2" height="10"></td>
    </tr>
    <tr> 
      <td>Emailaddress information:</td>
      <td><input name="mail_info" type="text" id="mail_info" value="<%=getConfig("mail_info")%>" size="40"></td>
    </tr>
    <tr>
      <td>Emailaddress noreply</td>
      <td><input name="mail_noreply" type="text" id="mail_noreply" value="<%=getConfig("mail_noreply")%>" size="40"></td>
    </tr>
    <tr> 
      <td>Emailaddress orders:</td>
      <td><input name="mail_orders" type="text" id="mail_orders" value="<%=getConfig("mail_orders")%>" size="40"></td>
    </tr>
    <tr align="center">
      <td align="left"><% sendOrderconfirmed = cint(getConfig("mail_sendOrderconfirmed")) %>
        Send email to emailaddress orders when order is confirmed </td>
	  <td align="left" valign="top"><select name="mail_Sendorderconfirmed">
	    <option value="1"<% if sendOrderconfirmed = 1 then %> selected="selected"<% end if %>>Yes</option>
	    <option value="0"<% if sendOrderconfirmed = 0 then %> selected="selected"<% end if %>>No</option>
	    </select>
	  </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Update", request.querystring("p"))%> &nbsp;&nbsp; <input type="reset" name="Submit2" value="Reset"></td>
    </tr>
  </table>
</td></tr>
</table>
</center>
</form>
<%
rsConfig.close
set rsConfig = nothing
%>