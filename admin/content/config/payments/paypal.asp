<% if request.querystring("opt") = "uninstall" then %>
<%
if intModuleRights = 2 then
	strSQL = "DELETE * FROM payment WHERE payment_ID = 2"
	adoCon.execute(strSQL)
end if
response.redirect("?p=" & request.querystring("p"))
%>
<% else %>
<%
if len(request.form()) > 0 then
	total_lang = request.form("totalLang")
	cc         = request.form("currency_code")
	business   = request.form("business")
	x          = 0
	
	options = cc & ";" & business
	
	for x = 1 to total_lang
		uid     = cint(request.form("uid_" & x))
		lang_id = request.form("lang_id_" & x)
		p_name  = request.form("name_" & x)
		
		if uid = 0 then
			strSQL = "INSERT INTO payment (payment_ID, payment_name, payment_lang_id, payment_options) VALUES("
			strSQL = strSQL & "2, '" & p_name & "'," & lang_id & ",'" & options & "');"
		else
			strSQL = "UPDATE payment SET payment_name = '" & p_name & "', payment_options = '" & options & "' WHERE payment_uid = " & uid
		end if
		adoCon.execute(strSQL)
	next
end if

set rsPaypal = server.createobject("ADODB.recordset")
rsPaypal.cursortype = 3

strSQL = "SELECT payment_UID, payment_options, payment_name, payment_lang_id FROM payment WHERE payment_ID = 2"
rsPaypal.open strSQL, adoCon

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_id, language_name FROM lang"
rsLang.open strSQL, adoCon

if not rsPaypal.eof then
	arrOptions = Split(rsPaypal("payment_options"), ";")
	
	cc = arrOptions(0)
	business = arrOptions(1)
end if
%>
<form name="form1" method="post" action="">
  <table width="100%" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
    <tr bgcolor="#666666"> 
      <td colspan="2"><strong><font color="#FFFFFF">Adjust payment details (Paypal)</font></strong></td>
    </tr>
    <tr> 
      <td colspan="2"><b>Paypal settings</b></td>
    </tr>
    <tr> 
      <td>Currency code:</td>
      <td><select name="currency_code" id="currency_code">
          <option value="USD"<% if cc = "USD" then %> selected="selected"<% end if %>>USD</option>
          <option value="EUR"<% if cc = "EUR" then %> selected="selected"<% end if %>>EUR</option>
          <option value="GBP"<% if cc = "GBP" then %> selected="selected"<% end if %>>GBP</option>
          <option value="CAD"<% if cc = "CAD" then %> selected="selected"<% end if %>>CAD</option>
          <option value="JPY"<% if cc = "JPY" then %> selected="selected"<% end if %>>JPY</option>
        </select></td>
    </tr>
    <tr> 
      <td width="120">
	  	Business<br />
      </td>
      <td valign="top"><input name="business" type="text" id="business" value="<%=business%>" size="40">
        (Email address on your PayPal account)</td>
    </tr>
    <tr> 
      <td colspan="2"><b>Translations:</b></td>
    </tr>
<%
x = 0
do while not rsLang.eof
	x = x + 1
	
	rsPaypal.filter = "payment_lang_id = " & rsLang("language_id")
	if not rsPaypal.eof then
		uid = rsPaypal("payment_uid")
		p_name = rsPaypal("payment_name")
	else
		uid = 0
		p_name = ""
	end if
%>
    <tr> 
      <td width="120">&nbsp;<%=rsLang("language_name")%></td>
      <td>
	  	<input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
		<input type="hidden" name="uid_<%=x%>" value="<%=uid%>" />
		<input name="name_<%=x%>" type="text" value="<%=p_name%>" size="40" />
	  </td>
    </tr>
<%
	rsLang.movenext
loop
%>
	<tr>
	  <td align="center" colspan="2">
	    <input type="hidden" name="totalLang" value="<%=rsLang.recordcount%>" />
	    <%=buildSubmitter("cmdSubmit", "Adjust paypal settings", request.querystring("p")) %>
		<input type="button" name="btnBack" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';" />
	  </td>
	</tr>
  </table>
</form>
<%
rsLang.close
set rsLang = nothing

rsPaypal.close
set rsPaypal = nothing
%>
<% end if %>