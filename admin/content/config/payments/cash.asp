<% if request.querystring("opt") = "uninstall" then %>
<%
if intModuleRights = 2 then
	strSQL = "DELETE * FROM payment WHERE payment_ID = 1"
	adoCon.execute(strSQL)
end if

response.redirect("?p=" & request.querystring("p"))
%>
<% else %>
<%
if len(request.form()) > 0 then
	total_lang = request.form("totalLang")
	x          = 0
	
	for x = 1 to total_lang
		uid = cint(request.form("uid_" & x))
		c_name = Replace(request.form("name_" & x), "'", "''")
		lang_id = request.form("lang_id_" & x)
		
		if uid = 0 then
			strSQL = "INSERT INTO payment (payment_ID, payment_name, payment_lang_id) VALUES("
			strSQL = strSQL & "1,'" & c_name & "'," & lang_id & ");"
		else
			strSQL = "UPDATE payment SET payment_name = '" & c_name & "' WHERE payment_UID = " & uid
		end if
		adoCon.execute(strSQL)
	next
end if

set rsCash = server.createobject("ADODB.recordset")
rsCash.cursortype = 3

strSQL = "SELECT payment_UID, payment_name, payment_lang_id FROM payment WHERE payment_id = 1"
rsCash.open strSQL, adoCon

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_name, language_id FROM lang"
rsLang.open strSQL, adoCon
%>
<form name="form1" method="post" action="">
  <table width="100%" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
    <tr bgcolor="#666666"> 
      <td colspan="2"><strong><font color="#FFFFFF">Adjust payment details (cash)</font></strong></td>
    </tr>
    <tr> 
      <td colspan="2"><b>Translations:</b></td>
    </tr>
    <%
x = 0
do while not rsLang.eof
	x = x + 1
	
	rsCash.filter = "payment_lang_id = " & rsLang("language_id")
	if not rsCash.eof then
		cash_uid = rsCash("payment_uid")
		cash_name = rsCash("payment_name")
	else
		cash_uid = 0
		cash_name = ""
	end if
%>
    <tr> 
      <td width="120">&nbsp;<%=rsLang("language_name")%></td>
      <td>
	  	<input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
	    <input type="hidden" name="uid_<%=x%>" value="<%=cash_uid%>" /> 
        <input name="name_<%=x%>" type="text" id="name_<%=x%>" value="<%=cash_name%>" size="40"> 
      </td>
    </tr>
    <%
	rsLang.movenext
loop
%>
    <tr align="center"> 
      <td colspan="2">
	    <input type="hidden" name="totalLang" value="<%=rsLang.recordcount%>" />
	    <%=buildSubmitter("cmdSubmit", "Adjust payment details", request.querystring("p"))%>
        <input name="btnBack" type="button" id="btnBack" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';" /></td>
    </tr>
  </table>
</form>
<%
rsCash.close
set rsCash = nothing

rsLang.close
set rsLang = nothing
%>
<% end if %>