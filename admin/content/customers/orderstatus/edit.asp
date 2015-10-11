<%
id = request.querystring("id")
if len(id) > 0 and isnumeric(id) then
	id = clng(id)
else
	response.redirect("?p=" & request.querystring("p"))
end if

if len(request.form()) > 0 then
	order_default = request.form("chkDefault")
	next_status   = request.form("next_status")
	total_lang    = request.form("total_lang")
	
	if isnumeric(next_status) then
		next_status = cint(next_status)
	else
		next_status = 0
	end if	
	
	if order_default = "yes" then
		strSQL = "UPDATE order_status SET order_default = 0"
		adoCon.execute(strSQL)
		strSQL = "UPDATE order_status SET order_default = -1 WHERE order_status_id = " & id
		adoCon.execute(strSQL)
		order_default = -1
	else
		order_default = 0
	end if
	
	strSQL = "UPDATE order_status SET next_status = " & next_status & " WHERE order_status_id = " & id
	adoCon.execute(strSQL)
	
	x = 0
	for x = 1 to total_lang
		isnew = cint(request.form("isnew_" & x))
		order_status = request.form("orderstatus_" & x)
		lang_id      = request.form("lang_id_" & x)
		unique_id    = request.form("unique_" & x)
		
		if isnew = 1 then
			strSQL = "INSERT INTO order_status(order_status_id, order_status, lang_id, order_default, next_status) VALUES("
			strSQL = strSQL & id & ",'" & order_status & "'," & lang_id & "," & order_default & "," & next_status & ");"
		else
			strSQL = "UPDATE order_status SET order_status = '" & order_status & "' WHERE order_status_unique_id = " & unique_id
		end if
		adoCon.Execute(strSQL)
	next
end if
set rsOrderstatus = server.createobject("ADODB.recordset")
rsOrderstatus.cursortype = 3

strSQL = "SELECT order_status_unique_id, order_status_id, order_status, lang_id, order_default, next_status FROM order_status"
rsOrderstatus.open strSQL, adoCon

rsOrderstatus.filter = "order_status_id = " & id
if not rsOrderstatus.eof then
	order_default = cint(rsOrderstatus("order_default"))
	next_status   = cint(rsOrderstatus("next_status"))
end if

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_id, language_name FROM lang WHERE language_show = -1"
rsLang.open strSQL, adoCon
%>
<form name="frmEditOrderstatus" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="100">Default status:</td>
      <td><input name="chkDefault" type="checkbox" id="chkDefault" value="yes"<% if order_default = -1 then %> checked="checked" disabled="disabled"<% end if %> /></td>
    </tr>
    <tr> 
      <td>Next status :</td>
      <td>
	    <select name="next_status">
		  <option value="0">None</option>
		  <%
		  rsOrderstatus.filter = "lang_id = " & default_lang_id
		  do while not rsOrderstatus.eof
		  	response.write("<option value=""" & rsOrderstatus("order_status_id") & """")
			if cint(rsOrderstatus("order_status_id")) = next_status then
				response.write(" selected=""selected""")
			end if
			response.write(">" & rsOrderstatus("order_status") & "</option>" & chr(10))
		  	rsOrderstatus.movenext
		  loop
		  %>
		</select>
	  </td>
    </tr>
	<tr>
	  <td colspan="2" height="10"></td>
	</tr>
	<%
	x = 0
	do while not rsLang.eof
		x = x + 1
		rsOrderstatus.filter = "order_status_id = " & id & " AND lang_id = " & rsLang("language_id")
		if not rsOrderstatus.eof then
			order_status = rsOrderstatus("order_status")
			unique_id    = rsOrderstatus("order_status_unique_id")
			isnew        = 0
		else
			order_status = ""
			isnew        = 1
			unique_id    = 0
		end if
	%>
    <tr>
      <td>&nbsp;<%=rsLang("language_name")%></td>
      <td>
		<input name="orderstatus_<%=x%>" type="text" id="orderstatus_<%=x%>" value="<%=order_status%>">
		<input name="lang_id_<%=x%>" type="hidden" value="<%=rsLang("language_id")%>" />
		<input type="hidden" name="isnew_<%=x%>" value="<%=isnew%>" />
		<input type="hidden" name="unique_<%=x%>" value="<%=unique_id%>" />
	  </td>
    </tr>
	<%
		rsLang.movenext
	loop
	%>
	<tr align="center"> 
      <td colspan="2">
	    <input type="hidden" name="total_lang" value="<%=x%>" />
	    <%=BuildSubmitter("submit","Update orderstatus", request.querystring("p"))%> 
        <input type="button" name="Cancel" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';">
      </td>
	</tr>
  </table>
</form>
