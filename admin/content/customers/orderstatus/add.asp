<%
if len(request.form()) > 0 then
	order_default = request.form("chkDefault")
	next_status   = request.form("next_status")
	total_lang    = request.form("total_lang")
	new_id        = cint(request.form("new_id"))
	
	if isnumeric(next_status) then
		next_status = cint(next_status)
	else
		next_status = 0
	end if
	
	if order_default = "yes" then
		order_default = -1
		strSQL = "UPDATE order_status SET order_default = 0 WHERE order_default = -1"
		adoCon.execute(strSQL)
	else
		order_default = 0
	end if
	
	x = 0
	for x = 1 to total_lang
		lang_id = request.form("lang_id_" & x)
		order_status = request.form("orderstatus_" & x)
		
		strSQL = "INSERT INTO order_status (order_status_id, order_status, lang_id, order_default, next_status) VALUES("
		strSQL = strSQL & new_id & ",'" & order_status & "'," & lang_id & "," & order_default & "," & next_status & ");"
		adoCon.execute(strSQL)
	next
	response.redirect("?p=" & request.querystring("p"))
end if
set rsOrderstatus = server.createobject("ADODB.recordset")
rsOrderstatus.cursortype = 3

strSQL = "SELECT order_status_id, order_status FROM order_status WHERE lang_id = " & default_lang_id
rsOrderstatus.open strSQL, adoCon

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_id, language_name FROM lang"
rsLang.open strSQL, adoCon
%>
<form name="form1" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="100">Default status:</td>
      <td><input name="chkDefault" type="checkbox" id="chkDefault" value="yes"></td>
    </tr>
    <tr> 
      <td>Next status :</td>
      <td> <select name="next_status">
          <option value="0">None</option>
          <%
		  do while not rsOrderstatus.eof
		  	response.write("<option value=""" & rsOrderstatus("order_status_id") & """>" & rsOrderstatus("order_status") & "</option>" & chr(10))
			order_high = rsOrderstatus("order_status_id")
		  	rsOrderstatus.movenext
		  loop
		  %>
        </select> </td>
    </tr>
    <tr> 
      <td colspan="2" height="10"></td>
    </tr>
    <%
	x = 0
	do while not rsLang.eof
		x = x + 1
	%>
    <tr> 
      <td>&nbsp;<%=rsLang("language_name")%></td>
      <td>
	    <input name="orderstatus_<%=x%>" type="text" id="orderstatus_<%=x%>" value="<%=order_status%>"> 
        <input name="lang_id_<%=x%>" type="hidden" value="<%=rsLang("language_id")%>" /> 
      </td>
    </tr>
    <%
		rsLang.movenext
	loop
	%>
    <tr align="center"> 
      <td colspan="2">
	    <input type="hidden" name="new_id" value="<%=order_high + 1%>" />
	    <input type="hidden" name="total_lang" value="<%=x%>" /> 
        <%=BuildSubmitter("submit","Add orderstatus", request.querystring("p"))%> 
        <input type="button" name="Cancel" value="Back" onClick="document.location='?p=<%=request.querystring("p")%>';"> 
      </td>
    </tr>
  </table>
</form>
<%
rsLang.close
set rsLang = nothing

rsOrderstatus.close
set rsOrderstatus = nothing
%>