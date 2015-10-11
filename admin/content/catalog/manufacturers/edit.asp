<%
man_id = request.querystring("mid")

if len(man_id) = 0 or not isnumeric(man_id) then response.redirect("?p=" & request.querystring("p"))

if len(request.form()) > 0 then
	man_name = makeChars(Replace(request.form("man_name"), "'", "''"))
	man_url  = request.form("man_url")
	
	if len(man_url) < 12 then
		man_url = ""
	end if
	
	if len(man_name) > 0 then
		strSQL = "UPDATE manufacturer SET manufacturer_name = '" & man_name & "', manufacturer_url = '" & man_url & "' WHERE manufacturer_id = " & man_id
		adoCon.execute(strSQL)
		
		response.redirect("?p=" & request.querystring("p"))
	end if
else
	set rsMan = server.createobject("ADODB.recordset")
	rsMan.cursortype = 3
	
	strSQL = "SELECT manufacturer_name, manufacturer_url FROM manufacturer WHERE manufacturer_id = " & man_id
	rsMan.open strSQL, adoCon
	
	if not rsMan.eof then
		man_name = rsMan("manufacturer_name")
		man_url  = rsMan("manufacturer_url")
	end if
	
	rsMan.close
	set rsMan = nothing
end if

if len(man_url) < 12 then
	man_url = "http://"
end if
%>
<form name="form1" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="120">Manufacturer:</td>
      <td><input name="man_name" type="text" id="man_name" value="<%=man_name%>" size="40"></td>
    </tr>
    <tr> 
      <td>Website:</td>
      <td><input name="man_url" type="text" id="man_url" value="<%=man_url%>" size="40"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Update manufacturer", request.querystring("p"))%> 
        <input type="button" name="Cancel" value="Cancel" onClick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
</form>
