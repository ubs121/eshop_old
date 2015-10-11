<%
if len(request.form()) > 0 then
	man_name = makeChars(Replace(request.form("man_name"), "'", "''"))
	man_url  = request.form("man_url")
	
	if len(man_url) = 12 then
		man_url = ""
	end if
	
	if len(man_name) > 0 then
		strSQL = "INSERT INTO manufacturer (manufacturer_name, manufacturer_url) VALUES('"
		strSQL = strSQL & man_name & "','" & man_url & "');"
		
		adoCon.execute(strSQL)
		response.redirect("?p=" & request.querystring("p"))
	end if
end if
%>
<form name="form1" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td width="120">Manufacturer:</td>
      <td><input name="man_name" type="text" id="man_name" size="40"></td>
    </tr>
    <tr> 
      <td>Website:</td>
      <td><input name="man_url" type="text" id="man_url" value="http://" size="40"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Add manufacturer", request.querystring("p"))%>
<input type="button" name="Cancel" value="Cancel" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
</form>
