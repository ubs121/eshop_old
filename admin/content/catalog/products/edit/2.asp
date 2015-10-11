<%
if len(request.form()) > 0 then
	old_lang_id = request.form("old_lang_id")
	lang_id     = cint(request.form("lang_id"))
	pid         = request.form("pid")
	action      = request.form("action")
	
	select case action
		case "change_lang":
			prod_descr = request.form("product_descr")
			prod_descr = Replace(prod_descr, "'", "''")
			prod_descr = makeChars(Replace(prod_descr, chr(10), "<br />"))
			pdid       = request.form("pdid")
			new_descr  = request.form("prod_new")
			
			if new_descr = 1 then
				strSQL = "INSERT INTO product_description (product_id, product_lang_id, product_description) VALUES("
				strSQL = strSQL & pid & "," & old_lang_id & ",'" & prod_descr & "');"
			else
				strSQL = "UPDATE product_description set product_description = '" & prod_descr & "' WHERE product_description_id = " & pdid
			end if
			adoCon.execute(strSQL)
		case "add_spec":
			spec_name  = makeChars(Replace(request.form("spec_name"), "'", "''"))
			spec_value = makeChars(Replace(request.form("spec_value"), "'", "''"))
			
			if len(spec_name) > 0 AND len(spec_value) > 0 then
				strSQL = "INSERT INTO product_info (product_id, product_lang_id, product_info_name, product_info_description) VALUES("
				strSQL = strSQL & pid & "," & lang_id & ",'" & spec_name & "','" & spec_value & "');"
				adoCon.execute(strSQL)
			end if
		case "complete":
			prod_descr = request.form("product_descr")
			prod_descr = Replace(prod_descr, "'", "''")
			prod_descr = makeChars(Replace(prod_descr, chr(10), "<br />"))
			pdid       = request.form("pdid")
			new_descr  = request.form("prod_new")
			
			if new_descr = 1 then
				strSQL = "INSERT INTO product_description (product_id, product_lang_id, product_description) VALUES("
				strSQL = strSQL & pid & "," & old_lang_id & ",'" & prod_descr & "');"
			else
				strSQL = "UPDATE product_description set product_description = '" & prod_descr & "' WHERE product_description_id = " & pdid
			end if
			adoCon.execute(strSQL)
			
			response.redirect("?p=" & request.querystring("p"))
		case "delete_spec":
			spec_id = request.form("spec_id")
			strSQL = "DELETE * FROM product_info WHERE product_info_id = " & spec_id
			adoCon.execute(strSQL)
	end select
else
	new_record = 1
	lang_id    = cint(default_lang_id)
	pid        = request.querystring("pid")
end if

set rsDescr = server.createobject("ADODB.recordset")
rsDescr.cursortype = 3

strSQL = "SELECT product_description_id, product_description FROM product_description WHERE product_id = " & pid & " AND product_lang_id = " & lang_id
rsDescr.open strSQL, adoCon

if not rsDescr.eof then
	pdid = rsDescr("product_description_id")
	prod_descr = rsDescr("product_description")
	prod_descr = Replace(prod_descr, "<br />", chr(10))
	prod_new   = 0
else
	pdid       = 0
	prod_descr = ""
	prod_new   = 1
end if

rsDescr.close
set rsDescr = nothing
%>
<script>
<!--
function setAction(action){
	document.getElementById("action").value = action;
	document.frmUpdateProduct.submit();
}

function deleteSpec(spec_id){
	document.getElementById("action").value = "delete_spec";
	document.getElementById("spec_id").value = spec_id;
	document.frmUpdateProduct.submit();
}
-->
</script>
 <form name="frmUpdateProduct" method="post" action="">
  <table width="500" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <td colspan="2">
	    <b>Step 2 - Language specific</b></td>
    </tr>
    <tr align="right"> 
      <td colspan="2"> Language&nbsp; <select name="lang_id" onChange="javascript:setAction('change_lang');">
          <%
			set rsLang = server.createobject("ADODB.recordset")
			rsLang.cursortype = 3
			
			strSQL = "SELECT language_id, language_name FROM lang"
			rsLang.open strSQL, adoCon
			
			do while not rsLang.eof
				response.write("<option value=""" & rsLang("language_id") & """")
				if cint(rsLang("language_id")) = lang_id then
					response.write(" selected=""selected""")
				end if
				response.write(">" & rsLang("language_name") & "</option>" & chr(13))
				rsLang.movenext
			loop
			%>
        </select> <hr /> </td>
    </tr>
    <tr> 
      <td colspan="2"><b>Description:</b></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><textarea name="product_descr" cols="80" rows="8" id="textarea"><%=prod_descr%></textarea>
        <br><input type="button" name="update_descr" value="Update" onclick="javascript:setAction('change_lang');" />
      </td>
    </tr>
    <tr> 
      <td colspan="2"> <hr /> <b>Specifications</b> </td>
    </tr>
    <%
  set rsSpec = server.createobject("ADODB.recordset")
  rsSpec.cursortype = 3
  
  strSQL = "SELECT product_info_id, product_info_name, product_info_description FROM product_info WHERE product_lang_id = " & lang_id & " AND product_id = " & pid
  rsSpec.open strSQL, adoCon
  
  do while not rsSpec.eof
  %>
    <tr> 
      <td width="120" bgcolor="#EEEEEE"><%=rsSpec("product_info_name")%> (<a href="javascript:deleteSpec('<%=rsSpec("product_info_id")%>');">Delete</a>)</td>
      <td><%=rsSpec("product_info_description")%></td>
    </tr>
    <%
  	rsSpec.movenext
  loop
  rsSpec.close
  set rsSpec = nothing
  %>
    <tr> 
      <td>Specification name:</td>
      <td><input name="spec_name" type="text" id="spec_name3"></td>
    </tr>
    <tr> 
      <td>Specification value:</td>
      <td><input name="spec_value" type="text" id="spec_value2"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input name="add_spec" type="button" id="add_spec" value="Add specification" onClick="javascript:setAction('add_spec');" /></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><hr />
	    <%=BuildSubmitter("cmdSubmit","Complete", request.querystring("p"))%>
	  </td>
    </tr>
  </table>
	    <input type="hidden" name="action" value="complete" id="action" />
		<input type="hidden" name="old_lang_id" value="<%=lang_id%>" />
		<input type="hidden" name="pid" value="<%=pid%>" />
		<input type="hidden" name="pdid" value="<%=pdid%>" />
		<input type="hidden" name="prod_new" id="prod_new" value="<%=prod_new%>" />
		<input type="hidden" name="spec_id" id="spec_id" value="" />
	  </form>