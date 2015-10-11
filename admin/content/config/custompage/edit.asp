<%
pid = request.querystring("pid")
if len(pid) = 0 or not isnumeric(pid) then response.redirect("?p=" & request.querystring("p"))

if len(request.form()) > 0 then
	total_pages = request.form("total_lang")
	x = 1
	
	for x = 1 to total_pages
		cp_uid = cint(request.form("cp_uid_" & x))
		lang_id = request.form("lang_id_" & x)
		page_name = makeChars(Replace(request.form("page_name_" & x), "'", "''"))
		page_content = makeChars(Replace(request.form("page_content_" & x), "'", "''"))
		
		if len(page_name) > 0 then
			if cp_uid = 0 then
				strSQL = "INSERT INTO custom_pages (page_id, page_lang_Id, page_name, page_content) VALUES("
				strSQL = strSQL & pid & "," & lang_id & ",'" & page_name & "','" & page_content & "');"
			else
				strSQL = "UPDATE custom_pages SET page_name = '" & page_name & "', page_content = '" & page_content & "' WHERE page_unique_id = " & cp_uid
			end if
			adoCon.execute(strSQL)
		end if
	next
end if

set rsCpage = server.createobject("ADODB.recordset")
rsCpage.cursortype = 3

strSQL = "SELECT page_unique_id, page_lang_id, page_content, page_name FROM custom_pages WHERE page_id = " & pid
rsCpage.open strSQL, adoCon

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_id, language_name, language_default FROM lang WHERE language_show = -1 ORDER BY language_name ASC;"
rsLang.open strSQL, adoCon

total_lang = rsLang.recordcount
%>
<script>
function hideAll(){
	total_lang = <%=total_lang%>;
	for(x=1; x <= total_lang; x++){
		div = document.getElementById("lang_" + x);
		div.style.visibility = "hidden";
		div.style.position   = "absolute";
	}
}

function switchLang(id){
	hideAll();
	div = document.getElementById("lang_" + id);
	div.style.visibility = "";
	div.style.position   = "";
}
</script>
<form name="form1" method="post" action="">
  <table width="600" align="center" cellpadding="2" cellspacing="2">
    <% if len(strError) > 0 then %>
    <tr> 
      <td><b><%=strError%></b></td>
    </tr>
    <% end if %>
    <tr> 
      <td>
	  	<table width="100%" cellpadding="2" cellspacing="1">
          <tr>
		  	<%
			x = 0
		  	do while not rsLang.eof
				x = x + 1
		  		if cint(rsLang("language_default")) = -1 then
					default_open = x
				end if
				response.write("<td align=""center"" bgcolor=""#EEEEEE""><b><a href=""javascript:switchLang('" & x & "');"">" & rsLang("language_name") & "</a></b></td>" & chr(10))
		  		rsLang.movenext
			loop
		  	%>
		  </tr>
		  <tr>
		    <td colspan="<%=total_lang%>">
			<%
			if not rsLang.bof then rsLang.movefirst
			
			x = 0
			do while not rsLang.eof
				x = x + 1
				
				rsCpage.filter = "page_lang_id = " & rsLang("language_id")
				if not rsCpage.eof then
					cp_uid = rsCpage("page_unique_id")
					cp_name = rsCpage("page_name")
					cp_content = rsCpage("page_content")
				else
					cp_uid = 0
					cp_name = ""
					cp_content = ""
				end if
				%>
				<div id="lang_<%=x%>" style="visibility: hidden; position: absolute;">
				<input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
				<input type="hidden" name="cp_uid_<%=x%>" value="<%=cp_uid%>" />
                <table width="700" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
                  <tr>
                    <td colspan="2" bgcolor="#666666"><font color="#FFFFFF"><b>Edit custompage (<%=rsLang("language_name")%>)</b></font></td>
						</tr>
						<tr>
							<td width="120">&nbsp;Page name:</td>
							<td><input name="page_name_<%=x%>" type="text" value="<%=cp_name%>" size="40" /></td>
						</tr>
						<tr>
						  <td colspan="2" align="center">
						            <%
								  	oFCKeditor.Value = cp_content
								   	oFCKeditor.Create "page_content_" & x
								   %>
						  </td>
						</tr>
					</table>
				</div>
				<%
				rsLang.movenext
			loop
			%>
			</td>
		  </tr>
		</table>
	  </td>
    </tr>
    <tr align="center"> 
      <td>
        <%=BuildSubmitter("submit","Update page", request.querystring("p"))%> <input type="button" name="Cancel" value="Back" onclick="document.location='?p=<%=request.querystring("p")%>';"></td>
    </tr>
  </table>
<input type="hidden" name="total_lang" value="<%=total_lang%>" />
</form>
<script>
switchLang('<%=default_open%>');
</script>

<%
rsLang.close
set rsLang = nothing
%>