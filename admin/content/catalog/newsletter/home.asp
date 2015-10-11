<%
if len(request.form()) > 0 then
	total_lang = request.form("total_lang")
	lang_init  = ""
	lang_count = 0
	
	for lang_count = 1 to total_lang
		lang_id = request.form("lang_id_" & lang_count)
		subject = Replace(request.form("subject_" & lang_count), "'", "''")
		content = Replace(request.form("content_" & lang_count), "'", "''")
		
		if len(subject) > 0 and len(content) > 0 then
			session("subject_" & lang_id) = subject
			session("content_" & lang_id) = content
			if len(lang_init) > 0 then
				lang_init = lang_init & ";" & lang_id
			else
				lang_init = lang_id
			end if
		end if
	next
	session("lang_init") = lang_init
	response.redirect("?p=" & request.querystring("p") & "&action=send")
end if
set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_id, language_name, language_default FROM lang"
rsLang.open strSQL, adoCon

total_lang = rsLang.recordcount

set rsNewsletter = server.createobject("ADODB.recordset")
rsNewsletter.cursortype = 3

strSQL = "SELECT user_email, user_lang_id FROM newsletter"
rsNewsletter.open strSQL, adoCon
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
<table width="700" align="center" cellpadding="2" cellspacing="2">
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
				
				rsNewsletter.filter = "user_lang_id = " & rsLang("language_id")
				subscriptions = rsNewsletter.recordcount
				%>
				<div id="lang_<%=x%>" style="visibility: hidden; position: absolute;">
				  <input type="hidden" name="lang_id_<%=x%>" value="<%=rsLang("language_id")%>" />
				  <table width="100%" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
				    <tr>
					  <td colspan="2" bgcolor="#666666"><font color="#FFFFFF"><b>Newsletter (<%=rsLang("language_name")%>)</b></font></td>
					</tr>
					<tr>
					  <td width="120">&nbsp;Total subscriptions:</td>
					  <td><%=subscriptions%></td>
					</tr>
					<tr>					  
					  <td>&nbsp;Subject:</td>
					  <td width="40">&nbsp;<input name="subject_<%=x%>" type="text" value="" size="40" /></td>
					</tr>
					<tr>
					  <td colspan="2" align="center">
				      <%
       					Dim oFCKeditor
					   	Set oFCKeditor = New FCKeditor
					   	oFCKeditor.BasePath = strVirtualPath & "admin/FCKeditor/"
					   	oFCKeditor.Height = "500"
					   	oFCKeditor.Create "content_" & x
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
  <tr>
    <td align="center">
	  <input type="hidden" name="total_lang" value="<%=total_lang%>" />
	  <%=buildSubmitter("cmdSend", "Send newsletter", request.querystring("p"))%>
	</td>
  </tr>
  <tr>
    <td><font size="1">*Newsletter for a language will only be sent when both 
      subject and body have content</font></td>
  </tr>
</table>
<script>
switchLang('<%=default_open%>');
</script>
<%
rsLang.close
set rsLang = nothing

rsNewsletter.close
set rsNewsletter = nothing
%>