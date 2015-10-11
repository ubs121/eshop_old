<%
nid = request.querystring("nid")

if len(nid) > 0 and isnumeric(nid) then
	nid = cint(nid)
else
	response.redirect("?p=" & request.querystring("p"))
end if

set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_name, language_id, language_default FROM lang WHERE language_show = -1"
rsLang.open strSQL, adoCon

total_lang = rsLang.recordcount

if len(request.form()) > 0 then
	x = 0
	for x = 1 to total_lang
		nuid = cint(request.form("nuid_" & x))
		newsTitle = makeChars(Replace(request.form("title_" & x), "'", "''"))
		newsBody  = Replace(request.form("body_" & x), "'", "''")
		lang_id   = request.form("lang_" & x)
		
		if nuid = 0 then
			if len(newsTitle) > 0 then
				strSQL = "INSERT INTO news (news_id, news_lang_id, news_title, news_content, news_date_added) VALUES("
				strSQL = strSQL & nid & "," & lang_id & ",'" & newsTitle & "','" & newsBody & "', now());"
			else
				strSQL = ""
			end if
		else
			if len(newsTitle) > 0 then
				strSQL = "UPDATE news SET news_title = '" & newsTitle & "', news_content = '" & newsBody & "' WHERE news_unique_id = " & nuid
			else
				strSQL = ""
			end if
		end if
		
		if len(strSQL) > 0 then
			adoCon.execute(strSQL)
		end if
	next
end if

set rsNews = server.createobject("ADODB.recordset")
rsNews.cursortype = 3

strSQL = "SELECT news_unique_id, news_title, news_content, news_lang_id FROM news WHERE news_id = " & nid
rsNews.open strSQL, adoCon
%>
<script>
var total_lang = <%=total_lang%>;

function hideAll(){
	total_lang = <%=total_lang%>;
	for(x=1; x <= total_lang; x++){
		div = document.getElementById("news_" + x);
		div.style.visibility = "hidden";
		div.style.position   = "absolute";
	}
}

function switchLang(id){
	hideAll();
	div = document.getElementById("news_" + id);
	div.style.visibility = "";
	div.style.position   = "";
}
</script>
<form name="frmAddNews" action="" method="post">
<table width="500" align="center" cellpadding="2" cellspacing="2">
  <tr>
<%
x = 0
do while not rsLang.eof
	x = x + 1
	response.write("<td align=""center"" bgcolor=""#EEEEEE""><a href=""javascript:switchLang('" & x & "');"">" & rsLang("language_name") & "</a></td>" & chr(10))
	rsLang.movenext
loop
response.write("</tr>" & chr(10))
%>
<tr>
  <td colspan="<%=total_lang%>">
  <%
  if not rsLang.bof then rsLang.movefirst
  x = 0
  
  do while not rsLang.eof
  	x = x + 1
	if cint(rsLang("language_id")) = default_lang_id then
		default_open = x
	end if
	
	rsNews.filter = "news_lang_id = " & rsLang("language_id")
	if not rsNews.eof then
		news_uid = rsNews("news_unique_id")
		news_title = rsNews("news_title")
		news_content = rsNews("news_content")
	else
		news_uid = 0
		news_title = ""
		news_content = ""
	end if
	
  	response.write("<div id=""news_" & x & """ style=""visibility: hidden; position: absolute;"">" & chr(10))
	%>
	    <table width="100%" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
          <tr> 
            <td colspan="2" bgcolor="#666666"><b><font color="#FFFFFF">Add news 
              (<%=rsLang("language_name")%>)</font></b></td>
          </tr>
          <tr> 
            <td width="80"><strong>News title:</strong></td>
            <td> <input name="title_<%=x%>" type="text" id="title_<%=x%>" value="<%=news_title%>" size="60" /> 
              <input name="lang_<%=x%>" type="hidden" id="lang_<%=x%>" value="<%=rsLang("language_id")%>" /> 
              <input name="nuid_<%=x%>" type="hidden" id="nuid_<%=x%>" value="<%=news_uid%>" /> 
            </td>
          </tr>
          <tr> 
            <td colspan="2"><strong>News body:</strong></td>
          </tr>
          <tr align="center"> 
            <td colspan="2">
			<%
			oFCKeditor.Value = news_content
			oFCKeditor.Create "body_" & x
			%>
            </td>
          </tr>
        </table>
	<%
	response.write("</div>" & chr(10))
  	rsLang.movenext
  loop
  %>
  </td>
</tr>
<tr>
      <td colspan="<%=total_lang%>" align="center">
	  <%=BuildSubmitter("cmdSubmit","Update newsitem", request.querystring("p"))%>&nbsp;
	  <input type="button" name="btnBack" value="Back" onClick="document.location='?p=<%=request.querystring("p")%>';"> 
      </td>
</tr>
</table>
</form>
<script>
switchLang('<%=default_open%>');
</script>
<%
rsLang.close
set rsLang = nothing

rsNews.close
set rsNews = nothing
%>