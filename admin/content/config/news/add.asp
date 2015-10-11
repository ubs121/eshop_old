<%
set rsLang = server.createobject("ADODB.recordset")
rsLang.cursortype = 3

strSQL = "SELECT language_name, language_id, language_default FROM lang WHERE language_show = -1"
rsLang.open strSQL, adoCon

total_lang = rsLang.recordcount

if len(request.form()) > 0 then
	x = 0
	isAdded = false
	redim newsTitles(total_lang)
	redim newsBodies(total_lang)
	redim Lang_id(total_lang)
	
	for x = 1 to total_lang
		newsTitles(x) = makeChars(Replace(request.form("title_" & x), "'", "''"))
		newsBodies(x) = makeChars(Replace(request.form("body_" & x), "'", "''"))
		Lang_id(x)    = request.form("lang_" & x)
	next
	
	x = 0
	set rsNews = server.createobject("ADODB.recordset")
	strSQL = "SELECT TOP 1 * FROM news ORDER BY news_id DESC;"
	
	rsNews.open strSQL, adoCon, 2, 2
	
	if not rsNews.eof then
		next_news_id = rsNews("news_id") + 1
	else
		next_news_id = 1
	end if
	
	for x = 1 to total_lang
		if len(newsTitles(x)) > 0 then
			strSQL = "INSERT INTO news (news_id, news_lang_id, news_title, news_content, news_date_added) VALUES("
			strSQL = strSQL & next_news_id & "," & Lang_id(x) & ",'" & newsTitles(x) & "','" & newsBodies(x) & "', now());"
			adoCon.execute(strSQL)
			isAdded = true
		end if
	next
	if isAdded then response.redirect("?p=" & request.querystring("p") & "&action=edit&nid=" & next_news_id)
end if
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
  	response.write("<div id=""news_" & x & """ style=""visibility: hidden; position: absolute;"">" & chr(10))
	%>
	    <table width="100%" cellspacing="2" cellpadding="4" style="border: solid 1px #000000;">
          <tr> 
            <td colspan="2" bgcolor="#666666"><b><font color="#FFFFFF">Add news 
              (<%=rsLang("language_name")%>)</font></b></td>
          </tr>
          <tr> 
            <td width="80"><strong>News title:</strong></td>
            <td> <input name="title_<%=x%>" type="text" id="title_<%=x%>" value="<%=request.form("title_" & x)%>" size="60"> 
              <input name="lang_<%=x%>" type="hidden" id="lang_<%=x%>" value="<%=rsLang("language_id")%>"></td>
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
	  <%=BuildSubmitter("cmdSubmit","Add newsitem", request.querystring("p"))%>&nbsp;
	  <input type="button" name="Cancel" value="Cancel" onClick="document.location='?p=<%=request.querystring("p")%>';"> 
      </td>
</tr>
</table>
</form>
<script>
switchLang('<%=default_open%>');
</script>