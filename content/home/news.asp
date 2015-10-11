<p class="pageheader"><%=strNews%></p>
<%
set rsNews = server.createobject("ADODB.recordset")
rsNews.cursortype = 3

strSQL = "SELECT TOP 5 news_id, news_title, news_date_added FROM news WHERE news_lang_id = " & session("language_id") & " ORDER BY news_id DESC;"
rsNews.open strSQL, adoCon

if not rsNews.eof then
	for x = 1 to 5
		if rsNews.eof then exit for
		
		news_date_added = rsNews("news_date_added")
		news_id         = rsNews("news_id")
		news_title      = rsNews("news_title")
%>
<p>
  <%=news_date_added%><br />
  <img src="includes/styles/<%=strStylesheet%>/before_news.gif" width="10" height="10" align="middle" /> 
  <b><a href="?mod=news&amp;news_id=<%=news_id%>"><%=news_title%></a></b>
</p>
	<%
		rsNews.movenext
	next
else
	response.write("&nbsp;")
end if
rsNews.close
set rsNews = nothing
%>

