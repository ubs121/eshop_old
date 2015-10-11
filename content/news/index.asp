<% if cint(strShowNews) = 0 then %>
<p class="pageheader"><%=strNews%></p>
<b><%=strNewsModuleDisabled%></b>
<% else %>
<%
nid = request.querystring("news_id")

if isnumeric(nid) and len(nid) > 0 then
	nid = cint(nid)
else
	response.redirect("?mod=home")
end if

'Get the specific article
set rsNewsItem = server.createobject("ADODB.recordset")
rsNewsItem.cursortype = 3

strSQL = "SELECT news_title, news_content, news_date_added FROM news WHERE news_id = " & nid & " AND news_lang_id = " & session("language_id")
rsNewsitem.open strSQL, adoCon

'Get the top 5 latest newsarticles
set rsNews = server.createobject("ADODB.recordset")
rsNews.cursortype = 3

strSQL = "SELECT TOP 5 news_id, news_title, news_date_added FROM news WHERE news_lang_id = " & session("language_id") & " ORDER BY news_id DESC;"
rsNews.open strSQL, adoCon
%>
<table width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td align="left" valign="top" class="home_toprow">
	<% if rsNewsitem.eof then %>
	<p class="pageheader"><%=strNews%></p>
	<b><%=strArticleNotAvailable%></b>
	<% else %>
	<p class="pageheader"><%=rsNewsItem("news_title")%></p>
	<%=rsNewsItem("news_content")%>
	<div id="newsdate"><%=rsNewsItem("news_date_added")%></div>
	<% end if %>
	</td>
	<td class="home_spacer">&nbsp;</td>
	<td width="300" align="left" valign="top" class="home_toprow">
	<p class="pageheader"><%=strNews%></p>
	<%
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
	%>
	</td>
  </tr>
<%
rsNewsItem.close
set rsNewsItem = nothing

rsNews.close
set rsNews = nothing
%>
<% end if %>