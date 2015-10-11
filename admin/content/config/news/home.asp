<script>
function deleteMe(id){
	del = window.confirm("Are you sure you want to delete this newsitem?");
	
	if(del)document.location = "?p=<%=request.querystring("p")%>&action=delete&nid=" + id;
}
</script>
<%
pp = request.querystring("pp")

if len(pp) > 0 and isnumeric(pp) then
	pp = cint(pp)
else
	pp = 1
end if

set rsNews = server.createobject("ADODB.recordset")
rsNews.cursortype = 3

strSQL = "SELECT news_id, news_title, news_date_added FROM news WHERE news_lang_id = " & default_lang_id & " ORDER BY news_id DESC;"
rsNews.open strSQL, adoCon

rsNews.pagesize = 20
news_pages      = rsNews.pagecount

if not rsNews.eof then rsNews.absolutepage = pp
%>
	  <table width="500" align="center" cellpadding="2" cellspacing="2">
	  <%
	  	x = 0
	  	for x = 1 to 20
	  		if rsNews.eof then exit for
			response.write("<tr>" & chr(10))
				response.write("<td>" & rsNews("news_title") & " [" & rsNews("news_date_added") & "]</td>" & chr(10))
				response.write("<td width=""120"">")
				select case intModuleRights
					case 1
						response.write("<a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;nid=" & rsNews("news_id") & """>View</a>")
					case 2
						response.write("<a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;nid=" & rsNews("news_id") & """>Edit</a> | ")
						response.write("<a href=""javascript:deleteMe('" & rsNews("news_id") & "');"">Delete</a>")
					case else
						response.write("&nbsp;")
				end select
				response.write("</td>" & chr(10)) 
			response.write("</tr>")
			rsNews.movenext
		next
		response.write("<tr>" & chr(10))
			response.write("<td colspan=""2"">Pages:")	
			x = 0
			for x = 1 to news_pages
				response.write("&nbsp;<a href=""?p=" & request.querystring("p") & "&amp;pp=" & x & """>" & x & "</a>")
			next
		response.write("</tr>" & chr(10))
	  %>
            <tr align="center"> 
              <td colspan="2">
			  <% if intModuleRights = 2 then %>
                <a href="?p=<%=request.querystring("p")%>&amp;action=add">add newsitem</a> 
                <% end if %>
			  </td>
            </tr>
          </table>
<%
rsNews.close
set rsNews = nothing
%>