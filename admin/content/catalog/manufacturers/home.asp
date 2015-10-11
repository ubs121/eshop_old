<%
pp = request.querystring("pp")

if len(pp) = 0 or not isnumeric(pp) then
	pp = 1
else
	pp = cint(pp)
end if

set rsMans = server.createobject("ADODB.recordset")
rsMans.cursortype = 3

strSQL = "SELECT manufacturer_id, manufacturer_name FROM manufacturer ORDER BY manufacturer_name ASC;"
rsMans.open strSQL, adoCon

rsMans.pagesize = 20
pages           = rsMans.pagecount

if not rsMans.eof then rsMans.absolutepage = pp
%>
<br />
<table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
<%
bgcolor = "#EEEEEE"
for man_looper = 1 to 20
	if rsMans.eof then exit for
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
	response.write("<tr bgcolor=""" & bgcolor & """>" & chr(13))
		response.write("<td>" & rsMans("manufacturer_name") & "</td>" & chr(13))
		if intModuleRights = 2 then
			response.write("<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;mid=" & rsMans("manufacturer_id") & """>Edit</a>")
		else
			response.write("<td width=""60"">&nbsp;</td>")
		end if
		if intModuleRights = 2 then
			response.write("<td width=""60"" align=""center""><a href=""?p=" & request.querystring("p") & "&amp;action=delete&amp;mid=" & rsMans("manufacturer_id") & """>Delete</a>")
		else
			response.write("<td width=""60"">&nbsp;</td>")
		end if
	response.write("</tr>")
	rsMans.movenext
next
%>
</table>
<p>
  Pages:
  	<%
	for page_looper = 1 to pages
		if page_looper = pp then
			response.write("&nbsp;<b>" & pp & "</b>")
		else
			response.write("&nbsp;<a href=""?p=" & request.querystring("p") & "&amp;pp=" & page_looper & """>" & page_looper & "</a>")
		end if
	next
	%>
</p>
<p align="center"> <a href="?p=<%=request.querystring("p")%>&amp;action=add">Add 
  a manufacturer </a></p>
