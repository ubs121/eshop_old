<%
user_type = cint(request.querystring("type"))
order     = request.querystring("order")
pp        = request.querystring("pp")

if len(pp) = 0 or not isnumeric(pp) then
	pp = 1
else
	pp = cint(pp)
end if
%>
<form name="frmCustomNav" id="frmCustomNav" method="get" action="">
  <input type="hidden" name="p" value="<%=request.querystring("p")%>" />
  <input type="hidden" name="pp" value="<%=pp%>" />
  <p align="right">
    <select name="type" disabled="disabled">
	  <option value="0">All customers</option>
	  <option value="">Confirmed customers</option>
	  <option value="">Unconfirmed customers</option>
	</select>
	<select name="order">
	  <option value="">-- Order By --</option>
	  <option value="user_lastname"<% if order="user_lastname" then %> selected="selected"<% end if %>>Name</option>
	  <option value="user_id"<% if order = "user_id" then %> selected="selected"<% end if %>>registration date</option>
	  <option value="user_email"<% if order = "user_email" then %> selected="selected"<% end if %>>Email</option>
	</select>
	<input type="submit" value="GO" name="cmdSubmit" />
	&nbsp;
  </p>
</form>
<%
set rsCustomers = server.createobject("ADODB.recordset")
rsCustomers.cursortype = 3

strSQL = "SELECT user_id, user_lastname, user_firstname, user_email FROM users"

select case order
	case "user_name":
		strSQL = strSQL & " ORDER BY user_lastname, user_firstname ASC;"
	case "user_id":
		strSQL = strSQL & " ORDER BY user_id DESC;"
	case "user_email":
		strSQL = strSQL & " ORDER BY user_email ASC;"
	case else
		strSQL = strSQL & " ORDER BY user_lastname, user_firstname ASC;"
end select

rsCustomers.open strSQL, adoCon

rsCustomers.pagesize = 20
pages = rsCustomers.pagecount

if not rsCustomers.eof then rsCustomers.absolutepage = pp
%>
<table width="500" align="center" cellpadding="2" cellspacing="2">
<%
bgcolor = "#EEEEEE"

for x = 1 to 20
	if rsCustomers.eof then exit for
	if bgcolor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
	
	user_id    = rsCustomers("user_id")
	user_name  = rsCustomers("user_lastname") & " " & rsCustomers("user_firstname")
	user_email = rsCustomers("user_email")
	
	response.write("<tr bgcolor=""" & bgcolor & """>" & chr(10))
		if intModuleRights = 2 then
			response.write("  <td><a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;cid=" & user_id & """>" & user_name & "</a></td>" & chr(10))
		else
			response.write("  <td>" & user_name & "</td>" & chr(10))
		end if
		response.write("  <td width=""150"">")
		if intModuleRights = 2 then
			if len(user_email) > 0 then
				response.write("<a href=""?p=" & request.querystring("p") & "&amp;action=sendmail&amp;cid=" & user_id & """>" & user_email & "</a>")
			else
				response.write("&nbsp;")
			end if
		else
			response.write("&nbsp;")
		end if
		response.write("</td>" & chr(10))
	response.write("</tr>" & chr(10))
	
	rsCustomers.movenext
next
%>
</table>
<p>
Pages: 
<%
for page_looper = 1 to pages
	if page_looper = pp then
		response.write("&nbsp;<b>" & page_looper & "</b>")
	else
		response.write("&nbsp;<a href=""?p=" & request.querystring("p") & "&amp;type=" & user_type & "&amp;order=" & order & "&amp;pp=" & page_looper & """>" & page_looper & "</a>")
	end if
next
%>
</p>
<%
rsCustomers.close
set rsCustomers = nothing
%>