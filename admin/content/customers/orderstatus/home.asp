<%
if len(request.form()) > 0 then
	errorID = 0
	total_orders = cint(request.form("total_orders"))
	redim order_status(total_orders, 2)
	redim order_status_ordered(total_orders, 2)
	x = 0
	for x = 1 to total_orders
		order_status(x, 1) = request.form("lsOrder_" & x)
		order_status(x, 2) = request.form("order_status_id_" & x)
	next
	x = 0
	for x = 1 to total_orders
		temp_order = cint(order_status(x, 1))
		temp_order_id = order_status(x, 2)
		
		y = 0
		for y = 1 to total_orders
			if y = temp_order then
				order_status_ordered(y, 1) = temp_order
				order_status_ordered(y, 2) = temp_order_id
				exit for
			end if
		next
	next
	
	x = 0
	for x = 1 to total_orders
		if len(order_status_ordered(x, 1)) = 0 then
			errorId = 1
			response.write("Error: Duplicate found")
			exit for
		end if
	next
	
	if errorId = 0 then
	x = 0
	redim SQL(total_orders)
		set rsOrderstatus = server.createobject("ADODB.recordset")
		rsOrderstatus.cursortype = 3
		
		for x = 1 to total_orders
			temp_sql = ""
			strSQL = "SELECT order_status_unique_id FROM order_status WHERE order_status_id = " & order_status_ordered(x, 2)
			rsOrderstatus.open strSQL, adoCon
			do while not rsOrderstatus.eof
				if len(temp_sql) = 0 then
					temp_sql = " WHERE order_status_unique_id = " & rsOrderstatus("order_status_unique_id")
				else
					temp_sql = temp_sql & " OR order_status_unique_id = " & rsOrderstatus("order_status_unique_id")
				end if
				rsOrderstatus.movenext
			loop
			temp_sql = "UPDATE order_status SET order_status_id = " & order_status_ordered(x, 1) & temp_sql
			SQL(x) = temp_sql
			rsOrderstatus.close
		next
		set rsOrderstatus = nothing
		
		x = 0
		for x = 1 to total_orders
			adoCon.execute(SQL(x))
		next
	end if
end if
set rsOrderstatus = server.createobject("ADODB.recordset")
rsOrderstatus.cursortype = 3

strSQL = "SELECT order_status_id, order_status FROM order_status WHERE lang_id = " & default_lang_id & " ORDER BY order_status_id ASC;"
rsOrderstatus.open strSQL, adoCon

totalOrders = rsOrderstatus.recordcount
%>
<form name="frmOrderstatus" method="post" action="">
<table width="500" align="center" cellpadding="2" cellspacing="2">
<%
bgColor = "#EEEEEE"
x = 0
do while not rsOrderstatus.eof
	x = x + 1
	if bgColor = "#EEEEEE" then
		bgcolor = "#FFFFFF"
	else
		bgcolor = "#EEEEEE"
	end if
%>
  <tr bgcolor="<%=bgcolor%>">
    <td width="20">
	  <select name="lsOrder_<%=x%>">
	    <%
		y = 0
		for y = 1 to totalOrders
			response.write("<option value=""" & y & """")
			if y = cint(rsOrderstatus("order_status_id")) then
				response.write(" selected=""selected""")
			end if
			response.write(">" & y & "</option>" & chr(10))
		next
		%>
	  </select>
	  <input type="hidden" value="<%=rsOrderstatus("order_status_id")%>" name="order_status_id_<%=x%>" />
	</td>
	<td>&nbsp;<%=rsOrderstatus("order_status")%></td>
	<td width="80" align="center">
<% if intModuleRights = 2 then %><a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;id=<%=rsOrderstatus("order_status_id")%>">edit</a><% else %>&nbsp;
<% end if %></td>
	<td width="80" align="center">
<% if intModuleRights = 2 then %><a href="?p=<%=request.querystring("p")%>&amp;action=delete&amp;id=<%=rsOrderstatus("order_status_id")%>">delete</a><% else %>&nbsp;<% end if %>
    </td>
  </tr>
<%
	rsOrderstatus.movenext
loop
%>
    <tr align="right" bgcolor="#999999"> 
      <td colspan="4" style="border: solid 1px #000000;"><input type="button" name="btnAddOrderstatus" value="Add orderstatus" onclick="javascript:document.location='?p=<%=request.querystring("p")%>&action=add';" />&nbsp;<input type="submit" name="Submit" value="Update order"></td>
  </tr>
</table>
<input type="hidden" name="total_orders" value="<%=x%>" />
</form>
<%
rsOrderstatus.close
set rsOrderstatus = nothing
%>