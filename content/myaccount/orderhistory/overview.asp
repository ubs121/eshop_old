<%
pp = request.querystring("pp")
if len(pp) > 0 and isnumeric(pp) then
	pp = cint(pp)
else
	pp = 1
end if

set rsStatus = server.createobject("ADODB.recordset")
rsStatus.cursortype = 3

strSQL = "SELECT order_status_ID, order_status FROM order_status WHERE lang_ID = " & session("language_id")
rsStatus.open strSQL, adoCon

set rsOrders = server.createobject("ADODB.recordset")
rsOrders.cursortype = 3

strSQL = "SELECT order_ID, total_price, date_ordered, date_confirmed, confirmed, status FROM orders WHERE user_ID = " & session("customer_id") & " ORDER BY order_ID DESC;"
rsOrders.open strSQL, adoCon

rsOrders.pageSize = strProductsPerPage
pages = rsOrders.pagecount

if not rsOrders.eof then rsOrders.absolutepage = pp
%>
<p><b><%=strOrderHistory%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="content" width="60">&nbsp;<b><%=strOrderNumber%></b></td>
		  <td class="content">&nbsp;<b><%=strDateOrdered%></b></td>
          <td class="content">&nbsp;<b><%=strDateConfirmed%></b></td>
		  <td class="content">&nbsp;<b><%=strStatus%></b></td>
		  <td class="content" align="right">&nbsp;<b><%=strTotal%></b></td>
		  <td class="content" width="100">&nbsp;</td>
        </tr>
<% if rsOrders.eof then %>
		<tr>
		  <td colspan="6">&nbsp;<%=strNoOrderHistory%></td>
		</tr>
<% else %>
<%
x = 0
for x = 1 to strProductsPerPage
	if rsOrders.eof then exit for
	
	if cint(rsOrders("confirmed")) = -1 then
		date_confirmed = rsOrders("date_confirmed")
	else
		date_confirmed = "<b><font color=""#FF0000"">" & strNotConfirmed & "</font></b>"
	end if
	
	rsStatus.filter = "order_status_id = " & rsOrders("status")
	if not rsStatus.eof then
		order_status = rsStatus("order_status")
	else
		order_status = strUnknown
	end if
%>
        <tr onmouseover="this.style.background='#EEEEEE';" onmouseout="this.style.background='';">
          <td class="content">&nbsp;<b><%=rsOrders("order_ID")%></b></td>
          <td class="content">&nbsp;<%=rsOrders("date_ordered")%></td>
          <td class="content">&nbsp;<%=date_confirmed%></td>
          <td class="content">&nbsp;<%=order_status%></td>
          <td class="content" align="right"><%=strCurrency%>&nbsp;<%=RoundNumber(rsOrders("total_price"))%></td>
          <td class="content" align="center"><a href="?mod=myaccount&amp;sub=orders_history&amp;action=view&amp;oid=<%=rsOrders("order_ID")%>"><img src="languages/<%=session("language")%>/images/details.gif" width="100" height="20" border="0" align="absmiddle" /></a></td>
        </tr>
<%
	rsOrders.movenext
next
%>
<% end if %>
      </table></td>
  </tr>
</table>
<p align="right">
<%
if pp > 1 then
	response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & pp - 1 & """>[&lt;&lt;" & strPrevious & "]</a>&nbsp;")
	if pp >= 5 then
		x = 0
		for x = 1 to 2
			response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & x & """>" & x & "</a>&nbsp;")
		next
		response.write("...&nbsp;")
		response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & x & """>" & pp - 1 & "</a>&nbsp;")
	else
		x = 0
		for x = 1 to pp - 1
			response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & x & """>" & x & "</a>&nbsp;")
		next
	end if
end if

response.write("<b>" & pp & "</b>&nbsp;")

if pages > pp then
	if pages >= pp + 5 then
		x = 0
		for x = pp + 1 to pp + 4
			response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & x & """>" & x & "</a>&nbsp;")
		next
		response.write("...&nbsp;")
		response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & pages & """>" & pages & "</a>&nbsp;")
	else
		x = 0
		for x = pp + 1 to pages
			response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & x & """>" & x & "</a>&nbsp;")
		next
	end if
	response.write("<a href=""?mod=myaccount&amp;sub=orders_history&amp;pp=" & pp + 1 & """>[" & strNext & "&gt;&gt;]</a>&nbsp;")
end if
%>
</p>
<%
rsOrders.close
set rsOrders = nothing

rsStatus.close
set rsStatus = nothing
%>