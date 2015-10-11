<%
if len(request.form()) > 0 then
	action        = cint(request.form("action"))
	total_records = cint(request.form("total_records"))
	
	'Delete records
	counter = 0
	if action = -1 then
		for counter = 1 to total_records
			chkOrder = request.form("chkOrder_" & counter)
			order_id = cint(request.form("order_" & counter))
			if chkOrder = "yes" then
				strSQL = "DELETE * FROM orders WHERE order_id = " & order_id
				adoCon.execute(strSQL)
				strSQL = "DELETE * FROM order_info WHERE order_id = " & order_id
				adoCon.execute(strSQL)
			end if
		next
	end if
	
	if action = -2 then
		for counter = 1 to total_records
			chkOrder = request.form("chkOrder_" & counter)
			order_id = cint(request.form("order_" & counter))
			if chkOrder = "yes" then
				strSQL = "UPDATE orders SET paid = 0 WHERE order_ID = " & order_ID
				adoCon.execute(strSQL)
			end if
		next
	end if
	
	if action = -3 then
		for counter = 1 to total_records
			chkOrder = request.form("chkOrder_" & counter)
			order_id = cint(request.form("order_" & counter))
			if chkOrder = "yes" then
				strSQL = "UPDATE orders SET paid = -1 WHERE order_ID = " & order_ID
				adoCon.execute(strSQL)
			end if
		next
	end if
	
	'change status
	counter = 0
	if action > 0 then
		for counter = 1 to total_records
			chkOrder = request.form("chkOrder_" & counter)
			order_id = cint(request.form("order_" & counter))
			
			if chkOrder = "yes" then
				strSQL = "UPDATE orders SET status = " & action & " WHERE order_id = " & order_id
				adoCon.Execute(strSQL)
			end if	
		next
	end if
end if
order_type = request.querystring("order_type")
pp         = request.querystring("pp")
oid        = request.querystring("oid")
stype      = request.querystring("stype")
customer_n = request.querystring("customer_name")
company_n  = request.querystring("company_name")

if len(pp) > 0 and isnumeric(pp) then
	pp = cint(pp)
else
	pp = 1
end if

set rsOrdertype = server.createobject("ADODB.recordset")
rsOrdertype.cursortype = 3

strSQL = "SELECT order_status_id, order_status, order_default FROM order_status WHERE lang_id = " & default_lang_id & " ORDER BY order_status_id ASC;"
rsOrderType.open strSQL, adoCon

if len(order_type) > 0 and isnumeric(order_type) then
	order_type = cint(order_type)
else
	rsOrderType.filter = "order_default = -1"
	if not rsOrderType.eof then
		order_type = cint(rsOrderType("order_status_id"))
	else
		order_type = 0
	end if
	rsOrderType.filter = ""
end if
%>
<script>
function viewSearch(div){
	div = document.getElementById("search_" + div);
	div_stat = div.style.visibility;
	
	document.getElementById("search_oid").style.visibility = "hidden";
	document.getElementById("search_adv").style.visibility = "hidden";
	document.getElementById("search_oid").style.position = "absolute";
	document.getElementById("search_adv").style.position = "absolute";
	
	if(div_stat=="hidden"){
		div.style.visibility = "";
		div.style.position = "";
	}
}
</script>
<form name="frmCustomNav" id="frmCustomNav" method="get" action="">
  <input type="hidden" name="p" value="<%=request.querystring("p")%>" />
  <p align="right">
  	Show: 
    <select name="order_type" onchange="javascript:document.frmCustomNav.submit();">
	  <option value="0">All orders</option>
	  <%
	  do while not rsOrderType.eof
	  	response.write("<option value=""" & rsOrderType("order_status_id") & """")
		if order_type = cint(rsOrderType("order_status_id")) then
			response.write(" selected=""selected""")
		end if
		response.write(">" & rsOrderType("order_status") & "</option>" & chr(10))
	  	rsOrderType.movenext
	  loop
	  
	  rsOrderType.close
	  set rsOrderType = nothing
	  %>
	</select>&nbsp;
  </p>
</form>
<%
set rsOrders = server.createobject("ADODB.recordset")
rsOrders.cursortype = 3

if stype = "oid" and isnumeric(oid) then
	if order_type = 0 then
		strSQL = "SELECT order_id, date_ordered, total_price, user_firstname, user_lastname, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id WHERE order_id LIKE " & oid & " ORDER BY order_id DESC;"
	else
		strSQL = "SELECT order_id, date_ordered, total_price, user_firstname, user_lastname, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id WHERE status = " & order_type & " AND order_id LIKE " & oid & " ORDER BY order_id DESC;"
	end if
elseif stype = "adv" then
	sql = 0
	if order_type > 0 then
		sql = 1
		strWhere = " WHERE status = " & order_type
	end if
	
	if len(customer_n) > 0 then
		if sql = 0 then
			sql = 1
			strWhere = " WHERE orders.user_id IN(select users.user_id FROM users WHERE ([user_lastname] & ' ' & [user_firstname]) LIKE '%" & customer_n & "%')"
		else
			strWhere = strWhere & " AND orders.user_id IN(select users.user_id FROM users WHERE ([user_lastname] & ' ' & [user_firstname]) LIKE '%" & customer_n & "%')"
		end if
	end if
	if len(company_n) > 0 then
		if sql = 0 then
			strWhere = " WHERE address_id IN(SELECT user_address_id FROM user_address WHERE user_company_name LIKE '%" & company_n & "%')"
		else
			strWhere = strWhere & " AND address_id IN(SELECT user_address_id FROM user_address WHERE user_company_name LIKE '%" & company_n & "%')"
		end if
	end if
	strSQL = "SELECT order_id, date_ordered, total_price, user_firstname, user_lastname, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id" & strWhere & " ORDER BY order_id DESC;"
else
	if order_type = 0 then
		strSQL = "SELECT order_id, date_ordered, total_price, user_firstname, user_lastname, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id ORDER BY order_id DESC;"
	else
		strSQL = "SELECT order_id, date_ordered, total_price, user_firstname, user_lastname, Paid FROM orders INNER JOIN users ON orders.user_id = users.user_id WHERE status = " & order_type & " ORDER BY order_id DESC;"
	end if
end if

rsOrders.open strSQL, adoCon

rsOrders.pagesize = 20
pages = rsOrders.pagecount

if not rsOrders.eof then rsOrders.absolutepage = pp
%>
 <table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
	<tr align="left" bgcolor="#999999"> 
      
    <td colspan="5" style="border: solid 1px #000000;"> <a href="javascript:viewSearch('oid');">search 
      order ID</a> | <a href="javascript:viewSearch('adv');">Advanced search</a> 
	  <div id="search_oid" style="visibility:hidden; margin-top: 2px; position: absolute;">
	  <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
		  <td><form name="frmSearchOID" method="get" action="">
                <input name="oid" type="text" id="oid" value="<%=request.querystring("oid")%>" size="6">
                <input name="cmdSearch" type="submit" id="cmdSearch" value="OK">
                <input name="p" type="hidden" id="p" value="<%=request.querystring("p")%>">
                <input name="order_type" type="hidden" id="order_type" value="<%=request.querystring("order_type")%>">
                <input name="stype" type="hidden" id="stype" value="oid">
              </form></td>
		</tr>
	  </table>
	  </div>
	  <div id="search_adv" style="visibility:hidden; margin-top: 2px; position: absolute;">
	  <form name="frmSearchAdv" method="get" action="">
	    <input type="hidden" name="p" value="<%=request.querystring("p")%>" />
		<input type="hidden" name="order_type" value="<%=request.querystring("order_type")%>" />
		<input type="hidden" name="stype" value="adv" />
	    <table width="100%" cellspacing="2" cellpadding="0">
          <tr> 
              <td width="120">Customername:</td>
              <td>
<input name="customer_name" type="text" id="customer_name" value="<%=request.querystring("customer_name")%>"></td>
          </tr>
          <tr> 
            <td>Companyname:</td>
              <td>
<input name="company_name" type="text" id="company_name" value="<%=request.querystring("company_name")%>"></td>
          </tr>
            <tr align="center"> 
              <td colspan="2"> 
                <input type="submit" name="Submit" value="Search">
              </td>
          </tr>
        </table>
		</form>
	  </div>
    </td>
	</tr>
	<form name="form1" method="post" action="">
    <%
	bgcolor = "#EEEEEE"
  	for x = 1 to 20
		if rsOrders.eof then exit for
		if bgcolor = "#EEEEEE" then
			bgcolor = "#FFFFFF"
		else
			bgcolor = "#EEEEEE"
		end if
		
		paid      = cint(rsOrders("paid"))
		if lcase(paid) = 0 then
			paid = "<img src=""../images/not_paid.gif"" width=""60"" height=""15"" align=""absmiddle"" alt=""Not paid"" />"
		else
			paid = "<img src=""../images/paid.gif"" width=""60"" height=""15"" align=""absmiddle"" alt=""Paid"" />"
		end if
    %>
    <tr bgcolor="<%=bgcolor%>"> 
      <td width="20" align="center"><input name="chkOrder_<%=x%>" type="checkbox" value="yes" /> 
        <input type="hidden" name="order_<%=x%>" value="<%=rsOrders("Order_id")%>" /></td>
      <td width="20"> 
        <% if intModuleRights = 2 then %>
	    <a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;oid=<%=rsOrders("Order_id")%>"> 
        <% else %>
		<a href="#"> 
        <% end if %>
        #<%=rsOrders("order_id")%></a> </td>
	  <td>
	  <%=rsOrders("user_lastname") & " " & rsOrders("user_firstname")%> [<%=rsOrders("date_ordered")%>]
	  </td>
	  <td><%=shop_currency & rsOrders("total_price")%></td>
	  <td width="60"><%=paid%></td>
    </tr>
    <%
  		rsOrders.movenext
	next
    %>
    <tr align="right" bgcolor="#999999"> 
      <td colspan="5" style="border: solid 1px #000000;"> <input type="hidden" name="total_records" value="<%=x-1%>" /> 
        <%
		set rsOrderType = server.createobject("ADODB.recordset")
		rsOrderType.cursortype = 3
		
	  	strSQL = "SELECT order_status_id, order_status FROM order_status WHERE lang_id = " & default_lang_id & " ORDER BY order_status_id ASC;"
	 	rsOrderType.open strSQL, adoCon
	  	%>
        With selected: 
        <select name="action">
          <option value="0">>>-->> select action</option>
          <option value="-1">delete</option>
		  <option value="-2">Change payment -> Not paid</option>
		  <option value="-3">Change payment -> Paid</option>
          <%
		  do while not rsOrderType.eof
		  	response.write("<option value=""" & rsOrderType("order_status_id") & """>Change status -> " & rsOrderType("order_status") & "</option>" & chr(10))
		  	rsOrderType.movenext
		  loop
		  
		  rsOrderType.close
		  set rsOrderType = nothing
		  %>
        </select> <%=BuildSubmitter("submit","OK", request.querystring("p"))%> </td>
    </tr>
 	</form>
  </table>
  <br />&nbsp;Pages: 
  <%
  ext = ""
  if stype = "oid" then
  	ext = "&amp;stype=oid&amp;oid=" & oid
  elseif stype = "adv" then
  	ext = "&amp;stype=adv&amp;customer_name=" & customer_n & "&amp;company_name=" & company_n
  end if
  for page_looper = 1 to pages
  	response.write("&nbsp;<a href=""?p=" & request.querystring("p") & "&amp;order_type=" & order_type & "&amp;pp=" & page_looper & ext & """>" & page_looper & "</a>")
  next
  %>
<script>
<% if stype = "oid" then %>
viewSearch('oid');
<% end if %>
<% if stype = "adv" then %>
viewSearch('adv');
<% end if %>
</script>
