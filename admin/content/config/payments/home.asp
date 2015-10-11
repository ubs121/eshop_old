<%
set rsPayment = server.createobject("ADODB.recordset")
rsPayment.cursortype = 3

strSQL = "SELECT payment_id, payment_name FROM payment WHERE payment_lang_id = " & default_lang_id
rsPayment.open strSQL, adoCon
%>
<table width="500" border="0" align="center" cellpadding="1" cellspacing="0">
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<%
	rsPayment.filter = "payment_ID = 1"
	if not rsPayment.eof then
		installed = true
		bg_color = "#339900"
	else
		installed = false
		bg_color = "#CC3300"
	end if
	%>
	<tr bgcolor="<%=bg_color%>">
		<td>&nbsp;<b>Cash</b></td>
		<td width="250" align="right">
			<% if installed then %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;opt=edit&amp;PID=1"><img src="images/button_edit.gif" width="71" height="21" alt="edit" /></a>
			<a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;opt=uninstall&amp;PID=1"><img src="images/button_uninstall.gif" width="71" height="21" alt="uninstall" /></a>
			<% end if %>
			<% else %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;opt=edit&amp;PID=1"><img src="images/button_install.gif" width="71" height="21" alt="install" /></a>
			<% end if %>
			<% end if %>
			&nbsp;
		</td>
	</tr>
	<%
	rsPayment.filter = "payment_ID = 2"
	if not rsPayment.eof then
		installed = true
		bg_color = "#339900"
	else
		installed = false
		bg_color = "#CC3300"
	end if
	%>
	<tr bgcolor="<%=bg_color%>">
		
    <td>&nbsp;<b>Paypal</b></td>
		<td width="250" align="right">
			<% if installed then %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=paypal&amp;opt=edit"><img src="images/button_edit.gif" width="71" height="21" alt="edit" /></a>
			<a href="?p=<%=request.querystring("p")%>&amp;action=paypal&amp;opt=uninstall"><img src="images/button_uninstall.gif" width="71" height="21" alt="uninstall" /></a>
			<% end if %>
			<% else %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=paypal&amp;opt=edit"><img src="images/button_install.gif" width="71" height="21" alt="install" /></a>
			<% end if %>
			<% end if %>
			&nbsp;
		</td>
	</tr>
	<%
	rsPayment.filter = "payment_ID = 3"
	if not rsPayment.eof then
		installed = true
		bg_color = "#339900"
	else
		installed = false
		bg_color = "#CC3300"
	end if
	%>
	<tr bgcolor="<%=bg_color%>">
		
    <td>&nbsp;<b>Ogone</b></td>
		<td width="250" align="right">
			<% if installed then %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=ogone&amp;opt=edit"><img src="images/button_edit.gif" width="71" height="21" alt="edit" /></a>
			<a href="?p=<%=request.querystring("p")%>&amp;action=ogone&amp;opt=uninstall"><img src="images/button_uninstall.gif" width="71" height="21" alt="uninstall" /></a>
			<% end if %>
			<% else %>
			<% if intModuleRights = 2 then %>
			<a href="?p=<%=request.querystring("p")%>&amp;action=ogone&amp;opt=edit"><img src="images/button_install.gif" width="71" height="21" alt="install" /></a>
			<% end if %>
			<% end if %>
			&nbsp;
		</td>
	</tr>
	<tr>
	  <td colspan="2" style="background: #000000; color: #FFFFFF; padding: 4; border-top: solid 1px #000000; border-bottom: solid 1px #000000; font-weight: bold; text-align: center;">Custom payments</td>
	</tr>
	<%
	hPid = 50
	rsPayment.filter = "payment_ID >= 50"
	do while not rsPayment.eof
		hPid = cint(rsPayment("payment_ID")) + 1
		response.write "<tr bgcolor=""#339900"">" & chr(10) & _
			"  <td>&nbsp;<b>" & rsPayment("payment_name") & "</b></td>" & chr(10) & _
			"  <td width=""250"" align=""right"">" & chr(10) & _
			"    <a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;opt=edit&amp;PID=" & rsPayment("payment_ID") & """>" & _
			"<img src=""images/button_edit.gif"" width=""71"" height=""21"" alt=""edit"" /></a>" & chr(10) & _
			"    <a href=""?p=" & request.querystring("p") & "&amp;action=edit&amp;opt=uninstall&amp;PID=" & rsPayment("payment_ID") & """>" & _
			"<img src=""images/button_uninstall.gif"" width=""71"" height=""21"" alt=""uninstall"" /></a>" & chr(10) & _
			"  </td>" & chr(10) & _
			"</tr>"
		rsPayment.movenext
	loop
	%>
	<tr>
	  <td align="center" colspan="2">
	  	<form name="frmAdd" action="?p=<%=request.querystring("p")%>&action=add" method="post">
		  <br />
		  <input type="hidden" name="pid" value="<%=hpid%>" />
	  	  <%=buildSubmitter("cmdAdd", "Add custom paymentmethod", request.querystring("p"))%>
		</form>
	</tr>
</table>
<br />
<%
rsPayment.close
set rsPayment = nothing
%>	