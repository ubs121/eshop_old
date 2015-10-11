<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="banner"><img src="includes/styles/<%=strStyleSheet%>/banner.jpg" /></td>
  </tr>
  <tr>
    <td class="header_menu">
	  <ul class="menulist" id="menulist">
	    <li><a href="?mod=myaccount"><% if session("customer_id") = "" OR len(session("customer_id")) = 0 then%><%=strLogin%><% else %><%=strMyAccount%><% end if %></a></li>
		<li><a href="?mod=cart&action=view"><%=strShoppingCart%></a></li>
		<li><a href="?mod=checkout"><%=strCheckOut%></a> <% if session("customer_id") <> "" AND len(session("customer_id")) > 0 then%></li>
		<li><a href="?mod=myaccount&amp;sub=logout"><%=strLogout%></a><% end if %></li>
	  </ul>
	 </td>
  </tr>
</table>