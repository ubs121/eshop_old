<%
if session("customer_id") = "" OR len(session("customer_id")) = 0 then
	response.redirect("?mod=myaccount&sub=login&red=myaccount")
end if
%> 
<p><b><%=strMyAccount%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" align="center" class="content"><img src="images/myaccount.jpg" width="65" height="65" alt="<%=strMyAccount%>" /></td>
          <td class="content">
		    <p><a href="?mod=myaccount&amp;sub=details">&raquo;&nbsp;<%=strViewAccountInformation%></a></p>
			<p><a href="?mod=myaccount&amp;sub=addressbook">&raquo;&nbsp;<%=strViewAddressbook%></a></p>
			<p><a href="?mod=myaccount&amp;sub=password">&raquo;&nbsp;<%=strChangePassword%></a></p>
			<p><a href="?mod=myaccount&amp;sub=orders_history">&raquo;&nbsp;<%=strOrderHistory%></a></p>
		  </td>
        </tr>
      </table></td>
  </tr>
</table>
<% if strShowNewsletter = 1 then %>
<p><b><%=strMySubscriptions%></b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
  <tr> 
    <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" class="content" align="center"><img src="images/mysubscriptions.jpg" width="65" height="65" alt="<%=strMySubscriptions%>" /></td>
          <td class="content">
		    <p><a href="?mod=myaccount&amp;sub=newsletter">&raquo;&nbsp;<%=strNewsLetter%></a></p>
		  </td>
        </tr>
      </table></td>
  </tr>
</table>
<% end if %>