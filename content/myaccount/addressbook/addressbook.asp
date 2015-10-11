<%
intError = 0
strError = ""

set rsAddressbook = server.createobject("ADODB.recordset")
rsAddressbook.cursortype = 3

strSQL = "SELECT * FROM user_address WHERE user_id = " & session("customer_id")
rsAddressbook.open strSQL, adoCon

'Primary address
rsAddressbook.filter = "user_default_address = -1"
%>
<p><b><%=strPrimaryAddress%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="left" valign="top" class="content"><%=strPrimaryExplanation%></td>
          <td width="120" align="center" class="content"><b><%=strPrimaryAddress%></b><br /><img src="images/pinned.gif" alt="<%=strPrimaryAddress%>" width="40" height="20" align="middle" /></td>
          <td width="200" class="content">
			<%=rsAddressbook("user_firstname") & " " & rsAddressbook("user_lastname")%><br />
			<%=rsAddressbook("user_street")%><br />
			<%=rsAddressbook("user_postcode") & " " & rsAddressbook("user_city")%><br />
			<%=rsAddressbook("user_province") & ", " & rsAddressbook("user_country")%>
		  </td>
        </tr>
      </table></td>
    </tr>
  </table>
<p><b><%=strAddressbookEntries%></b></p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top: solid 1px #C2C2C2">
    <tr> 
      <td class="table_content"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <!-- default address -->
        <tr> 
          <td width="20" class="content">&nbsp;</td>
          <td class="content">
		    &nbsp;<b><%=rsAddressbook("user_lastname") & " " & rsAddressbook("user_firstname")%></b>&nbsp;<i>(<%=strPrimaryAddress%>)</i><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_street")%><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_postcode") & " " & rsAddressbook("user_city")%><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_province") & ", " & rsAddressbook("user_country")%>
		  </td>
          <td width="200" class="content">&nbsp;<a href="?mod=myaccount&amp;sub=addressbook&amp;action=edit&amp;id=<%=rsAddressbook("user_address_id")%>"><%=strEdit%></a></td>
        </tr>
	  <!-- other addresses -->
	  <%
	  rsAddressbook.filter = "user_default_address = 0"
	  do while not rsAddressbook.eof
	  %>
        <tr> 
          <td width="20" class="content">&nbsp;</td>
          <td class="content">
		    &nbsp;<b><%=rsAddressbook("user_lastname") & " " & rsAddressbook("user_firstname")%></b><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_street")%><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_postcode") & " " & rsAddressbook("user_city")%><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsAddressbook("user_province") & ", " & rsAddressbook("user_country")%>
		  </td>
          <td width="200" class="content">&nbsp;<a href="?mod=myaccount&amp;sub=addressbook&amp;action=edit&amp;id=<%=rsAddressbook("user_address_id")%>"><%=strEdit%></a> | <a href="?mod=myaccount&amp;sub=addressbook&amp;action=delete&amp;id=<%=rsAddressbook("user_address_id")%>"><%=strDelete%></a></td>
        </tr>
	  <%
	  	rsAddressbook.movenext
	  loop
	  %>
      </table></td>
    </tr>
  </table>
<%
rsAddressbook.close
set rsAddressbook = nothing
%>
<br />
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="50%"><a href="?mod=myaccount"><img src="languages/<%=session("language")%>/images/button_back.gif" alt="<%=strBack%>" width="122" height="22" border="0" /></a></td>
    <td width="50%" align="right"><a href="?mod=myaccount&amp;sub=addressbook&amp;action=add"><img src="languages/<%=session("language")%>/images/button_add_address.gif" alt="<%=strBack%>" width="122" height="22" border="0" /></a></td>
  </tr>
</table>