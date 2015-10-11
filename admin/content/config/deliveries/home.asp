<%
set rsDeliveries = server.createobject("ADODB.recordset")
rsDeliveries.cursortype = 3

strSQL = "SELECT delivery_id, delivery_name, a FROM delivery WHERE lang_id = " & default_lang_id & " ORDER BY delivery_name ASC;"
rsDeliveries.open strSQL, adoCon
%>
<script>
function deleteMe(id){
	doIt = confirm("Are you sure you want to delete this deliverymethod?");
	if(doIt)document.location = "?p=<%=request.querystring("p")%>&action=delete&did=" + id;
}
</script>
<table width="500" align="center" cellpadding="2" cellspacing="2" style="border: solid 1px #000000;">
  <tr bgcolor="#666666"> 
    <td><font color="#FFFFFF"><strong>&nbsp;Name</strong></font></td>
    <td width="150"><font color="#FFFFFF"><strong>&nbsp;Actions</strong></font></td>
  </tr>
	<%
	do while not rsDeliveries.eof
		if rsDeliveries("a") = "1" then
			opt = "f"
		else
			opt = "v"
		end if
	%>
  <tr> 
    <td>&nbsp;<%=rsDeliveries("delivery_name")%></td>
    <td>
	  &nbsp;
	  <% if intModuleRights = 2 then %>
	  <a href="?p=<%=request.querystring("p")%>&amp;action=edit&amp;did=<%=rsDeliveries("delivery_ID")%>&amp;opt=<%=opt%>">Edit</a> | 
	  <a href="javascript:deleteMe('<%=rsDeliveries("delivery_ID")%>');">Delete</a>
	  <% end if %>
	</td>
  </tr>
	<%
		rsDeliveries.movenext
	loop  
	%>
</table>
<% if intModuleRights = 2 then %>
<form id="frmAdd" name="frmAdd" method="post" action="?p=<%=request.querystring("p")%>&amp;action=add">
  <p align="center">
    <select name="slType" id="slType">
      <option value="f">Fixed price</option>
      <option value="v">Variable price</option>
    </select>
    <%=buildSubmitter("cmdAdd", "Add deliverymethod", request.querystring("p"))%>  </p>
</form>
<p align="center"><a href="?p=<%=request.querystring("p")%>&amp;action=add"></a></p>
<% end if %>