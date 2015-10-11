<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
      <td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">News</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td>
	<%
	action = request.querystring("action")
	select case action
		case "add":
			%><!-- #include file="news/add.asp" --><%
		case "edit":
			%><!-- #include file="news/edit.asp" --><%
		case "delete":
			%><!-- #include file="news/delete.asp" --><%
		case else
			%><!-- #include file="news/home.asp" --><%
	end select
	%>
	</td>
  </tr>
</table>
</center>