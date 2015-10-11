<!-- #include file="header.inc.asp" -->
<html>
<head>
  <title><%=strPageTitle%></title>
  <meta http-equiv="Content-Type" content="text/html; charset=<%=strCharset%>">
  <link href="includes/styles/<%=strStyleSheet%>.css" rel="stylesheet" type="text/css">
  <script type="text/javascript" language="JavaScript">
  <!--
  var strQueryStrings = "[<%=request.querystring()%>]";
  -->
  </script>
  <script language="JavaScript" type="text/javascript" src="includes/js/common.js"></script>
</head>
<body dir="<%=strTextDirectory%>">
<table cellspacing="0" cellpadding="0" class="maintable" width="100%">
  <tr>  
    <td colspan="3">
	  <!-- #include file="banner.inc.asp" -->
	</td>
  </tr>
  <tr> 
    <td style="width: 150px; padding: 5px 5px 5px 5px;">
	  <% if module <> "setlanguage" then %>
	  <!-- #include file="leftmenu.inc.asp" -->
	  <% end if %>
	</td>
    <td><!-- #include file="content.asp" --></td>
    <td style="width: 150px; padding: 5px;">
	  <!-- #include file="rightmenu.inc.asp" -->
	</td>
  </tr>
</table>
<div class="copyright">
	<%
		call WriteCopy()
	%>
</div>
<div id="tipDiv" style="position:absolute; visibility:hidden; z-index:100"></div>
</body>
</html>
<!-- #include file="footer.inc.asp" -->
