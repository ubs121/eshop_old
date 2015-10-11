<!-- #include file="common.asp" -->
<% intModuleRights = getRights(request.querystring("p")) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Admin control panel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="includes/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<% if CheckLogin() = 1 then %>
<%
	if len(p) = 0 then
		p = "cp"
	end if

end if
%>
<br><br>
	  <!-- #include file="content/login/index.asp" -->
</body>
</html>
