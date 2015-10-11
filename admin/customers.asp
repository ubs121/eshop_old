<!-- #include file="common.asp" -->
<!-- #include file="fckEditor/fckeditor.asp" -->
<% if CheckLogin() = 0 then response.redirect("login.asp?p=login") %>
<% intModuleRights = getRights(request.querystring("p")) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Customers</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="includes/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<!-- #include file="content/customers/index.asp" -->
<script src="includes/tooltip.js"></script>
</body>
</html>
