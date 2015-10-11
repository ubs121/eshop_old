<!-- #include file="common.asp" -->
<!-- #include file="fckEditor/fckeditor.asp" -->
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = strVirtualPath & "admin/FCKeditor/"
oFCKeditor.Height = "500"
%>
<% if CheckLogin() = 0 then response.redirect("login.asp?p=login") %>
<% intModuleRights = getRights(request.querystring("p")) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Admin control panel: Configuration</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="includes/style.css" rel="stylesheet" type="text/css">
</head>

<body>
	  <!-- #include file="content/config/index.asp" -->
</body>
</html>
