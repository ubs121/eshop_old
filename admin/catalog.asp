<!-- #include file="common.asp" -->
<!-- #include file="includes/functions/sendmail.asp" -->
<!-- #include file="fckEditor/fckeditor.asp" -->
<% if CheckLogin() = 0 then response.redirect("login.asp?p=login") %>
<% intModuleRights = getRights(request.querystring("p")) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Catalog</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="includes/style.css" rel="stylesheet" type="text/css">
<script>
<!--
function doUpload(upload_type){
	uploader = window.open("upload.asp?first=1&type=" + upload_type,"upload","toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=1,width=400,height=30,left = 0,top = 0");
}

function showFilebrowser(folder){
	filebrowser = window.open("filebrowser.asp?folder=" + folder, "filebrowser", "toolbar=0, scrollbars=0, location=0, statusbar=0, menubar=0, resizable=0, width=580, height=350,left=0, top=0");
}
-->
</script>
</head>
<body>
<!-- #include file="content/catalog/index.asp" -->
</body>
</html>